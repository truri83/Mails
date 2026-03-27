Attribute VB_Name = "modMailExtract"
Option Compare Database
Option Explicit

' ===========================================================================
' modMailExtract - Datenextraktion aus Outlook COM-Objekten
' ===========================================================================
' Zentrales Modul fuer das "Extract-Release-Process" Pattern:
'   1. EXTRACT: Alle benoetigten Daten aus einem RDO-Objekt lesen
'   2. RELEASE: COM-Objekt sofort freigeben (Outlook entlasten)
'   3. PROCESS: Hash, DB, Dateien - alles OHNE COM-Referenz
'
' Definiert alle projektweiten Datentypen (Public Types):
'   - TypMailDaten         (Mail-Metadaten + Content)
'   - TypEmpfaengerDaten   (Empfaenger einer Mail)
'   - TypAnhangDaten       (Anhang-Metadaten + Temp-Pfad)
'   - TypMailKomplett      (Container: Mail + Empfaenger + Anhaenge)
'   - TypDateiOperation    (Queue-Eintrag: Temp -> Netzwerk)
'
' Abhaengigkeiten: modOutlookConnect (SMTP-Aufloesung),
'                  modGlobals (Konstanten, MAPI-Tags),
'                  modCrypto (Hash-Generierung),
'                  modStringUtils (Bereinigung, Pfade)
' ===========================================================================


' ---------------------------------------------------------------------------
' DATENTYPEN
' ---------------------------------------------------------------------------

' Mail-Metadaten (alles, was aus dem COM-Objekt gelesen wird)
Public Type TypMailDaten
    EntryID             As String
    Betreff             As String
    AbsenderName        As String
    AbsenderEmail       As String
    AbsenderEmailTyp    As String
    EmpfangenAm         As Date
    GesendetAm          As Date
    Groesse             As Long
    Wichtigkeit         As Integer
    Gelesen             As Boolean
    HatAnhaenge         As Boolean
    AnhangAnzahl        As Integer
    MessageClass        As String
    InternetMessageID   As String
    InReplyTo           As String
    DisplayTo           As String
    HTMLBody            As String
    PlainTextBody       As String
    IstGueltig          As Boolean
End Type

' Empfaenger einer Mail
Public Type TypEmpfaengerDaten
    Name    As String
    Email   As String
    Typ     As String   ' "To", "CC", "BCC"
End Type

' Anhang-Metadaten + Temp-Pfad
Public Type TypAnhangDaten
    Dateiname       As String
    Groesse         As Long
    MimeType        As String
    AnhangTyp       As Integer   ' 1=Datei, 5=OLE, 6=Mail
    IstVersteckt    As Boolean
    TempPfad        As String    ' Lokaler Temp-Pfad (nach Extraktion)
End Type

' Komplett-Paket: Mail + Empfaenger + Anhaenge + Temp-Dateien
Public Type TypMailKomplett
    Mail                As TypMailDaten
    Empfaenger()        As TypEmpfaengerDaten
    EmpfaengerAnzahl    As Integer
    Anhaenge()          As TypAnhangDaten
    AnhangAnzahl        As Integer
    MSGTempPfad         As String   ' Temp-Pfad fuer MSG-Export
    UniqueHash          As String   ' SHA256 Duplikat-Hash
End Type

' Datei-Operation fuer die Queue (Temp -> Netzwerk)
Public Type TypDateiOperation
    QuellPfad       As String   ' Lokaler Temp-Pfad
    ZielPfad        As String   ' Netzwerk-Zielpfad
    OperationsTyp   As String   ' "MSG" oder "Anhang"
    EmailID         As Long     ' FK fuer DB-Update nach Kopie
    AnhangID        As Long     ' FK fuer DB-Update nach Kopie
    Versuche        As Integer  ' Bisherige Versuche
End Type


' ---------------------------------------------------------------------------
' TEMP-ZAEHLER (modulweit, fuer eindeutige Temp-Dateinamen)
' ---------------------------------------------------------------------------
Private m_lngTempZaehler As Long


' ===========================================================================
' KOMPLETT-EXTRAKTION (Hauptfunktion)
' ===========================================================================

' Liest ALLE benoetigten Daten aus einem RDO-Mail-Objekt in eine
' TypMailKomplett-Struktur. Nach diesem Aufruf kann das COM-Objekt
' sofort freigegeben werden.
'
' Parameter:
'   objRDOMail      - Redemption RDOMail-Objekt
'   mk              - ByRef: Wird mit allen extrahierten Daten gefuellt
'   blnAnhangZuTemp - Anhaenge in Temp-Ordner extrahieren?
'   blnMSGZuTemp    - MSG-Datei in Temp-Ordner speichern?
'   blnAnhangFilter - Signatur-Bilder filtern? (Hidden/Type<>1)
Public Sub ExtrahiereKomplett(objRDOMail As Object, _
                               ByRef mk As TypMailKomplett, _
                               Optional ByVal blnAnhangZuTemp As Boolean = True, _
                               Optional ByVal blnMSGZuTemp As Boolean = True, _
                               Optional ByVal blnAnhangFilter As Boolean = True)
    On Error GoTo ErrHandler

    ' 1. Skalare Eigenschaften + Content
    mk.Mail = ExtrahiereMailDaten(objRDOMail)

    ' 2. Hash berechnen (benoetigt extrahierte Daten)
    mk.UniqueHash = GeneriereMailHash(mk.Mail.Betreff, mk.Mail.AbsenderEmail, _
                                      mk.Mail.DisplayTo, mk.Mail.EmpfangenAm)

    ' 3. Empfaenger
    Call ExtrahiereEmpfaengerIn(mk, objRDOMail)

    ' 4. Anhaenge (optional mit Temp-Speicherung + Filter)
    If mk.Mail.AnhangAnzahl > 0 Then
        Call ExtrahiereAnhaengeIn(mk, objRDOMail, blnAnhangZuTemp, blnAnhangFilter)
    End If

    ' 5. MSG in Temp speichern
    If blnMSGZuTemp Then
        mk.MSGTempPfad = SpeichereMSGZuTemp(objRDOMail)
    End If

    mk.Mail.IstGueltig = True
    Exit Sub

ErrHandler:
    mk.Mail.IstGueltig = False
    LogVBAError "ExtrahiereKomplett [" & Left(Nz(mk.Mail.Betreff, "?"), 30) & "]"
End Sub


' ===========================================================================
' BASISDATEN EXTRAHIEREN (Skalare Felder + Content)
' ===========================================================================

Private Function ExtrahiereMailDaten(objRDOMail As Object) As TypMailDaten
    Dim md As TypMailDaten
    On Error Resume Next

    ' --- Identifikation ---
    md.EntryID = Nz(objRDOMail.EntryID, "")
    If Err.Number <> 0 Then Err.Clear

    ' --- Betreff ---
    md.Betreff = Nz(objRDOMail.Subject, "")
    If Err.Number <> 0 Then Err.Clear

    ' --- Absender ---
    md.AbsenderName = Nz(objRDOMail.SenderName, "")
    If Err.Number <> 0 Then Err.Clear

    md.AbsenderEmailTyp = Nz(objRDOMail.SenderEmailType, "SMTP")
    If Err.Number <> 0 Then Err.Clear: md.AbsenderEmailTyp = "SMTP"

    ' SMTP-Adresse aufloesen (via modOutlookConnect)
    On Error GoTo 0
    On Error Resume Next
    md.AbsenderEmail = GetAbsenderSMTP(objRDOMail)
    If Err.Number <> 0 Then Err.Clear: md.AbsenderEmail = DEFAULT_EMAIL

    ' --- Zeiten ---
    md.EmpfangenAm = Nz(objRDOMail.ReceivedTime, Now)
    If Err.Number <> 0 Then Err.Clear: md.EmpfangenAm = Now

    md.GesendetAm = Nz(objRDOMail.SentOn, md.EmpfangenAm)
    If Err.Number <> 0 Then Err.Clear: md.GesendetAm = md.EmpfangenAm

    ' --- Groesse + Status ---
    md.Groesse = objRDOMail.Size
    If Err.Number <> 0 Then Err.Clear: md.Groesse = 0

    md.Wichtigkeit = objRDOMail.Importance
    If Err.Number <> 0 Then Err.Clear: md.Wichtigkeit = 1

    md.Gelesen = Not objRDOMail.UnRead
    If Err.Number <> 0 Then Err.Clear: md.Gelesen = False

    ' --- Anhaenge ---
    md.AnhangAnzahl = objRDOMail.Attachments.Count
    If Err.Number <> 0 Then Err.Clear: md.AnhangAnzahl = 0
    md.HatAnhaenge = (md.AnhangAnzahl > 0)

    ' --- MessageClass ---
    md.MessageClass = Nz(objRDOMail.MessageClass, "IPM.Note")
    If Err.Number <> 0 Then Err.Clear: md.MessageClass = "IPM.Note"

    ' --- MAPI Properties ---
    md.InternetMessageID = objRDOMail.Fields(PR_INTERNET_MESSAGE_ID)
    If Err.Number <> 0 Then Err.Clear: md.InternetMessageID = ""

    md.DisplayTo = objRDOMail.Fields(PR_DISPLAY_TO)
    If Err.Number <> 0 Then Err.Clear: md.DisplayTo = ""

    ' --- Transport-Headers -> In-Reply-To ---
    Dim strHeaders As String
    strHeaders = objRDOMail.Fields(PR_TRANSPORT_MESSAGE_HEADERS)
    If Err.Number <> 0 Then Err.Clear: strHeaders = ""
    md.InReplyTo = ParseHeaderField(strHeaders, "In-Reply-To")

    ' --- Content (HTML + Plain) ---
    md.HTMLBody = Nz(objRDOMail.HTMLBody, "")
    If Err.Number <> 0 Then Err.Clear: md.HTMLBody = ""

    md.PlainTextBody = Nz(objRDOMail.Body, "")
    If Err.Number <> 0 Then Err.Clear: md.PlainTextBody = ""

    On Error GoTo 0
    md.IstGueltig = True
    ExtrahiereMailDaten = md
End Function


' ===========================================================================
' EMPFAENGER EXTRAHIEREN
' ===========================================================================

Private Sub ExtrahiereEmpfaengerIn(ByRef mk As TypMailKomplett, objRDOMail As Object)
    On Error Resume Next

    Dim intCount As Integer
    intCount = objRDOMail.Recipients.Count
    If Err.Number <> 0 Then Err.Clear: intCount = 0

    mk.EmpfaengerAnzahl = intCount
    If intCount = 0 Then Exit Sub

    ReDim mk.Empfaenger(0 To intCount - 1)

    Dim objRec As Object
    Dim i As Integer: i = 0

    For Each objRec In objRDOMail.Recipients
        If i > intCount - 1 Then Exit For

        mk.Empfaenger(i).Name = Nz(objRec.Name, "")
        If Err.Number <> 0 Then Err.Clear: mk.Empfaenger(i).Name = ""

        mk.Empfaenger(i).Email = GetSMTPFromRecipient(objRec)
        If Err.Number <> 0 Then Err.Clear: mk.Empfaenger(i).Email = ""

        Select Case objRec.Type
            Case 1: mk.Empfaenger(i).Typ = "To"
            Case 2: mk.Empfaenger(i).Typ = "CC"
            Case 3: mk.Empfaenger(i).Typ = "BCC"
            Case Else: mk.Empfaenger(i).Typ = "?"
        End Select
        If Err.Number <> 0 Then Err.Clear

        i = i + 1
    Next objRec

    ' Tatsaechliche Anzahl (falls weniger gelesen)
    mk.EmpfaengerAnzahl = i
    On Error GoTo 0
End Sub


' ===========================================================================
' ANHAENGE EXTRAHIEREN
' ===========================================================================

Private Sub ExtrahiereAnhaengeIn(ByRef mk As TypMailKomplett, _
                                  objRDOMail As Object, _
                                  ByVal blnZuTemp As Boolean, _
                                  ByVal blnFilter As Boolean)
    On Error Resume Next

    Dim intCount As Integer
    intCount = objRDOMail.Attachments.Count
    If Err.Number <> 0 Then Err.Clear: intCount = 0

    mk.AnhangAnzahl = intCount
    If intCount = 0 Then Exit Sub

    ReDim mk.Anhaenge(0 To intCount - 1)

    Dim objAtt As Object
    Dim a As Integer
    Dim blnSave As Boolean

    For a = 1 To intCount
        Set objAtt = objRDOMail.Attachments(a)
        If Err.Number <> 0 Then Err.Clear: GoTo NaechsterAnhang

        mk.Anhaenge(a - 1).Dateiname = Nz(objAtt.FileName, "")
        If Err.Number <> 0 Then Err.Clear

        mk.Anhaenge(a - 1).Groesse = objAtt.Size
        If Err.Number <> 0 Then Err.Clear: mk.Anhaenge(a - 1).Groesse = 0

        mk.Anhaenge(a - 1).MimeType = Nz(objAtt.MimeTag, "")
        If Err.Number <> 0 Then Err.Clear

        mk.Anhaenge(a - 1).AnhangTyp = objAtt.Type
        If Err.Number <> 0 Then Err.Clear: mk.Anhaenge(a - 1).AnhangTyp = ATT_BY_VALUE

        mk.Anhaenge(a - 1).IstVersteckt = objAtt.Hidden
        If Err.Number <> 0 Then Err.Clear: mk.Anhaenge(a - 1).IstVersteckt = False

        ' Entscheidung: Anhang in Temp speichern?
        blnSave = blnZuTemp
        If blnFilter Then
            ' Mit Filter: nur echte Dateien (Type=1), nicht versteckt
            blnSave = blnSave And (mk.Anhaenge(a - 1).AnhangTyp = ATT_BY_VALUE) _
                      And Not mk.Anhaenge(a - 1).IstVersteckt
        End If

        If blnSave Then
            mk.Anhaenge(a - 1).TempPfad = SpeichereAnhangZuTemp(objAtt)
        End If

NaechsterAnhang:
    Next a
    On Error GoTo 0
End Sub


' ===========================================================================
' TEMP-DATEI-OPERATIONEN
' ===========================================================================

' Speichert einen einzelnen Anhang in den Temp-Ordner
Private Function SpeichereAnhangZuTemp(objAtt As Object) As String
    On Error GoTo ErrHandler

    Dim strTempDir As String
    strTempDir = HoleTempPfad()
    ErstelleOrdner strTempDir

    m_lngTempZaehler = m_lngTempZaehler + 1

    Dim strExt As String
    strExt = HoleEndung(Nz(objAtt.FileName, ""))
    If strExt = "" Then strExt = "bin"

    Dim strPfad As String
    strPfad = strTempDir & "att_" & Format(Now, "yyyymmdd_hhnnss") & _
              "_" & Format(m_lngTempZaehler, "0000") & "." & strExt

    objAtt.SaveAsFile strPfad
    SpeichereAnhangZuTemp = strPfad
    LogTrace "Anhang-Temp: " & objAtt.FileName & " -> " & strPfad, "EXTRACT"
    Exit Function

ErrHandler:
    LogWarn "Anhang-Temp fehlgeschlagen: " & Nz(objAtt.FileName, "?") & _
            " - " & Err.Description, "EXTRACT"
    SpeichereAnhangZuTemp = ""
End Function


' Speichert eine MSG-Datei in den Temp-Ordner
Private Function SpeichereMSGZuTemp(objRDOMail As Object) As String
    On Error GoTo ErrHandler

    Dim strTempDir As String
    strTempDir = HoleTempPfad()
    ErstelleOrdner strTempDir

    m_lngTempZaehler = m_lngTempZaehler + 1

    Dim strPfad As String
    strPfad = strTempDir & "msg_" & Format(Now, "yyyymmdd_hhnnss") & _
              "_" & Format(m_lngTempZaehler, "0000") & ".msg"

    objRDOMail.SaveAs strPfad
    SpeichereMSGZuTemp = strPfad
    LogTrace "MSG-Temp: " & strPfad, "EXTRACT"
    Exit Function

ErrHandler:
    LogWarn "MSG-Temp fehlgeschlagen: " & Err.Description, "EXTRACT"
    SpeichereMSGZuTemp = ""
End Function


' ===========================================================================
' HILFSFUNKTIONEN
' ===========================================================================

' Gibt den Temp-Verzeichnispfad zurueck (mit abschliessendem \)
Public Function HoleTempPfad() As String
    Dim strTemp As String
    strTemp = LeseConfig("TempPfad", "")

    If strTemp = "" Then
        strTemp = Environ("TEMP")
        If Right(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
        strTemp = strTemp & "OutlookSync\"
    Else
        If Right(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
    End If

    HoleTempPfad = strTemp
End Function


' Temp-Dateien im OutlookSync-Temp-Ordner aufraeumen
Public Sub BereinigeTempDateien()
    On Error Resume Next
    Dim strTempDir As String
    strTempDir = HoleTempPfad()

    If Dir(strTempDir, vbDirectory) = "" Then Exit Sub

    Dim strDatei As String
    strDatei = Dir(strTempDir & "*.*")
    Do While strDatei <> ""
        Kill strTempDir & strDatei
        strDatei = Dir()
    Loop

    On Error GoTo 0
    LogDebug "Temp-Dateien bereinigt: " & strTempDir, "EXTRACT"
End Sub


' Setzt Temp-Zaehler zurueck (am Anfang eines Sync-Laufs)
Public Sub ResetTempZaehler()
    m_lngTempZaehler = 0
End Sub


' ===========================================================================
' HEADER-PARSER (ehemals in modSync)
' ===========================================================================

' Extrahiert ein Feld aus Internet-Transport-Headers
' Format: "FieldName: value\r\n"
Public Function ParseHeaderField(ByVal strHeaders As String, _
                                  ByVal strFieldName As String) As String
    On Error Resume Next
    Dim pos As Long, posEnd As Long
    Dim strSearch As String

    ParseHeaderField = ""
    If Len(strHeaders) = 0 Then Exit Function

    ' Im Header suchen (Format: "FieldName: value\r\n")
    strSearch = vbCrLf & strFieldName & ": "
    pos = InStr(1, strHeaders, strSearch, vbTextCompare)

    ' Auch am Anfang der Headers pruefen
    If pos = 0 Then
        strSearch = strFieldName & ": "
        If LCase(Left(strHeaders, Len(strSearch))) = LCase(strSearch) Then
            pos = 1
        Else
            Exit Function
        End If
    Else
        pos = pos + 2  ' vbCrLf ueberspringen
    End If

    pos = pos + Len(strFieldName) + 2  ' "FieldName: " ueberspringen
    posEnd = InStr(pos, strHeaders, vbCrLf)
    If posEnd = 0 Then posEnd = Len(strHeaders) + 1

    ParseHeaderField = Trim(Mid(strHeaders, pos, posEnd - pos))
End Function
