Attribute VB_Name = "modAsyncBuffer"
Option Compare Database
Option Explicit

' ===========================================================================
' modAsyncBuffer - Schreib-Puffer und Datei-Queue
' ===========================================================================
' Sammelt extrahierte Mail-Daten im Speicher und schreibt sie in
' Transaktions-Batches in die Datenbank. Datei-Operationen (MSG/Anhaenge
' von Temp auf Netzlaufwerk kopieren) werden separat in einer Queue
' verarbeitet mit Retry-Logik bei Netzwerkfehlern.
'
' MUSTER:
'   BufferInit
'   BufferSetzeKontext lngSyncLaufID, lngOrdnerID, ...
'   For Each Mail:
'     ExtrahiereKomplett objMail, mk
'     BufferHinzufuegen mk
'     If BufferIstVoll Then BufferFlush
'   Next
'   BufferFlush            ' Restliche Daten schreiben
'   DateiQueueVerarbeiten  ' Dateien kopieren (Temp -> Netzwerk)
'   BufferLeeren
'
' Performance-Vorteile:
'   - DB-Writes in Transaktionen (25x weniger Netzwerk-Roundtrips)
'   - Batch-Duplikatpruefung per SQL IN-Clause
'   - Datei-Kopie mit exponentiellem Retry-Backoff
'   - DoEvents zwischen Batches (UI bleibt responsiv)
'
' Abhaengigkeiten: modMailExtract (Types), modDAO, modKontakte,
'                  modCrypto, modStringUtils, modLogging, modGlobals
' ===========================================================================


' ---------------------------------------------------------------------------
' MODUL-VARIABLEN: Schreib-Puffer
' ---------------------------------------------------------------------------
Private m_aBuffer()         As TypMailKomplett
Private m_lngBufferAnzahl   As Long
Private m_lngBufferMax      As Long
Private m_blnInitialisiert  As Boolean

' Kontext (wird einmal pro SyncFolder gesetzt)
Private m_lngKontextSyncLaufID  As Long
Private m_lngKontextOrdnerID    As Long
Private m_strKontextProjekt     As String
Private m_strKontextPhase       As String
Private m_strKontextExportBase  As String
Private m_blnKontextAnhaenge    As Boolean
Private m_blnKontextMSG         As Boolean
Private m_blnKontextFilter      As Boolean

' ---------------------------------------------------------------------------
' MODUL-VARIABLEN: Datei-Queue
' ---------------------------------------------------------------------------
Private m_aDateiQueue()         As TypDateiOperation
Private m_lngDateiQueueAnzahl   As Long

' ---------------------------------------------------------------------------
' STATISTIK (kumuliert ueber alle Flushes eines Sync-Laufs)
' ---------------------------------------------------------------------------
Private m_lngGesamtNeu          As Long
Private m_lngGesamtDuplikate    As Long
Private m_lngGesamtFehler       As Long
Private m_lngGesamtVerarbeitet  As Long


' ===========================================================================
' INITIALISIERUNG
' ===========================================================================

' Puffer initialisieren (am Anfang jedes SyncFolder-Aufrufs)
' Standard-Groesse 25 Items - konfigurierbar in tblConfig.BufferGroesse
Public Sub BufferInit(Optional ByVal lngMaxGroesse As Long = 0)
    If lngMaxGroesse <= 0 Then
        lngMaxGroesse = CLng(LeseConfig("BufferGroesse", "25"))
    End If
    If lngMaxGroesse < 5 Then lngMaxGroesse = 5
    If lngMaxGroesse > 500 Then lngMaxGroesse = 500

    m_lngBufferMax = lngMaxGroesse
    m_lngBufferAnzahl = 0
    ReDim m_aBuffer(0 To m_lngBufferMax - 1)

    m_lngDateiQueueAnzahl = 0
    ReDim m_aDateiQueue(0 To 99)

    m_lngGesamtNeu = 0
    m_lngGesamtDuplikate = 0
    m_lngGesamtFehler = 0
    m_lngGesamtVerarbeitet = 0

    m_blnInitialisiert = True
    LogDebug "Buffer initialisiert: Max=" & m_lngBufferMax, "BUFFER"
End Sub


' Sync-Kontext setzen (Infos die fuer alle Mails im Batch gelten)
Public Sub BufferSetzeKontext(ByVal lngSyncLaufID As Long, _
                               ByVal lngOrdnerID As Long, _
                               ByVal strProjekt As String, _
                               ByVal strPhase As String, _
                               ByVal strExportBase As String, _
                               ByVal blnAnhaenge As Boolean, _
                               ByVal blnMSG As Boolean, _
                               ByVal blnFilter As Boolean)
    m_lngKontextSyncLaufID = lngSyncLaufID
    m_lngKontextOrdnerID = lngOrdnerID
    m_strKontextProjekt = strProjekt
    m_strKontextPhase = strPhase
    m_strKontextExportBase = strExportBase
    m_blnKontextAnhaenge = blnAnhaenge
    m_blnKontextMSG = blnMSG
    m_blnKontextFilter = blnFilter
End Sub


' ===========================================================================
' PUFFER-OPERATIONEN
' ===========================================================================

' Mail-Daten zum Puffer hinzufuegen
' Gibt True zurueck wenn erfolgreich
Public Function BufferHinzufuegen(ByRef mk As TypMailKomplett) As Boolean
    If Not m_blnInitialisiert Then BufferInit

    ' Puffer-Array vergroessern wenn noetig
    If m_lngBufferAnzahl > UBound(m_aBuffer) Then
        ReDim Preserve m_aBuffer(0 To m_lngBufferAnzahl + m_lngBufferMax - 1)
    End If

    m_aBuffer(m_lngBufferAnzahl) = mk
    m_lngBufferAnzahl = m_lngBufferAnzahl + 1
    m_lngGesamtVerarbeitet = m_lngGesamtVerarbeitet + 1

    BufferHinzufuegen = True
End Function


' Prueft ob der Puffer voll ist (Schwellwert erreicht)
Public Function BufferIstVoll() As Boolean
    BufferIstVoll = (m_lngBufferAnzahl >= m_lngBufferMax)
End Function


' Aktuelle Puffergroesse (Anzahl Eintraege)
Public Function BufferGroesse() As Long
    BufferGroesse = m_lngBufferAnzahl
End Function


' ===========================================================================
' STATISTIK-ABFRAGEN
' ===========================================================================

Public Function BufferGesamtNeu() As Long
    BufferGesamtNeu = m_lngGesamtNeu
End Function

Public Function BufferGesamtDuplikate() As Long
    BufferGesamtDuplikate = m_lngGesamtDuplikate
End Function

Public Function BufferGesamtFehler() As Long
    BufferGesamtFehler = m_lngGesamtFehler
End Function

Public Function BufferGesamtVerarbeitet() As Long
    BufferGesamtVerarbeitet = m_lngGesamtVerarbeitet
End Function

Public Function BufferStatistik() As String
    BufferStatistik = "Verarbeitet=" & m_lngGesamtVerarbeitet & _
                      " Neu=" & m_lngGesamtNeu & _
                      " Dup=" & m_lngGesamtDuplikate & _
                      " Err=" & m_lngGesamtFehler
End Function


' ===========================================================================
' PUFFER FLUSH (Kernroutine: Batch-Schreiben in DB)
' ===========================================================================

' Schreibt alle gepufferten Mails in die Datenbank.
' Verwendet eine DAO-Transaktion fuer Atomizitaet + Performance.
' Gibt Anzahl neu gespeicherter Mails zurueck.
Public Function BufferFlush() As Long
    On Error GoTo ErrHandler

    If m_lngBufferAnzahl = 0 Then BufferFlush = 0: Exit Function

    Dim lngNeu As Long, lngDup As Long, lngErr As Long
    lngNeu = 0: lngDup = 0: lngErr = 0

    LogDebug "BufferFlush: " & m_lngBufferAnzahl & " Eintraege...", "BUFFER"

    ' --- 1. Batch-Duplikatpruefung ---
    Dim arrHashes() As String
    Dim arrExistiert() As Boolean
    Dim n As Long

    ReDim arrHashes(0 To m_lngBufferAnzahl - 1)
    ReDim arrExistiert(0 To m_lngBufferAnzahl - 1)

    For n = 0 To m_lngBufferAnzahl - 1
        arrHashes(n) = m_aBuffer(n).UniqueHash
    Next n

    Call PruefeHashes(arrHashes, arrExistiert)

    ' --- 2. Transaktion: Nicht-Duplikate in DB schreiben ---
    Dim ws As DAO.Workspace
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans

    On Error GoTo RollbackHandler

    For n = 0 To m_lngBufferAnzahl - 1
        ' Duplikat ueberspringen
        If arrExistiert(n) Then
            lngDup = lngDup + 1
            LogTrace "Duplikat: " & Left(m_aBuffer(n).Mail.Betreff, 40), "BUFFER"
            Call LoescheTempDateien(m_aBuffer(n))
            GoTo NaechsterEintrag
        End If

        ' Ungueltige Mail ueberspringen
        If Not m_aBuffer(n).Mail.IstGueltig Then
            lngErr = lngErr + 1
            Call LoescheTempDateien(m_aBuffer(n))
            GoTo NaechsterEintrag
        End If

        ' Mail verarbeiten und in DB speichern
        Dim lngEmailID As Long
        lngEmailID = SpeichereMailAusBuffer(m_aBuffer(n))

        If lngEmailID > 0 Then
            lngNeu = lngNeu + 1
        Else
            lngErr = lngErr + 1
            Call LoescheTempDateien(m_aBuffer(n))
        End If

NaechsterEintrag:
    Next n

    ws.CommitTrans
    On Error GoTo ErrHandler

    ' Statistik aktualisieren
    m_lngGesamtNeu = m_lngGesamtNeu + lngNeu
    m_lngGesamtDuplikate = m_lngGesamtDuplikate + lngDup
    m_lngGesamtFehler = m_lngGesamtFehler + lngErr

    ' Puffer leeren (Daten sind in DB)
    m_lngBufferAnzahl = 0
    If m_lngBufferMax > 0 Then
        ReDim m_aBuffer(0 To m_lngBufferMax - 1)
    End If

    LogDebug "BufferFlush: Neu=" & lngNeu & " Dup=" & lngDup & " Err=" & lngErr, "BUFFER"
    BufferFlush = lngNeu
    Exit Function

RollbackHandler:
    Dim strErr As String
    strErr = Err.Description
    On Error Resume Next
    ws.Rollback
    On Error GoTo 0

    LogError "BufferFlush ROLLBACK: " & strErr, "BUFFER"
    ' Bei Rollback: Temp-Dateien aufbewahren (koennen beim naechsten Versuch gebraucht werden)
    m_lngGesamtFehler = m_lngGesamtFehler + m_lngBufferAnzahl
    m_lngBufferAnzahl = 0
    If m_lngBufferMax > 0 Then
        ReDim m_aBuffer(0 To m_lngBufferMax - 1)
    End If
    BufferFlush = 0
    Exit Function

ErrHandler:
    LogVBAError "BufferFlush"
    BufferFlush = 0
End Function


' ===========================================================================
' EINZELNE MAIL AUS PUFFER IN DB SCHREIBEN (innerhalb Transaktion)
' ===========================================================================

Private Function SpeichereMailAusBuffer(ByRef mk As TypMailKomplett) As Long
    On Error GoTo ErrHandler

    Dim lngEmailID      As Long
    Dim lngKontaktID    As Long
    Dim lngThreadID     As Long
    Dim strEmailTyp     As String

    ' --- Absender-Kontakt ermitteln/erstellen ---
    If mk.Mail.AbsenderEmailTyp = "EX" Then
        strEmailTyp = "EX"
    Else
        strEmailTyp = "SMTP"
    End If
    lngKontaktID = GetOderErstelleKontakt(mk.Mail.AbsenderName, _
                                           mk.Mail.AbsenderEmail, strEmailTyp)

    ' --- Thread ermitteln/erstellen ---
    lngThreadID = GetOderErstelleThread(mk.Mail.Betreff, mk.Mail.InternetMessageID, _
                                         mk.Mail.InReplyTo, mk.Mail.EmpfangenAm, _
                                         mk.Mail.AbsenderName)

    ' --- Email-Datensatz speichern ---
    lngEmailID = SpeichereEmail( _
        mk.Mail.EntryID, mk.UniqueHash, lngThreadID, m_lngKontextOrdnerID, _
        lngKontaktID, m_lngKontextSyncLaufID, mk.Mail.Betreff, _
        mk.Mail.AbsenderName, mk.Mail.AbsenderEmail, mk.Mail.EmpfangenAm, _
        mk.Mail.GesendetAm, mk.Mail.Groesse, mk.Mail.Wichtigkeit, _
        mk.Mail.Gelesen, mk.Mail.HatAnhaenge, mk.Mail.AnhangAnzahl, _
        mk.Mail.MessageClass, mk.Mail.InternetMessageID)

    If lngEmailID = 0 Then
        SpeichereMailAusBuffer = 0
        Exit Function
    End If

    ' --- Content speichern ---
    Call SpeichereEmailContent(lngEmailID, mk.Mail.HTMLBody, mk.Mail.PlainTextBody)

    ' --- Empfaenger speichern ---
    Dim e As Integer
    Dim lngEmpfKontaktID As Long
    For e = 0 To mk.EmpfaengerAnzahl - 1
        If IstGueltigeEmail(mk.Empfaenger(e).Email) Then
            lngEmpfKontaktID = GetOderErstelleKontakt(mk.Empfaenger(e).Name, _
                                                       mk.Empfaenger(e).Email)
            ' E-Mail aktualisieren wenn bisherige ungueltig
            If lngEmpfKontaktID > 0 Then
                Call AktualisiereKontaktEmail(lngEmpfKontaktID, mk.Empfaenger(e).Email)
            End If
        Else
            lngEmpfKontaktID = 0
        End If

        Call SpeichereEmpfaenger(lngEmailID, lngEmpfKontaktID, _
                                  mk.Empfaenger(e).Typ, mk.Empfaenger(e).Name, _
                                  mk.Empfaenger(e).Email)
    Next e

    ' --- Anhang-Metadaten speichern + Datei-Queue ---
    Dim a As Integer
    Dim lngAnhangID As Long
    For a = 0 To mk.AnhangAnzahl - 1
        lngAnhangID = SpeichereAnhangMetadaten(lngEmailID, mk.Anhaenge(a).Dateiname, _
                                                 mk.Anhaenge(a).Groesse, mk.Anhaenge(a).MimeType, _
                                                 mk.Anhaenge(a).AnhangTyp, mk.Anhaenge(a).IstVersteckt)

        ' Datei-Queue: Temp -> Netzwerk (nur wenn Temp-Pfad vorhanden)
        If mk.Anhaenge(a).TempPfad <> "" And lngAnhangID > 0 Then
            Dim strZielDir As String
            strZielDir = NormalisierePfad(m_strKontextExportBase & m_strKontextProjekt & "\" & _
                                           m_strKontextPhase & "\Anhaenge\EmailID_" & lngEmailID & "\")
            Dim strEndung As String
            strEndung = HoleEndung(mk.Anhaenge(a).Dateiname)
            If strEndung = "" Then strEndung = "bin"

            Dim strZielPfad As String
            strZielPfad = EindeutigerDateipfad(strZielDir, _
                          BereinigeDateiname(mk.Anhaenge(a).Dateiname, 100), strEndung)

            Call DateiQueueHinzufuegen(mk.Anhaenge(a).TempPfad, strZielPfad, _
                                       "Anhang", lngEmailID, lngAnhangID)
        End If
    Next a

    ' --- MSG Datei-Queue ---
    If mk.MSGTempPfad <> "" Then
        Dim strMSGDir As String
        strMSGDir = NormalisierePfad(m_strKontextExportBase & m_strKontextProjekt & "\" & _
                                      m_strKontextPhase & "\MSG\")
        Dim strMSGDatei As String
        strMSGDatei = Format(mk.Mail.EmpfangenAm, "yyyymmdd_hhnn") & "_" & _
                       BereinigeDateiname(mk.Mail.Betreff, 50)

        Dim strMSGZiel As String
        strMSGZiel = EindeutigerDateipfad(strMSGDir, strMSGDatei, "msg")

        Call DateiQueueHinzufuegen(mk.MSGTempPfad, strMSGZiel, "MSG", lngEmailID, 0)
    End If

    ' --- Status ---
    Call SpeichereEmailStatus(lngEmailID, "Verarbeitet", "Sync", "Import via BufferFlush")

    SpeichereMailAusBuffer = lngEmailID
    Exit Function

ErrHandler:
    LogVBAError "SpeichereMailAusBuffer [" & Left(Nz(mk.Mail.Betreff, "?"), 30) & "]"
    SpeichereMailAusBuffer = 0
End Function


' ===========================================================================
' BATCH-DUPLIKATPRUEFUNG
' ===========================================================================

' Prueft mehrere Hashes auf einmal gegen die DB.
' Viel schneller als N einzelne DCount-Aufrufe bei Netzwerk-Backend.
Private Sub PruefeHashes(ByRef arrHashes() As String, ByRef arrExistiert() As Boolean)
    On Error GoTo ErrHandler

    Dim n As Long
    Dim lngAnzahl As Long
    lngAnzahl = UBound(arrHashes) + 1

    ' Fuer kleine Batches: einzelne Pruefung (einfacher)
    If lngAnzahl <= 5 Then
        For n = 0 To lngAnzahl - 1
            arrExistiert(n) = ExistiertMailHash(arrHashes(n))
        Next n
        Exit Sub
    End If

    ' Fuer groessere Batches: IN-Query
    ' SQL: SELECT UniqueHash FROM tblEmails WHERE UniqueHash IN ('h1','h2',...)
    Dim strIN As String
    strIN = ""
    For n = 0 To lngAnzahl - 1
        If strIN <> "" Then strIN = strIN & ","
        strIN = strIN & "'" & SQLSafe(arrHashes(n)) & "'"
    Next n

    ' Alle existierenden Hashes in Dictionary laden
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT UniqueHash FROM tblEmails WHERE UniqueHash IN (" & _
                               strIN & ")", dbOpenSnapshot)

    Do While Not rs.EOF
        dict(CStr(rs!UniqueHash)) = True
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing: Set db = Nothing

    ' Ergebnis-Array fuellen
    For n = 0 To lngAnzahl - 1
        arrExistiert(n) = dict.Exists(arrHashes(n))
    Next n

    Set dict = Nothing
    Exit Sub

ErrHandler:
    ' Fallback: einzelne Pruefung
    LogWarn "Batch-Hash-Pruefung fehlgeschlagen, Fallback auf Einzel: " & _
            Err.Description, "BUFFER"
    On Error Resume Next
    For n = 0 To UBound(arrHashes)
        arrExistiert(n) = ExistiertMailHash(arrHashes(n))
        If Err.Number <> 0 Then arrExistiert(n) = False: Err.Clear
    Next n
    On Error GoTo 0
End Sub


' ===========================================================================
' DATEI-QUEUE (Temp -> Netzwerk)
' ===========================================================================

' Datei-Operation zur Queue hinzufuegen
Public Sub DateiQueueHinzufuegen(ByVal strQuelle As String, _
                                  ByVal strZiel As String, _
                                  ByVal strTyp As String, _
                                  ByVal lngEmailID As Long, _
                                  ByVal lngAnhangID As Long)
    ' Queue vergroessern wenn noetig
    If m_lngDateiQueueAnzahl > UBound(m_aDateiQueue) Then
        ReDim Preserve m_aDateiQueue(0 To m_lngDateiQueueAnzahl + 99)
    End If

    With m_aDateiQueue(m_lngDateiQueueAnzahl)
        .QuellPfad = strQuelle
        .ZielPfad = strZiel
        .OperationsTyp = strTyp
        .EmailID = lngEmailID
        .AnhangID = lngAnhangID
        .Versuche = 0
    End With

    m_lngDateiQueueAnzahl = m_lngDateiQueueAnzahl + 1
End Sub


' Datei-Queue abarbeiten (alle Dateien von Temp auf Ziellaufwerk kopieren)
' Gibt Anzahl erfolgreich kopierter Dateien zurueck.
Public Function DateiQueueVerarbeiten() As Long
    On Error GoTo ErrHandler

    If m_lngDateiQueueAnzahl = 0 Then DateiQueueVerarbeiten = 0: Exit Function

    Dim lngOK As Long, lngFail As Long
    Dim intMaxRetries As Integer
    lngOK = 0: lngFail = 0
    intMaxRetries = CInt(LeseConfig("NetzwerkRetries", "3"))

    LogDebug "DateiQueue: " & m_lngDateiQueueAnzahl & " Dateien verarbeiten...", "BUFFER"

    Dim n As Long
    For n = 0 To m_lngDateiQueueAnzahl - 1
        If g_blnAbbrechen Then Exit For

        If DateiKopierenMitRetry(m_aDateiQueue(n).QuellPfad, _
                                  m_aDateiQueue(n).ZielPfad, _
                                  intMaxRetries) Then
            ' Erfolg: DB-Pfad aktualisieren
            Select Case m_aDateiQueue(n).OperationsTyp
                Case "MSG"
                    Call SetzeEmailMSGPfad(m_aDateiQueue(n).EmailID, _
                                           m_aDateiQueue(n).ZielPfad)
                Case "Anhang"
                    Call AktualisiereAnhangPfad(m_aDateiQueue(n).AnhangID, _
                                                m_aDateiQueue(n).ZielPfad)
            End Select

            lngOK = lngOK + 1

            ' Temp-Datei loeschen (erfolgreich kopiert)
            On Error Resume Next
            Kill m_aDateiQueue(n).QuellPfad
            On Error GoTo ErrHandler
        Else
            lngFail = lngFail + 1
            LogWarn "Datei-Kopie endgueltig fehlgeschlagen: " & _
                    m_aDateiQueue(n).QuellPfad, "BUFFER"
        End If

        ' UI responsiv halten (alle 10 Dateien)
        If n Mod 10 = 0 Then DoEvents
    Next n

    LogInfo "DateiQueue: " & lngOK & " OK, " & lngFail & " fehlgeschlagen " & _
            "von " & m_lngDateiQueueAnzahl & " gesamt", "BUFFER"

    ' Queue leeren
    m_lngDateiQueueAnzahl = 0
    ReDim m_aDateiQueue(0 To 99)

    DateiQueueVerarbeiten = lngOK
    Exit Function

ErrHandler:
    LogVBAError "DateiQueueVerarbeiten"
    DateiQueueVerarbeiten = 0
End Function


' Datei-Queue Groesse abfragen
Public Function DateiQueueGroesse() As Long
    DateiQueueGroesse = m_lngDateiQueueAnzahl
End Function


' ===========================================================================
' DATEI-KOPIE MIT RETRY + EXPONENTIELLES BACKOFF
' ===========================================================================

' Kopiert eine Datei mit konfigurierbarer Retry-Anzahl
' Bei Netzwerk-Fehlern: warten, Pause verdoppeln, nochmal versuchen
Private Function DateiKopierenMitRetry(ByVal strQuelle As String, _
                                        ByVal strZiel As String, _
                                        ByVal intMaxVersuche As Integer) As Boolean
    On Error GoTo ErrHandler

    ' Quelldatei pruefen
    If Dir(strQuelle) = "" Then
        LogWarn "Quelldatei nicht gefunden: " & strQuelle, "BUFFER"
        DateiKopierenMitRetry = False
        Exit Function
    End If

    ' Zielverzeichnis erstellen (mit Retry)
    Dim strZielDir As String
    strZielDir = Left(strZiel, InStrRev(strZiel, "\"))
    If strZielDir <> "" Then ErstelleOrdner strZielDir

    ' Retry-Schleife mit exponentiellem Backoff
    Dim intVersuch As Integer
    Dim lngPause As Long
    lngPause = CLng(LeseConfig("NetzwerkRetryPause", "2000"))

    For intVersuch = 1 To intMaxVersuche
        On Error Resume Next
        FileCopy strQuelle, strZiel

        If Err.Number = 0 Then
            ' Erfolg!
            On Error GoTo 0
            LogTrace "Datei kopiert: " & strZiel, "BUFFER"
            DateiKopierenMitRetry = True
            Exit Function
        End If

        ' Fehler
        Dim strFehler As String
        strFehler = Err.Description
        Err.Clear
        On Error GoTo ErrHandler

        If intVersuch < intMaxVersuche Then
            LogWarn "Datei-Kopie Versuch " & intVersuch & "/" & intMaxVersuche & _
                    " fehlgeschlagen: " & strFehler & _
                    " - Retry in " & lngPause & "ms", "BUFFER"
            Sleep lngPause
            ' Exponentielles Backoff
            lngPause = lngPause * 2
        Else
            LogError "Datei-Kopie endgueltig fehlgeschlagen nach " & intMaxVersuche & _
                     " Versuchen: " & strFehler, "BUFFER"
        End If
    Next intVersuch

    DateiKopierenMitRetry = False
    Exit Function

ErrHandler:
    DateiKopierenMitRetry = False
End Function


' ===========================================================================
' AUFRAEUMEN
' ===========================================================================

' Puffer und Queue komplett leeren + Statistik zuruecksetzen
Public Sub BufferLeeren()
    m_lngBufferAnzahl = 0
    m_lngDateiQueueAnzahl = 0

    If m_lngBufferMax > 0 Then
        ReDim m_aBuffer(0 To m_lngBufferMax - 1)
    End If
    ReDim m_aDateiQueue(0 To 99)

    m_blnInitialisiert = False
    LogDebug "Buffer geleert", "BUFFER"
End Sub


' Temp-Dateien eines einzelnen TypMailKomplett loeschen
' (z.B. wenn Duplikat oder Fehler)
Private Sub LoescheTempDateien(ByRef mk As TypMailKomplett)
    On Error Resume Next

    ' MSG Temp
    If mk.MSGTempPfad <> "" Then
        Kill mk.MSGTempPfad
        mk.MSGTempPfad = ""
    End If

    ' Anhang Temp
    Dim a As Integer
    For a = 0 To mk.AnhangAnzahl - 1
        If mk.Anhaenge(a).TempPfad <> "" Then
            Kill mk.Anhaenge(a).TempPfad
            mk.Anhaenge(a).TempPfad = ""
        End If
    Next a

    On Error GoTo 0
End Sub
