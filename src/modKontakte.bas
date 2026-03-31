Option Compare Database
Option Explicit

' ===========================================================================
' modKontakte - Kontaktverarbeitung und Namens-Analyse
' ===========================================================================
' Konsolidiert die gesamte Kontaktlogik aus dem alten Projekt:
'   - Name parsing (Titel/Vorname/Nachname/Institution)
'   - Fallback-Name aus E-Mail-Adresse
'   - Geschlechtserkennung fuer Anrede
'   - Kontakt-Plausibilitaetspruefung
'   - Anredeform-Generierung
'   - Domain-basiertes Lernen von bestehenden Kontakten
'
' Abhaengigkeiten: modStringUtils (CapitalizeWord, IsAlphaOnly,
'                  IstGueltigeEmail, ExtrahiereInstitution, ExtrahiereDomain,
'                  BereinigeNameInput, BildeSortiername)
' ===========================================================================


' ---------------------------------------------------------------------------
' TYPE: Geparster Kontaktname
' ---------------------------------------------------------------------------
Public Type TypKontaktName
    Titel           As String   ' "Dr.", "Prof. Dr.", etc.
    Vorname         As String
    Nachname        As String
    Namenszusatz    As String   ' Mittlere Namensteile
    Institution     As String   ' Aus Klammer oder Domain
    Anzeigename     As String   ' Originalname bereinigt
    Sortiername     As String   ' "Nachname, Vorname"
    IstPlausibel    As Boolean  ' Plausibilitaet ok?
End Type


' ===========================================================================
' NAMENS-PARSING (aus Anzeigename -> Strukturierter Kontaktname)
' ===========================================================================

' Zerlegt einen Roh-Namen in Titel/Vorname/Nachname/Namenszusatz/Institution.
' Erkennt Formate:
'   "Dr. Max Mueller"             -> Titel=Dr., Vorname=Max, Nachname=Mueller
'   "Mueller, Max"                -> Vorname=Max, Nachname=Mueller
'   "Prof. Dr. Hans von Hohenheim" -> Titel=Prof. Dr., Vorname=Hans, Nachname=Hohenheim
'   "Jan Meier (LUBW)"           -> Vorname=Jan, Nachname=Meier, Institution=LUBW
Public Function ParseKontaktName(ByVal strNameRaw As String, _
                                  Optional ByVal strEmail As String = "") As TypKontaktName
    Dim result As TypKontaktName

    ' --- Bereinigung ---
    strNameRaw = BereinigeNameInput(strNameRaw)
    strNameRaw = Replace(strNameRaw, "'", "")
    result.Anzeigename = strNameRaw

    ' --- Technischer Fallback: Name = E-Mail oder leer? ---
    If strNameRaw = "" Or _
       (IstGueltigeEmail(strNameRaw) And LCase(strNameRaw) = LCase(strEmail)) Then
        result = FallbackNameAusEmail(strEmail)
        Exit Function
    End If

    ' --- Institution aus Klammern extrahieren: "Name (Firma)" ---
    If InStr(strNameRaw, "(") > 0 And InStr(strNameRaw, ")") > InStr(strNameRaw, "(") Then
        Dim strKlammer As String
        strKlammer = Mid(strNameRaw, InStr(strNameRaw, "(") + 1, _
                         InStr(strNameRaw, ")") - InStr(strNameRaw, "(") - 1)
        result.Institution = CapitalizeWord(Trim(strKlammer))
        strNameRaw = Trim(Replace(strNameRaw, "(" & strKlammer & ")", ""))
    End If

    ' --- Titel extrahieren ---
    Dim arrTitel As Variant
    Dim t As Long
    arrTitel = Array("Prof. Dr.", "Dr. med.", "Dr.", "Prof.", _
                     "Dipl.-Ing.", "Dipl.-Kfm.", "Mag.", "Ing.")

    For t = LBound(arrTitel) To UBound(arrTitel)
        If LCase(Left(strNameRaw, Len(arrTitel(t)))) = LCase(arrTitel(t)) Then
            result.Titel = CStr(arrTitel(t))
            strNameRaw = Trim(Mid(strNameRaw, Len(arrTitel(t)) + 1))
            Exit For
        End If
    Next t

    ' --- Format: "Nachname, Vorname" ---
    If InStr(strNameRaw, ",") > 0 Then
        Dim parts() As String
        parts = Split(strNameRaw, ",")
        result.Nachname = CapitalizeWord(Trim(parts(0)))
        result.Vorname = CapitalizeWord(Trim(parts(1)))
        GoTo Finish
    End If

    ' --- Standardformat: "Vorname [Zusatz] Nachname" ---
    Dim arrParts() As String
    arrParts = Split(strNameRaw, " ")

    Select Case UBound(arrParts)
        Case 0
            ' Nur ein Wort -> Nachname
            result.Nachname = CapitalizeWord(arrParts(0))
        Case 1
            ' Zwei Worte -> Vorname Nachname
            result.Vorname = CapitalizeWord(arrParts(0))
            result.Nachname = CapitalizeWord(arrParts(1))
        Case Else
            ' Drei+ Worte -> Vorname [Zusatz...] Nachname
            result.Vorname = CapitalizeWord(arrParts(0))
            result.Nachname = CapitalizeWord(arrParts(UBound(arrParts)))

            ' Mittlere Teile als Namenszusatz
            If UBound(arrParts) >= 2 Then
                Dim strMid As String
                Dim m As Long
                strMid = ""
                For m = 1 To UBound(arrParts) - 1
                    If strMid <> "" Then strMid = strMid & " "
                    strMid = strMid & arrParts(m)
                Next m
                result.Namenszusatz = strMid
            End If
    End Select

Finish:
    ' Institution aus Domain wenn noch leer
    If result.Institution = "" And strEmail <> "" Then
        result.Institution = ExtrahiereInstitution(strEmail)
    End If

    ' Sortiername bilden
    result.Sortiername = BildeSortiername(result.Nachname, result.Vorname)

    ' Anzeigename ggf. setzen
    If result.Anzeigename = "" Then
        result.Anzeigename = Trim(result.Vorname & " " & result.Nachname)
    End If

    ' Plausibilitaet pruefen
    result.IstPlausibel = IstKontaktPlausibel(result.Vorname, result.Nachname)

    ParseKontaktName = result
End Function


' ===========================================================================
' FALLBACK NAME AUS E-MAIL (wenn kein Anzeigename vorhanden)
' ===========================================================================

' Leitet Vor-/Nachname aus dem E-Mail-Format ab:
'   "vorname.nachname@firma.de"  -> Vorname=Vorname, Nachname=Nachname
'   "v.nachname@firma.de"        -> Nachname=Nachname
'   "info@firma.de"              -> Institution=Firma
Public Function FallbackNameAusEmail(ByVal strEmail As String) As TypKontaktName
    Dim result As TypKontaktName

    If InStr(strEmail, "@") = 0 Then
        result.Nachname = DEFAULT_NAME
        result.Anzeigename = DEFAULT_NAME
        FallbackNameAusEmail = result
        Exit Function
    End If

    Dim strUser As String
    strUser = Split(strEmail, "@")(0)

    result.Institution = ExtrahiereInstitution(strEmail)

    ' User-Part aufteilen (punkt/unterstrich/bindestrich)
    strUser = Replace(strUser, "_", ".")
    strUser = Replace(strUser, "-", ".")

    Dim arrUser() As String
    arrUser = Split(strUser, ".")

    If UBound(arrUser) >= 1 Then
        ' Zwei+ Teile: vorname.nachname
        result.Vorname = CapitalizeWord(arrUser(0))
        result.Nachname = CapitalizeWord(arrUser(UBound(arrUser)))
    ElseIf UBound(arrUser) = 0 Then
        ' Ein Teil: nur Nachname (oder Institution)
        If IstSystemEmail(strEmail) Then
            ' System-Adressen: Institution als Nachname
            result.Nachname = result.Institution
            If result.Nachname = "" Then result.Nachname = CapitalizeWord(arrUser(0))
        Else
            result.Nachname = CapitalizeWord(arrUser(0))
        End If
    End If

    ' Plausibilitaet und Sortierung
    result.IstPlausibel = IstKontaktPlausibel(result.Vorname, result.Nachname)
    result.Sortiername = BildeSortiername(result.Nachname, result.Vorname)

    If result.Vorname <> "" Then
        result.Anzeigename = result.Vorname & " " & result.Nachname
    ElseIf result.Nachname <> "" Then
        result.Anzeigename = result.Nachname
    Else
        result.Anzeigename = DEFAULT_NAME
    End If

    FallbackNameAusEmail = result
End Function


' ===========================================================================
' SYSTEM-E-MAIL ERKENNUNG
' ===========================================================================

' Erkennt generische Adressen wie info@, noreply@, support@ etc.
Public Function IstSystemEmail(ByVal strEmail As String) As Boolean
    If InStr(strEmail, "@") = 0 Then IstSystemEmail = False: Exit Function

    Dim strUser As String
    strUser = LCase(Split(strEmail, "@")(0))

    Dim arrSystem As Variant, s As Variant
    arrSystem = Array("info", "kontakt", "mail", "service", "support", _
                      "admin", "post", "webmaster", "noreply", "no.reply", _
                      "no-reply", "newsletter", "office", "empfang")

    For Each s In arrSystem
        If strUser = CStr(s) Or Left(strUser, Len(CStr(s)) + 1) = CStr(s) & "." Then
            IstSystemEmail = True
            Exit Function
        End If
    Next s

    IstSystemEmail = False
End Function


' ===========================================================================
' GESCHLECHTSERKENNUNG (fuer Anrede)
' ===========================================================================

' Gibt "m", "w" oder "" (unbekannt) zurueck.
' Basiert auf einer einfachen Liste gaengiger deutscher Vornamen.
Public Function ErkenneGeschlecht(ByVal strVorname As String) As String
    Dim strLow As String
    strLow = LCase(Trim(strVorname))
    If strLow = "" Then ErkenneGeschlecht = "": Exit Function

    Dim arrW As Variant, arrM As Variant, v As Variant

    arrW = Array("anna", "andrea", "angelika", "anke", "barbara", "birgit", _
                 "brigitte", "christina", "christa", "claudia", "cornelia", _
                 "daniela", "doris", "elena", "elisabeth", "eva", _
                 "gabriele", "gudrun", "heike", "helga", "ines", "ingrid", _
                 "jana", "julia", "karin", "katja", "katrin", "laura", "lena", _
                 "lisa", "maria", "marina", "martina", "monika", "nadine", _
                 "nicole", "petra", "renate", "ruth", "sabine", "sandra", _
                 "silke", "simone", "sophie", "stefanie", "susanne", "tanja", _
                 "ulrike", "ursula", "ute", "verena")

    arrM = Array("alexander", "andreas", "axel", "bernd", "bernhard", _
                 "christian", "christoph", "daniel", "dieter", "dirk", _
                 "erik", "florian", "frank", "franz", "georg", "gerald", _
                 "gerd", "gerhard", "guenter", "hans", "harald", "heinz", _
                 "helmut", "jan", "jens", "joerg", "johann", "juergen", _
                 "karl", "klaus", "lars", "lukas", "manfred", "markus", _
                 "martin", "matthias", "max", "michael", "norbert", "olaf", _
                 "oliver", "patrick", "paul", "peter", "philipp", "rainer", _
                 "ralph", "rolf", "sebastian", "stefan", "stephan", _
                 "thomas", "tobias", "ulrich", "uwe", "volker", "walter", _
                 "werner", "wolfgang")

    For Each v In arrW
        If strLow = CStr(v) Then ErkenneGeschlecht = "w": Exit Function
    Next v

    For Each v In arrM
        If strLow = CStr(v) Then ErkenneGeschlecht = "m": Exit Function
    Next v

    ErkenneGeschlecht = ""
End Function


' ===========================================================================
' ANREDEFORM GENERIEREN
' ===========================================================================

' Erzeugt eine formale Anrede:
'   "Sehr geehrter Herr Dr. Mueller"
'   "Sehr geehrte Frau Prof. Dr. Meier"
'   "Sehr geehrte/r Mueller" (wenn Geschlecht unbekannt)
Public Function BildeAnrede(ByVal strTitel As String, _
                             ByVal strVorname As String, _
                             ByVal strNachname As String) As String
    Dim strGeschlecht As String
    strGeschlecht = ErkenneGeschlecht(strVorname)

    Dim strPrefix As String
    Select Case strGeschlecht
        Case "m": strPrefix = "Sehr geehrter Herr"
        Case "w": strPrefix = "Sehr geehrte Frau"
        Case Else: strPrefix = "Sehr geehrte/r"
    End Select

    If strNachname = "" And strVorname = "" Then
        BildeAnrede = "Sehr geehrte Damen und Herren"
        Exit Function
    End If

    Dim strResult As String
    strResult = strPrefix

    If strTitel <> "" Then strResult = strResult & " " & strTitel
    If strNachname <> "" Then
        strResult = strResult & " " & strNachname
    ElseIf strVorname <> "" Then
        strResult = strResult & " " & strVorname
    End If

    BildeAnrede = Trim(strResult)
End Function


' ===========================================================================
' KONTAKT-PLAUSIBILITAET
' ===========================================================================

' Pruefen ob Vor-/Nachname plausibel sind (nicht zu kurz, nur Alpha-Zeichen)
Public Function IstKontaktPlausibel(ByVal strVorname As String, _
                                     ByVal strNachname As String) As Boolean
    IstKontaktPlausibel = True

    ' Nachname: mindestens 2 Zeichen, nur Alpha
    If Len(Trim(strNachname)) < 2 Then IstKontaktPlausibel = False: Exit Function
    If Not IsAlphaOnly(strNachname) Then IstKontaktPlausibel = False: Exit Function

    ' Vorname: optional, aber wenn vorhanden dann plausibel
    If Len(Trim(strVorname)) > 0 Then
        If Len(Trim(strVorname)) < 2 Then IstKontaktPlausibel = False: Exit Function
        If Not IsAlphaOnly(strVorname) Then IstKontaktPlausibel = False: Exit Function
    End If
End Function


' ===========================================================================
' DOMAIN-BASIERTES LERNEN
' ===========================================================================

' Versucht fuer eine E-Mail-Domain bereits bekannte Kontaktdaten
' (Institution, Titel-Muster) aus bestehenden Kontakten zu uebernehmen.
' Gibt ein Dictionary mit erkannten Werten zurueck (leer wenn nichts gefunden).
Public Function LerneVonDomain(ByVal strEmail As String) As Object
    On Error GoTo ErrHandler

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim strDomain As String
    strDomain = ExtrahiereDomain(strEmail)
    If strDomain = "" Then Set LerneVonDomain = dict: Exit Function

    ' Basis-Domain extrahieren (z.B. "hydron-gmbh" aus "hydron-gmbh.de")
    Dim strBase As String
    strBase = strDomain
    If InStr(strBase, ".") > 0 Then strBase = Split(strBase, ".")(0)

    ' Suche: anderer Kontakt mit gleicher Domain, der Vor+Nachname hat
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb

    Set rs = db.OpenRecordset( _
        "SELECT TOP 1 Vorname, Nachname, Titel, Institution FROM [" & TBL_KONTAKTE & "] " & _
        "WHERE Email LIKE '*@" & strBase & "*' " & _
        "AND Nz(Vorname,'') <> '' AND Nz(Nachname,'') <> ''", dbOpenSnapshot)

    If Not rs.EOF Then
        If Not IsNull(rs!Institution) And Nz(rs!Institution, "") <> "" Then
            dict.Add "Institution", rs!Institution
        End If
        ' Hinweis: Wir uebernehmen NUR Institution von Domain-Geschwistern.
        ' Vor-/Nachname sind individuell, nicht domain-basiert lernbar.
    End If

    rs.Close: Set rs = Nothing: Set db = Nothing
    Set LerneVonDomain = dict
    Exit Function

ErrHandler:
    LogWarn "LerneVonDomain fehlgeschlagen fuer '" & strEmail & "': " & Err.Description, "KONTAKT"
    Set LerneVonDomain = dict
End Function


' ===========================================================================
' E-MAIL-AKTUALISIERUNG BEI KONTAKTEN
' ===========================================================================

' Prueft ob ein Kontakt eine bessere E-Mail benoetigt
' (aktuell "Unbekannt" oder /O= Exchange-Pfad, neue E-Mail ist valide)
Public Function BrauchtEmailUpdate(ByVal lngKontaktID As Long, _
                                    ByVal strNeueEmail As String) As Boolean
    On Error GoTo ErrHandler
    BrauchtEmailUpdate = False

    If Not IstGueltigeEmail(strNeueEmail) Then Exit Function
    If LCase(strNeueEmail) = LCase(DEFAULT_EMAIL) Then Exit Function
    If Left(strNeueEmail, 3) = "/O=" Then Exit Function

    Dim strAlt As String
    strAlt = Nz(DLookup("Email", TBL_KONTAKTE, "KontaktID=" & lngKontaktID), "")

    If Trim(strAlt) = "" _
       Or LCase(Trim(strAlt)) = LCase(DEFAULT_EMAIL) _
       Or Left(strAlt, 3) = "/O=" Then
        BrauchtEmailUpdate = True
    End If
    Exit Function

ErrHandler:
    BrauchtEmailUpdate = False
End Function


' Aktualisiert die E-Mail eines Kontakts wenn die bisherige ungueltig ist
Public Sub AktualisiereKontaktEmail(ByVal lngKontaktID As Long, _
                                     ByVal strNeueEmail As String)
    On Error GoTo ErrHandler

    If Not BrauchtEmailUpdate(lngKontaktID, strNeueEmail) Then Exit Sub

    CurrentDb.Execute "UPDATE [" & TBL_KONTAKTE & "] SET " & _
                      "Email='" & SQLSafe(strNeueEmail) & "', " & _
                      "AktualisiertAm=" & SQLJetzt() & " " & _
                      "WHERE KontaktID=" & lngKontaktID, dbFailOnError

    LogDebug "Kontakt-Email aktualisiert: ID=" & lngKontaktID & " -> " & strNeueEmail, "KONTAKT"
    Exit Sub

ErrHandler:
    LogWarn "Kontakt-Email-Update fehlgeschlagen: " & Err.Description, "KONTAKT"
End Sub


