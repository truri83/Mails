Option Compare Database
Option Explicit

' ===========================================================================
' modStringUtils - String-Bereinigung, Validierung, Pfad- und Datei-Utilities
' ===========================================================================
' v0.4: Zentralisiert ALLE String-/Pfad-/Dateinamen-Operationen:
'
' VALIDIERUNG:
'   IstGueltigeEmail     - RegExp E-Mail-Pruefung
'   IsAlphaOnly          - Nur Buchstaben/Umlaute/Bindestrich
'
' TEXT-BEREINIGUNG:
'   BereinigeBetreff     - RE:/AW:/FW:/WG:/EXTERN entfernen
'   BereinigeDateiname   - Illegale Zeichen, Laengenlimit
'   BereinigeAnhangName  - Anhang-spezifisch mit Extension-Handling
'   BereinigeKontaktInfo - Name + Email standardisieren
'   BereinigeNameInput   - Sonderzeichen aus Namen
'   CapitalizeWord       - Gross-/Kleinschreibung + Adelspraedikate
'   SQLSafe              - SQL-Injection-Schutz
'   SQLDatum             - Locale-sicheres Datum fuer Jet SQL
'   SQLJetzt             - SQLDatum(Now) Kurzform
'   SQLNurDatum          - Nur Datum ohne Uhrzeit
'   KuerzeAbsender       - Absender auf Kurzform
'   RTrimSpecial         - Trailing Punkte/Leerzeichen (Windows)
'
' DATEINAME-GENERIERUNG:
'   GeneriereMailDateiname  - YYYYMMDD_HHNN_Absender_Betreff.ext
'   EntferneEndung          - Extension abtrennen
'   HoleEndung              - Extension extrahieren (lowercase)
'
' PFAD-OPERATIONEN:
'   NormalisierePfad      - UNC-sicher, Trailing Backslash
'   ErstelleOrdner        - Rekursiv mit UNC-Support
'   EindeutigerDateipfad  - Vermeidet Ueberschreiben (Name (1).ext)
'   BaueAblagePfad        - Basis/Projekt/Phase/Jahr/Monat
'   KuerzePfadSicher      - MAX_PATH-sichere Kuerzung
'
' ALLGEMEIN:
'   SafeText               - Null-sicheres Left(Nz(x), n)
'
' DOMAIN/INSTITUTION:
'   ExtrahiereDomain       - user@rps.bwl.de -> rps.bwl.de
'   ExtrahiereInstitution  - Domain -> RPS, LUBW etc.
'   BildeSortiername       - Nachname, Vorname
'   BereinigeOutlookPfad   - MAPI-Pfad normalisieren
'
' Abhaengigkeiten: modLogging (nur fuer ErstelleOrdner-Fehler)
' ===========================================================================


' ---------------------------------------------------------------------------
' E-MAIL-VALIDIERUNG per RegExp
' ---------------------------------------------------------------------------
Public Function IstGueltigeEmail(ByVal strEmail As String) As Boolean
    On Error Resume Next
    Dim objRE As Object

    If Nz(strEmail, "") = "" Then IstGueltigeEmail = False: Exit Function
    If Left(strEmail, 3) = "/O=" Then IstGueltigeEmail = False: Exit Function

    Set objRE = CreateObject("VBScript.RegExp")
    objRE.Pattern = "^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
    objRE.IgnoreCase = True
    objRE.Global = False

    IstGueltigeEmail = objRE.Test(strEmail)
    Set objRE = Nothing
End Function


' ---------------------------------------------------------------------------
' BETREFF BEREINIGEN (RE:, AW:, FW:, WG:, FWD:, EXTERN etc. entfernen)
' ---------------------------------------------------------------------------
Public Function BereinigeBetreff(ByVal strBetreff As String) As String
    On Error Resume Next
    Dim objRE As Object
    Set objRE = CreateObject("VBScript.RegExp")

    objRE.Pattern = "^(?:(?:RE|AW|FW|FWD|WG)\s*[:::]\s*|EXTERN\s+)+"
    objRE.IgnoreCase = True
    objRE.Global = True

    BereinigeBetreff = Trim(objRE.Replace(Nz(strBetreff, ""), ""))
    If BereinigeBetreff = "" Then BereinigeBetreff = "(Kein Betreff)"
    Set objRE = Nothing
End Function


' ---------------------------------------------------------------------------
' DATEINAMEN BEREINIGEN (illegale Zeichen entfernen/ersetzen)
' ---------------------------------------------------------------------------
Public Function BereinigeDateiname(ByVal strName As String, _
                                    Optional ByVal intMaxLen As Integer = 100) As String
    Dim s As String
    s = Nz(strName, "Unbenannt")

    ' Unzulaessige Zeichen entfernen/ersetzen
    s = Replace(s, "/", "-")
    s = Replace(s, "\", "-")
    s = Replace(s, ":", "-")
    s = Replace(s, "*", "")
    s = Replace(s, "?", "")
    s = Replace(s, Chr(34), "")   ' Anfuehrungszeichen
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "|", "")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    ' Mehrfach-Leerzeichen reduzieren
    Dim objRE As Object
    Set objRE = CreateObject("VBScript.RegExp")
    objRE.Pattern = "\s+"
    objRE.Global = True
    s = objRE.Replace(s, " ")
    Set objRE = Nothing

    s = Trim(s)
    If Len(s) > intMaxLen Then s = Left(s, intMaxLen)
    If s = "" Then s = "Unbenannt"

    BereinigeDateiname = s
End Function


' ---------------------------------------------------------------------------
' DATEIENDUNG ENTFERNEN
' ---------------------------------------------------------------------------
Public Function EntferneEndung(ByVal strDateiname As String) As String
    Dim pos As Long
    pos = InStrRev(strDateiname, ".")
    If pos > 1 Then
        EntferneEndung = Left(strDateiname, pos - 1)
    Else
        EntferneEndung = strDateiname
    End If
End Function


' ---------------------------------------------------------------------------
' DATEIENDUNG EXTRAHIEREN (lowercase)
' ---------------------------------------------------------------------------
Public Function HoleEndung(ByVal strDateiname As String) As String
    Dim pos As Long
    pos = InStrRev(strDateiname, ".")
    If pos > 0 And pos < Len(strDateiname) Then
        HoleEndung = LCase(Mid(strDateiname, pos + 1))
    Else
        HoleEndung = ""
    End If
End Function


' ---------------------------------------------------------------------------
' KONTAKTINFO BEREINIGEN (Name + Email)
' Gibt Dictionary zurueck mit Keys "Name" und "Email"
' ---------------------------------------------------------------------------
Public Function BereinigeKontaktInfo(ByVal strName As String, _
                                      ByVal strEmail As String, _
                                      Optional ByVal strDefaultName As String = "Unbekannt", _
                                      Optional ByVal strDefaultEmail As String = "unbekannt@example.com") As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Name bereinigen
    strName = Trim(Nz(strName, ""))
    strName = Replace(strName, Chr(34), "")   ' Doppelte Anfuehrungszeichen
    strName = Replace(strName, "'", "")
    strName = Replace(strName, ";", "")
    If strName = "" Then strName = strDefaultName

    ' Email bereinigen
    strEmail = LCase(Trim(Nz(strEmail, "")))
    strEmail = Replace(strEmail, "mailto:", "")
    strEmail = Replace(strEmail, Chr(34), "")
    strEmail = Replace(strEmail, "'", "")
    strEmail = Replace(strEmail, ";", "")
    If Not IstGueltigeEmail(strEmail) Then strEmail = strDefaultEmail

    dict.Add "Name", strName
    dict.Add "Email", strEmail
    Set BereinigeKontaktInfo = dict
End Function


' ---------------------------------------------------------------------------
' PFAD NORMALISIEREN (Backslash am Ende, doppelte Backslashes entfernen)
' ---------------------------------------------------------------------------
Public Function NormalisierePfad(ByVal strPfad As String, _
                                  Optional ByVal blnAlsOrdner As Boolean = True) As String
    If Nz(strPfad, "") = "" Then
        NormalisierePfad = Environ("USERPROFILE") & PATH_DEFAULT_FALLBACK
        Exit Function
    End If

    Dim blnUNC As Boolean
    strPfad = Trim(strPfad)
    blnUNC = (Left(strPfad, 2) = "\\")

    ' Trailing Backslashes entfernen
    Do While Right(strPfad, 1) = "\"
        strPfad = Left(strPfad, Len(strPfad) - 1)
    Loop

    ' Doppelte Backslashes entfernen (ab Position 2 wegen UNC)
    Do While InStr(2, strPfad, "\\") > 0
        strPfad = Left(strPfad, 1) & Replace(Mid(strPfad, 2), "\\", "\")
    Loop

    If blnAlsOrdner Then
        NormalisierePfad = strPfad & "\"
    Else
        NormalisierePfad = strPfad
    End If
End Function


' ---------------------------------------------------------------------------
' ORDNER ERSTELLEN (rekursiv, inkl. Unterordner)
' ---------------------------------------------------------------------------
Public Sub ErstelleOrdner(ByVal strOrdner As String)
    If Nz(strOrdner, "") = "" Then Exit Sub
    If Dir(strOrdner, vbDirectory) <> "" Then Exit Sub

    Dim arrTeile    As Variant
    Dim i           As Long
    Dim strTeil     As String
    Dim blnUNC      As Boolean

    blnUNC = (Left(strOrdner, 2) = "\\")
    arrTeile = Split(strOrdner, "\")

    If blnUNC Then
        If UBound(arrTeile) < 2 Then Exit Sub
        strTeil = "\\" & arrTeile(2) & "\"
        i = 3
    Else
        strTeil = arrTeile(0) & "\"
        i = 1
    End If

    For i = i To UBound(arrTeile)
        If arrTeile(i) <> "" Then
            strTeil = strTeil & arrTeile(i) & "\"
            If Dir(strTeil, vbDirectory) = "" Then
                On Error Resume Next
                MkDir strTeil
                If Err.Number <> 0 Then
                    LogError "ErstelleOrdner fehlgeschlagen: " & strTeil & " - " & Err.Description
                    Err.Clear
                    Exit Sub
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Sub


' ---------------------------------------------------------------------------
' EINDEUTIGEN DATEIPFAD GENERIEREN (vermeidet Ueberschreiben)
' ---------------------------------------------------------------------------
Public Function EindeutigerDateipfad(ByVal strOrdner As String, _
                                      ByVal strDateiname As String, _
                                      ByVal strEndung As String) As String
    Dim objFSO  As Object
    Dim strPfad As String
    Dim i       As Long

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strDateiname = EntferneEndung(BereinigeDateiname(strDateiname))
    If strEndung = "" Then strEndung = "bin"
    strPfad = strOrdner & strDateiname & "." & strEndung

    i = 1
    Do While objFSO.FileExists(strPfad)
        strPfad = strOrdner & strDateiname & " (" & i & ")." & strEndung
        i = i + 1
    Loop

    EindeutigerDateipfad = strPfad
    Set objFSO = Nothing
End Function


' ---------------------------------------------------------------------------
' SQL-SICHERE ZEICHENKETTE (Einfache Anfuehrungszeichen escapen)
' ---------------------------------------------------------------------------
Public Function SQLSafe(ByVal strInput As String) As String
    SQLSafe = Replace(Nz(strInput, ""), "'", "''")
End Function


' ---------------------------------------------------------------------------
' SQL-DATUM: Locale-sicheres Datum fuer Jet/ACE SQL-Statements
' ---------------------------------------------------------------------------
' PROBLEM: Format(Now, "mm/dd/yyyy") ersetzt "/" durch den lokalen
'   Datumstrenner (DE: "." statt "/"). Jet erwartet aber US-Format.
'   => #03.05.2026# ist UNGUELTIG, #03/05/2026# ist korrekt.
'
' LOESUNG: Diese Funktionen erzeugen IMMER locale-sichere SQL-Datum-Literals.
'   Format$ mit escaped Separatoren (\/) erzwingt literale Schrägstriche.
'
' VERWENDUNG:
'   db.Execute "UPDATE t SET Datum=" & SQLDatum(Now) & " WHERE ID=1"
'   db.Execute "INSERT INTO t (Datum) VALUES (" & SQLJetzt() & ")"
'
' ACHTUNG: NIEMALS direkt Format(datum, "mm/dd/yyyy") in SQL verwenden!
'          IMMER SQLDatum() oder SQLJetzt() nutzen!
' ---------------------------------------------------------------------------

' Gibt ein Date als Jet-SQL-Literal zurueck: #mm/dd/yyyy hh:nn:ss#
Public Function SQLDatum(ByVal dtWert As Date) As String
    SQLDatum = "#" & Format$(dtWert, "mm\/dd\/yyyy hh\:nn\:ss") & "#"
End Function

' Kurzform: Aktueller Zeitpunkt als SQL-Literal
Public Function SQLJetzt() As String
    SQLJetzt = SQLDatum(Now)
End Function

' Nur Datum ohne Uhrzeit: #mm/dd/yyyy#
Public Function SQLNurDatum(ByVal dtWert As Date) As String
    SQLNurDatum = "#" & Format$(dtWert, "mm\/dd\/yyyy") & "#"
End Function


' ===========================================================================
' CAPITALIZE / NAME-BEREINIGUNG (v0.2)
' ===========================================================================

' ---------------------------------------------------------------------------
' WORT CAPITALISIEREN mit Sonder-Handling fuer Adelspraedikate
' "von", "van", "de", "der", "den" bleiben klein
' Unterstuetzt Bindestriche: "Mueller-Schmidt" -> "Mueller-Schmidt"
' ---------------------------------------------------------------------------
Public Function CapitalizeWord(ByVal strInput As String) As String
    Dim partsSpace() As String, partsHyphen() As String
    Dim i As Long, j As Long

    strInput = Trim(strInput)
    If Len(strInput) = 0 Then CapitalizeWord = "": Exit Function

    partsSpace = Split(strInput, " ")
    For i = LBound(partsSpace) To UBound(partsSpace)
        partsHyphen = Split(partsSpace(i), "-")
        For j = LBound(partsHyphen) To UBound(partsHyphen)
            If Len(partsHyphen(j)) > 0 Then
                Select Case LCase(partsHyphen(j))


                
                    Case "von", "van", "de", "der", "den"
                        partsHyphen(j) = LCase(partsHyphen(j))
                    Case Else
                        partsHyphen(j) = UCase(Left(partsHyphen(j), 1)) & LCase(Mid(partsHyphen(j), 2))
                End Select
            End If
        Next j
        partsSpace(i) = Join(partsHyphen, "-")
    Next i

    CapitalizeWord = Join(partsSpace, " ")
End Function


' ---------------------------------------------------------------------------
' PRUEFEN OB NUR ALPHABETISCHE ZEICHEN + Umlaute + Bindestrich
' ---------------------------------------------------------------------------
Public Function IsAlphaOnly(ByVal strValue As String) As Boolean
    Dim objRE As Object
    Set objRE = CreateObject("VBScript.RegExp")
    objRE.Pattern = "^[a-zA-ZaeoeueAeOeUess\-]+$"
    objRE.IgnoreCase = True
    objRE.Global = False
    IsAlphaOnly = objRE.Test(Trim(strValue))
    Set objRE = Nothing
End Function


' ---------------------------------------------------------------------------
' DOMAIN AUS E-MAIL EXTRAHIEREN (z.B. "user@rps.bwl.de" -> "rps.bwl.de")
' ---------------------------------------------------------------------------
Public Function ExtrahiereDomain(ByVal strEmail As String) As String
    If InStr(strEmail, "@") > 0 Then
        ExtrahiereDomain = LCase(Split(strEmail, "@")(1))
    Else
        ExtrahiereDomain = ""
    End If
End Function


' ---------------------------------------------------------------------------
' INSTITUTION AUS E-MAIL-DOMAIN ABLEITEN
' Erkennt bekannte Behoerden/Institutionen; sonst Domain-Basis capitaliziert
' ---------------------------------------------------------------------------
Public Function ExtrahiereInstitution(ByVal strEmail As String) As String
    On Error Resume Next

    Dim strDomain As String
    strDomain = ExtrahiereDomain(strEmail)
    If strDomain = "" Then ExtrahiereInstitution = "": Exit Function

    ' TLD entfernen (.de, .com, .org etc.)
    Dim tlds As Variant, tld As Variant
    tlds = Array(".de", ".org", ".com", ".net", ".eu", ".gov", ".bund")
    For Each tld In tlds
        If Right(strDomain, Len(tld)) = CStr(tld) Then
            strDomain = Left(strDomain, Len(strDomain) - Len(CStr(tld)))
            Exit For
        End If
    Next tld

    ' Bekannte Institutionen mappen
    Select Case True
        Case InStr(strDomain, "rp-karlsruhe") > 0:  ExtrahiereInstitution = "RPK": Exit Function
        Case InStr(strDomain, "rp-freiburg") > 0:   ExtrahiereInstitution = "RPF": Exit Function
        Case InStr(strDomain, "rp-stuttgart") > 0:   ExtrahiereInstitution = "RPS": Exit Function
        Case InStr(strDomain, "rps") > 0:            ExtrahiereInstitution = "RPS": Exit Function
        Case InStr(strDomain, "lubw") > 0:           ExtrahiereInstitution = "LUBW": Exit Function
        Case InStr(strDomain, "lgl") > 0:            ExtrahiereInstitution = "LGL": Exit Function
        Case InStr(strDomain, "rpt") > 0:            ExtrahiereInstitution = "RPT": Exit Function
    End Select

    ' Fallback: erster Teil der Domain capitaliziert
    Dim strBase As String
    If InStr(strDomain, ".") > 0 Then
        strBase = Split(strDomain, ".")(0)
    Else
        strBase = strDomain
    End If

    ExtrahiereInstitution = CapitalizeWord(strBase)
End Function


' ---------------------------------------------------------------------------
' NAMEN-EINGABE BEREINIGEN (Sonderzeichen, Doppel-Leerzeichen, Trim)
' ---------------------------------------------------------------------------
Public Function BereinigeNameInput(ByVal strName As String) As String
    strName = Replace(strName, "|", "")
    strName = Replace(strName, ":", "")
    strName = Replace(strName, "/", "")
    strName = Replace(strName, ";", "")
    strName = Replace(strName, Chr(34), "")    ' Anfuehrungszeichen
    strName = Replace(strName, Chr(160), "")   ' Nicht-druckbares Leerzeichen
    strName = Replace(strName, "  ", " ")
    BereinigeNameInput = Trim(strName)
End Function


' ---------------------------------------------------------------------------
' SORTIERNAME BILDEN ("Nachname, Vorname")
' ---------------------------------------------------------------------------
Public Function BildeSortiername(ByVal strNachname As String, ByVal strVorname As String) As String
    Dim s As String
    s = Trim(strNachname)
    If Len(Trim(strVorname)) > 0 Then s = s & ", " & Trim(strVorname)
    BildeSortiername = s
End Function


' ===========================================================================
' ABSENDER KUERZEN (v0.4 - aus modFileManager zentralisiert)
' ===========================================================================

' Absender-Name auf Kurzform bringen:
'   "torsten.kugler@rps.bwl.de" -> "torsten kugler"
'   "Dr. Max Mueller" -> "Max Mueller"
' Fuer Dateinamen und Ordnernamen geeignet.
Public Function KuerzeAbsender(ByVal strAbsender As String, _
                                Optional ByVal intMaxLen As Integer = 20) As String
    If Nz(strAbsender, "") = "" Then
        KuerzeAbsender = ""
        Exit Function
    End If

    Dim s As String
    s = Trim(strAbsender)

    ' Wenn Email-Adresse: Teil vor @
    Dim lngAt As Long
    lngAt = InStr(s, "@")
    If lngAt > 1 Then
        s = Left(s, lngAt - 1)
        s = Replace(s, ".", " ")
    End If

    ' Nur erstes+zweites Wort behalten
    Dim arrTeile As Variant
    arrTeile = Split(Trim(s), " ")
    If UBound(arrTeile) >= 1 Then
        s = arrTeile(0) & " " & arrTeile(1)
    ElseIf UBound(arrTeile) >= 0 Then
        s = arrTeile(0)
    End If

    KuerzeAbsender = Left(s, intMaxLen)
End Function


' ===========================================================================
' RTRIM SPECIAL (v0.4 - Trailing Punkte/Leerzeichen/Bindestriche)
' ===========================================================================

' Windows verbietet Ordner/Dateien mit "." oder " " am Ende.
' Entfernt diese + Bindestriche, Fallback auf strDefault wenn leer.
Public Function RTrimSpecial(ByVal s As String, _
                              Optional ByVal strDefault As String = "Mail") As String
    Do While Len(s) > 0
        Dim c As String
        c = Right(s, 1)
        If c = " " Or c = "." Or c = "-" Then
            s = Left(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    If s = "" Then s = strDefault
    RTrimSpecial = s
End Function


' ===========================================================================
' MAIL-DATEINAME GENERIEREN (v0.4 - zentralisiert)
' ===========================================================================

' Erzeugt einen bereinigten Dateinamen im Format:
'   YYYYMMDD_HHNN_Absender_Betreff.ext
'
' Parameter:
'   strAbsender  - Absendername oder -email
'   strBetreff   - Betreff der Mail (wird automatisch bereinigt)
'   dtZeitpunkt  - Datum/Uhrzeit (EmpfangenAm oder GesendetAm)
'   strEndung    - Dateiendung ohne Punkt (z.B. "msg")
'   intMaxLen    - Maximale Gesamtlaenge inkl. Extension (Default 100)
'
' Rueckgabe: Bereinigter Dateiname mit Extension
Public Function GeneriereMailDateiname(ByVal strAbsender As String, _
                                       ByVal strBetreff As String, _
                                       ByVal dtZeitpunkt As Date, _
                                       ByVal strEndung As String, _
                                       Optional ByVal intMaxLen As Integer = 100) As String
    Dim strName As String
    strName = Format(dtZeitpunkt, "yyyymmdd\_hhnn")

    ' Absender kuerzen
    Dim strAbsKurz As String
    strAbsKurz = BereinigeDateiname(KuerzeAbsender(strAbsender), 20)
    If strAbsKurz <> "" And strAbsKurz <> "Unbenannt" Then
        strName = strName & "_" & strAbsKurz
    End If

    ' Betreff bereinigen und anhaengen
    Dim strBetrKurz As String
    strBetrKurz = BereinigeDateiname(BereinigeBetreff(strBetreff), 40)
    If strBetrKurz <> "" And strBetrKurz <> "Unbenannt" And strBetrKurz <> "(Kein Betreff)" Then
        strName = strName & "_" & strBetrKurz
    End If

    ' Extension anhaengen
    If strEndung = "" Then strEndung = "bin"
    strEndung = LCase(Replace(strEndung, ".", ""))

    ' Gesamtlaenge pruefen (Reserve fuer ".ext")
    Dim intMaxBase As Integer
    intMaxBase = intMaxLen - Len(strEndung) - 1
    If Len(strName) > intMaxBase Then
        strName = Left(strName, intMaxBase)
    End If
    strName = RTrimSpecial(strName, "Mail")

    GeneriereMailDateiname = strName & "." & strEndung
End Function


' ===========================================================================
' ANHANG-DATEINAME BEREINIGEN (v0.4 - zentralisiert)
' ===========================================================================

' Bereinigt einen Anhang-Dateinamen und kuerzt bei Bedarf.
' Erhaelt die Extension intakt.
'
' Parameter:
'   strOriginalName - Originaler Dateiname des Anhangs
'   intMaxLen       - Maximale Gesamtlaenge (Default 100)
Public Function BereinigeAnhangName(ByVal strOriginalName As String, _
                                     Optional ByVal intMaxLen As Integer = 100) As String
    Dim strClean As String
    strClean = BereinigeDateiname(strOriginalName, intMaxLen + 20)

    Dim strExt As String
    strExt = HoleEndung(strClean)

    Dim strBase As String
    strBase = EntferneEndung(strClean)

    ' Kuerzen unter Beibehaltung der Extension
    If strExt <> "" Then
        Dim intMaxBase As Integer
        intMaxBase = intMaxLen - Len(strExt) - 1
        If intMaxBase < 5 Then intMaxBase = 5
        If Len(strBase) > intMaxBase Then
            strBase = Left(strBase, intMaxBase)
        End If
        strBase = RTrimSpecial(strBase, "Anhang")
        BereinigeAnhangName = strBase & "." & strExt
    Else
        If Len(strClean) > intMaxLen Then
            strClean = Left(strClean, intMaxLen)
        End If
        BereinigeAnhangName = RTrimSpecial(strClean, "Anhang")
    End If
End Function


' ===========================================================================
' ABLAGE-BASISPFAD BAUEN (v0.4 - zentralisiert)
' ===========================================================================

' Baut den Basispfad fuer die Mail-Ablage:
'   <Basis>\<Projekt>\<Phase>\YYYY\MM\
'
' Parameter:
'   strBasis    - Export-Basispfad (z.B. "\\Server\Share\Mails\")
'   strProjekt  - Projektname (wird bereinigt)
'   strPhase    - Phasenname (wird bereinigt)
'   dtDatum     - Datum fuer Jahr/Monat-Ordner
'
' Rueckgabe: Pfad mit abschliessendem Backslash
Public Function BaueAblagePfad(ByVal strBasis As String, _
                                ByVal strProjekt As String, _
                                ByVal strPhase As String, _
                                ByVal dtDatum As Date) As String

    BaueAblagePfad = NormalisierePfad(strBasis) & _
                     BereinigeDateiname(strProjekt, 30) & "\" & _
                     BereinigeDateiname(strPhase, 30) & "\" & _
                     Format(dtDatum, "yyyy") & "\" & _
                     Format(dtDatum, "mm") & "\"
End Function


' ===========================================================================
' PFAD SICHER KUERZEN (v0.4 - MAX_PATH-Schutz)
' ===========================================================================

' Konstante: Windows MAX_PATH = 260, minus 12 Reserve fuer 8.3-Dateinamen
Private Const MAX_PFAD_LAENGE As Long = 248

' Kuerzt einen Pfad auf MAX_PATH-sichere Laenge.
' Kuerzt den letzten Pfadbestandteil (Ordner/Dateiname).
' Gibt gekuerzten Pfad zurueck, oder Original wenn kurz genug.
Public Function KuerzePfadSicher(ByVal strPfad As String, _
                                  Optional ByVal lngMaxLen As Long = 0) As String
    If lngMaxLen <= 0 Then lngMaxLen = MAX_PFAD_LAENGE

    If Len(strPfad) <= lngMaxLen Then
        KuerzePfadSicher = strPfad
        Exit Function
    End If

    ' Letzten Bestandteil finden und kuerzen
    Dim lngUeber As Long
    lngUeber = Len(strPfad) - lngMaxLen

    Dim lngLastSep As Long
    lngLastSep = InStrRev(strPfad, "\")
    If lngLastSep > 0 And lngLastSep < Len(strPfad) Then
        Dim strBasis As String
        strBasis = Left(strPfad, lngLastSep)
        Dim strLetzer As String
        strLetzer = Mid(strPfad, lngLastSep + 1)

        If Len(strLetzer) > lngUeber + 5 Then
            strLetzer = Left(strLetzer, Len(strLetzer) - lngUeber - 5)
            strLetzer = RTrimSpecial(strLetzer, "X")
        End If

        KuerzePfadSicher = strBasis & strLetzer
    Else
        ' Kein Separator: einfach abschneiden
        KuerzePfadSicher = Left(strPfad, lngMaxLen)
    End If
End Function


' ===========================================================================
' SAFE TEXT - Null-sicheres Kuerzen (v0.4.1)
' ===========================================================================

' Universelle Abkuerzung fuer das ueberall vorkommende Pattern:
'   Left(Nz(varWert, ""), intMaxLen)
'
' Parameter:
'   varWert     - Variant (kann Null/Nothing/String sein)
'   intMaxLen   - Maximale Laenge (Default 255, typisch fuer TEXT-Felder)
'   strDefault  - Fallback wenn Null/leer (Default "")
'
' Beispiel: SafeText(objMail.Subject, 255) statt Left(Nz(objMail.Subject, ""), 255)
Public Function SafeText(ByVal varWert As Variant, _
                          Optional ByVal intMaxLen As Integer = 255, _
                          Optional ByVal strDefault As String = "") As String
    Dim s As String
    s = Nz(varWert, strDefault)
    If Len(s) > intMaxLen Then s = Left(s, intMaxLen)
    SafeText = s
End Function


' ===========================================================================
' OUTLOOK-PFAD BEREINIGEN (v0.4.1)
' ===========================================================================

' Bereinigt einen Outlook-MAPI-Ordnerpfad:
'   - Fuehrende "\\" entfernen (Namespace-Prefix)
'   - Doppelte Backslashes normalisieren
'   - URL-kodierte Slashes dekodieren (%2F)
'
' Wird von OeffneOrdner (modOutlookConnect) verwendet.
Public Function BereinigeOutlookPfad(ByVal strPfad As String) As String
    If Len(strPfad) = 0 Then
        BereinigeOutlookPfad = ""
        Exit Function
    End If

    ' Fuehrende "\\" vom MAPI-Namespace entfernen
    If Left(strPfad, 2) = "\\" Then strPfad = Mid(strPfad, 3)

    ' URL-kodierte Slashes (aus Webmail/EWS)
    strPfad = Replace(strPfad, "%2F", "/")

    ' Doppelte Backslashes normalisieren
    strPfad = Replace(strPfad, "\\", "\")

    BereinigeOutlookPfad = strPfad
End Function


