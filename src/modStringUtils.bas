Attribute VB_Name = "modStringUtils"
Option Compare Database
Option Explicit

' ===========================================================================
' modStringUtils - String-Bereinigung, Validierung und Pfad-Utilities
' ===========================================================================
' Konsolidiert alle String-Operationen:
'   - E-Mail-Validierung (RegExp)
'   - Betreff-Bereinigung (RE:/AW:/FW:/WG:/EXTERN entfernen)
'   - Dateinamen-Bereinigung (illegale Zeichen)
'   - Kontaktinfo-Bereinigung (Name + Email)
'   - Pfad-Normalisierung und Ordner-Erstellung
'   - Eindeutige Dateipfad-Generierung
'   - SQL-Escaping
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

    objRE.Pattern = "^(?:(RE|AW|FW|FWD|WG|EXTERN)(\s*[::])\s*)+"
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
        NormalisierePfad = Environ("USERPROFILE") & "\OutlookSync\"
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
