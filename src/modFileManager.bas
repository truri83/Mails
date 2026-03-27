Attribute VB_Name = "modFileManager"
Option Compare Database
Option Explicit

' ===========================================================================
' modFileManager - Dateiablage + Netzwerk-Resilienz
' ===========================================================================
' v0.4: Verwaltet die Ablagestruktur fuer .msg-Dateien und Anhaenge
' auf dem Netzlaufwerk. Inkl. Netzwerk-Pruefung, Retry-Logik und
' Wiederaufnahme nach VPN-Verlust.
'
' ABLAGESTRUKTUR:
'   <ExportBasis>\<Projekt>\<Phase>\
'     YYYY\
'       MM\
'         YYYYMMDD_HHNN_Absender_Betreff\
'           Mail.msg
'           Anhaenge\
'             Dateiname.ext
'
' Pfadlaenge: Max 260 Zeichen (Windows-Limit)
'   ExportBasis ~80 + Projekt ~20 + Phase ~20 + YYYY\MM\ = 8
'   + Ordnername max 80 + Dateiname max 50 = ~258
'
' Oeffentliche Routinen:
'   BaueMailOrdnerPfad  - Erstellt Pfad fuer eine Mail
'   BaueAnhangPfad      - Erstellt Pfad fuer einen Anhang
'   IstNetzwerkOK       - Prueft Netzwerk-Erreichbarkeit
'   WarteAufNetzwerk    - Wartet bis Netzwerk wieder da (mit Timeout)
'   KopiereNachNetzwerk - Kopiert Datei mit Retry + Backoff
'   ErstelleAblagePfad  - Baut den vollstaendigen Zielpfad
'
' Abhaengigkeiten: modStringUtils, modLogging
' ===========================================================================


' ---------------------------------------------------------------------------
' KONSTANTEN
' ---------------------------------------------------------------------------
Private Const MAX_ORDNERNAME     As Long = 80    ' Max Laenge Mail-Ordnername
Private Const MAX_DATEINAME      As Long = 50    ' Max Laenge Dateiname
Private Const MAX_PFAD           As Long = 248   ' Windows MAX_PATH - 12 Reserve
Private Const NETZWERK_TIMEOUT   As Long = 300   ' Sekunden Warten auf Netzwerk
Private Const NETZWERK_POLL      As Long = 5000  ' ms zwischen Netzwerk-Pruefungen

' Win32 API fuer GetTickCount
#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If


' ===========================================================================
' ABLAGESTRUKTUR: Mail-Ordner bauen
' ===========================================================================

' Gibt den vollstaendigen Ordnerpfad fuer eine Mail zurueck:
'   <Basis>\<Projekt>\<Phase>\YYYY\MM\YYYYMMDD_HHNN_Absender_Betreff\
'
' Parameter:
'   strBasis     - Export-Basispfad (z.B. "\\Server\Share\Mails\")
'   strProjekt   - Projektname
'   strPhase     - Phasenname
'   dtEmpfangen  - Empfangsdatum der Mail
'   strAbsender  - Absendername oder -email
'   strBetreff   - Betreff der Mail
'
' Rueckgabe: Ordnerpfad MIT abschliessendem Backslash
Public Function BaueMailOrdnerPfad(ByVal strBasis As String, _
                                    ByVal strProjekt As String, _
                                    ByVal strPhase As String, _
                                    ByVal dtEmpfangen As Date, _
                                    ByVal strAbsender As String, _
                                    ByVal strBetreff As String) As String

    ' Basis + Projekt + Phase
    Dim strPfad As String
    strPfad = NormalisierePfad(strBasis) & _
              BereinigeDateiname(strProjekt, 30) & "\" & _
              BereinigeDateiname(strPhase, 30) & "\"

    ' Jahr + Monat
    strPfad = strPfad & Format(dtEmpfangen, "yyyy") & "\" & _
              Format(dtEmpfangen, "mm") & "\"

    ' Mail-Ordnername: YYYYMMDD_HHNN_Absender_Betreff
    Dim strOrdner As String
    strOrdner = Format(dtEmpfangen, "yyyymmdd\_hhnn")

    ' Absender kuerzen (nur Teil vor @, oder erstes Wort)
    Dim strAbsKurz As String
    strAbsKurz = KuerzeAbsender(strAbsender)
    If strAbsKurz <> "" Then
        strOrdner = strOrdner & "_" & BereinigeDateiname(strAbsKurz, 20)
    End If

    ' Betreff kuerzen
    Dim strBetrKurz As String
    strBetrKurz = BereinigeDateiname(BereinigeBetreff(strBetreff), 40)
    If strBetrKurz <> "" And strBetrKurz <> "Unbenannt" Then
        strOrdner = strOrdner & "_" & strBetrKurz
    End If

    ' Gesamtlaenge pruefen
    If Len(strOrdner) > MAX_ORDNERNAME Then
        strOrdner = Left(strOrdner, MAX_ORDNERNAME)
    End If
    ' Trailing Leerzeichen/Punkte entfernen (Windows verbietet das)
    strOrdner = RTrimSpecial(strOrdner)

    strPfad = strPfad & strOrdner & "\"

    ' MAX_PATH pruefen
    If Len(strPfad) > MAX_PFAD Then
        ' Ordnername weiter kuerzen
        Dim lngUeber As Long
        lngUeber = Len(strPfad) - MAX_PFAD
        strOrdner = Left(strOrdner, Len(strOrdner) - lngUeber - 5)
        strOrdner = RTrimSpecial(strOrdner)
        strPfad = NormalisierePfad(strBasis) & _
                  BereinigeDateiname(strProjekt, 30) & "\" & _
                  BereinigeDateiname(strPhase, 30) & "\" & _
                  Format(dtEmpfangen, "yyyy") & "\" & _
                  Format(dtEmpfangen, "mm") & "\" & _
                  strOrdner & "\"
    End If

    BaueMailOrdnerPfad = strPfad
End Function


' ===========================================================================
' ABLAGESTRUKTUR: Anhang-Pfad bauen
' ===========================================================================

' Gibt den Dateipfad fuer einen Anhang zurueck:
'   <MailOrdner>\Anhaenge\Dateiname.ext
'
' Verwendet EindeutigerDateipfad um Ueberschreiben zu vermeiden.
Public Function BaueAnhangPfad(ByVal strMailOrdner As String, _
                                ByVal strDateiname As String) As String

    Dim strAnhOrdner As String
    strAnhOrdner = NormalisierePfad(strMailOrdner) & "Anhaenge\"

    Dim strEndung As String
    strEndung = HoleEndung(strDateiname)
    If strEndung = "" Then strEndung = "bin"

    Dim strName As String
    strName = BereinigeDateiname(EntferneEndung(strDateiname), MAX_DATEINAME)

    BaueAnhangPfad = EindeutigerDateipfad(strAnhOrdner, strName, strEndung)
End Function


' ===========================================================================
' ABLAGESTRUKTUR: MSG-Pfad bauen
' ===========================================================================

' Gibt den Dateipfad fuer die .msg-Datei zurueck:
'   <MailOrdner>\Mail.msg
Public Function BaueMSGPfad(ByVal strMailOrdner As String) As String
    BaueMSGPfad = EindeutigerDateipfad(NormalisierePfad(strMailOrdner), "Mail", "msg")
End Function


' ===========================================================================
' NETZWERK-PRUEFUNG
' ===========================================================================

' Prueft ob ein Netzwerk-Pfad erreichbar ist (schneller Dir()-Test)
Public Function IstNetzwerkOK(ByVal strNetzwerkPfad As String) As Boolean
    On Error Resume Next

    Dim strTest As String
    strTest = Dir(NormalisierePfad(strNetzwerkPfad), vbDirectory)

    IstNetzwerkOK = (Err.Number = 0 And strTest <> "")
    Err.Clear
    On Error GoTo 0
End Function


' Wartet bis Netzwerk wieder erreichbar ist (nach VPN-Verlust etc.)
' Gibt True zurueck wenn Netzwerk innerhalb Timeout zurueckgekehrt.
' Gibt False zurueck wenn Timeout ueberschritten oder Abbruch.
'
' Parameter:
'   strNetzwerkPfad - UNC-Pfad zum Pruefen
'   lngTimeoutSek   - Max. Wartezeit in Sekunden (0 = Default NETZWERK_TIMEOUT)
Public Function WarteAufNetzwerk(ByVal strNetzwerkPfad As String, _
                                  Optional ByVal lngTimeoutSek As Long = 0) As Boolean
    If lngTimeoutSek <= 0 Then lngTimeoutSek = NETZWERK_TIMEOUT

    If IstNetzwerkOK(strNetzwerkPfad) Then
        WarteAufNetzwerk = True
        Exit Function
    End If

    LogWarn "Netzwerk nicht erreichbar: " & strNetzwerkPfad & _
            " - Warte bis zu " & lngTimeoutSek & "s...", "FILEMANAGER"

    Dim lngStart As Long
    lngStart = GetTickCount()
    Dim lngEnde As Long
    lngEnde = lngStart + (lngTimeoutSek * 1000)

    Do
        Sleep NETZWERK_POLL
        DoEvents

        ' Abbruch-Pruefung
        If g_blnAbbrechen Then
            LogWarn "Netzwerk-Warten abgebrochen durch Benutzer", "FILEMANAGER"
            WarteAufNetzwerk = False
            Exit Function
        End If

        If IstNetzwerkOK(strNetzwerkPfad) Then
            Dim lngGewartet As Long
            lngGewartet = (GetTickCount() - lngStart) \ 1000
            LogInfo "Netzwerk wieder da nach " & lngGewartet & "s", "FILEMANAGER"
            WarteAufNetzwerk = True
            Exit Function
        End If
    Loop While GetTickCount() < lngEnde

    LogError "Netzwerk-Timeout nach " & lngTimeoutSek & "s: " & strNetzwerkPfad, "FILEMANAGER"
    WarteAufNetzwerk = False
End Function


' ===========================================================================
' DATEI-KOPIE MIT NETZWERK-RESILIENZ
' ===========================================================================

' Kopiert eine Datei von Quelle nach Ziel mit:
'   - Zielverzeichnis automatisch erstellen
'   - Retry bei Netzwerk-Fehler + exponentielles Backoff
'   - Bei Netzwerk-Verlust: WarteAufNetzwerk() aufrufen
'   - Abbruch-faehig (g_blnAbbrechen)
'
' Gibt True zurueck bei Erfolg.
Public Function KopiereNachNetzwerk(ByVal strQuelle As String, _
                                     ByVal strZiel As String, _
                                     Optional ByVal intMaxVersuche As Integer = 3, _
                                     Optional ByVal lngBasisPause As Long = 2000) As Boolean
    On Error GoTo ErrHandler

    ' Quelldatei pruefen
    If Dir(strQuelle) = "" Then
        LogWarn "Quelldatei nicht gefunden: " & strQuelle, "FILEMANAGER"
        KopiereNachNetzwerk = False
        Exit Function
    End If

    ' Zielverzeichnis
    Dim strZielDir As String
    strZielDir = Left(strZiel, InStrRev(strZiel, "\"))
    If strZielDir <> "" Then ErstelleOrdner strZielDir

    ' Retry-Schleife
    Dim intVersuch As Integer
    Dim lngPause As Long
    lngPause = lngBasisPause

    For intVersuch = 1 To intMaxVersuche
        If g_blnAbbrechen Then
            KopiereNachNetzwerk = False
            Exit Function
        End If

        On Error Resume Next
        FileCopy strQuelle, strZiel

        If Err.Number = 0 Then
            On Error GoTo 0
            LogTrace "Datei kopiert: " & strZiel, "FILEMANAGER"
            KopiereNachNetzwerk = True
            Exit Function
        End If

        ' Fehler auswerten
        Dim lngErr As Long: lngErr = Err.Number
        Dim strErr As String: strErr = Err.Description
        Err.Clear
        On Error GoTo ErrHandler

        ' Netzwerk-Fehler? (typische Codes: 52, 53, 70, 75, 76)
        If IstNetzwerkFehler(lngErr) Then
            LogWarn "Netzwerk-Fehler bei Kopie (Versuch " & intVersuch & "): " & strErr, "FILEMANAGER"

            ' Warten bis Netzwerk wieder da
            If Not WarteAufNetzwerk(strZielDir) Then
                LogError "Netzwerk nicht wiederhergestellt - Kopie abgebrochen: " & strQuelle, "FILEMANAGER"
                KopiereNachNetzwerk = False
                Exit Function
            End If

            ' Zielverzeichnis ggf. nochmal erstellen
            ErstelleOrdner strZielDir
        Else
            ' Sonstiger Fehler: Retry mit Backoff
            If intVersuch < intMaxVersuche Then
                LogWarn "Datei-Kopie Versuch " & intVersuch & "/" & intMaxVersuche & _
                        ": " & strErr & " - Retry in " & lngPause & "ms", "FILEMANAGER"
                Sleep lngPause
                lngPause = lngPause * 2
                If lngPause > 16000 Then lngPause = 16000
            End If
        End If
    Next intVersuch

    LogError "Datei-Kopie endgueltig fehlgeschlagen nach " & intMaxVersuche & _
             " Versuchen: " & strQuelle, "FILEMANAGER"
    KopiereNachNetzwerk = False
    Exit Function

ErrHandler:
    LogVBAError "KopiereNachNetzwerk"
    KopiereNachNetzwerk = False
End Function


' ===========================================================================
' HILFSFUNKTIONEN (privat)
' ===========================================================================

' KuerzeAbsender und RTrimSpecial -> zentralisiert in modStringUtils (v0.4)

' Prueft ob ein VBA-Fehlercode ein Netzwerk-Problem anzeigt
Private Function IstNetzwerkFehler(ByVal lngErrNum As Long) As Boolean
    Select Case lngErrNum
        Case 52     ' Bad file name or number
        Case 53     ' File not found
        Case 57     ' Device I/O error
        Case 67     ' Too many files
        Case 70     ' Permission denied
        Case 71     ' Disk not ready
        Case 75     ' Path/File access error
        Case 76     ' Path not found
        Case Else
            IstNetzwerkFehler = False
            Exit Function
    End Select
    IstNetzwerkFehler = True
End Function
