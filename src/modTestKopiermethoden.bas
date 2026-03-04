Attribute VB_Name = "modTestKopiermethoden"
Option Compare Database
Option Explicit

' ===========================================================================
' modTestKopiermethoden - Testmodul: Welche Hintergrund-Kopiermethoden gehen?
' ===========================================================================
' Testet systematisch alle realistischen Optionen fuer asynchrones/
' nicht-blockierendes Datei-Kopieren unter folgenden Einschraenkungen:
'   - Kein Admin-Zugriff
'   - PowerShell-Scripting ggf. blockiert (Execution Policy)
'   - Access VBA 64-Bit Umgebung
'   - Netzlaufwerke als Ziel
'
' AUFRUF IM DIREKTFENSTER:
'   TestAlleKopiermethoden            ' Alle Tests auf einmal
'   TestAlleKopiermethoden "Z:\"     ' Mit eigenem Netzlaufwerk-Pfad
'
' EINZELTESTS:
'   TestShellCmdCopy                  ' cmd /c copy (async via Shell)
'   TestShellRobocopy                 ' robocopy (Windows built-in)
'   TestShellXcopy                    ' xcopy (Windows built-in)
'   TestShellStartCopy                ' cmd /c start /b copy (detached)
'   TestWScriptShellRun               ' WScript.Shell.Run (COM, async)
'   TestCopyFileExAPI                 ' CopyFileEx mit Callback (WinAPI)
'   TestTimerBasiert                  ' Access Form-Timer (chunk-weise)
'   TestBITSAdmin                     ' BITS Service (bitsadmin.exe)
'   TestPowerShellCopy                ' PowerShell Start-Job (wahrscheinl. blockiert)
'   TestFileCopyBaseline              ' VBA FileCopy - Baseline (synchron)
'
' Jeder Test:
'   1. Erstellt eine Test-Datei (konfigurierbare Groesse)
'   2. Versucht die Kopiermethode
'   3. Prueft ob Zieldatei angekommen ist
'   4. Misst Zeit + blockierend/nicht-blockierend
'   5. Raeumt auf
'
' Abhaengigkeiten: Keine (komplett eigenstaendig)
' ===========================================================================


' ---------------------------------------------------------------------------
' WINDOWS API DEKLARATIONEN
' ---------------------------------------------------------------------------
#If VBA7 Then
    ' Shell-Prozess-Handle warten (fuer Timing)
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
         ByVal dwProcessId As Long) As LongPtr

    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long

    Private Declare PtrSafe Function CloseHandle Lib "kernel32" _
        (ByVal hObject As LongPtr) As Long

    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As LongPtr, lpExitCode As Long) As Long

    Private Declare PtrSafe Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)

    ' CopyFileEx mit Progress-Callback
    Private Declare PtrSafe Function CopyFileEx Lib "kernel32" Alias "CopyFileExA" _
        (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, _
         ByVal lpProgressRoutine As LongPtr, ByVal lpData As LongPtr, _
         ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long

    ' GetTickCount fuer Zeitmessung
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
         ByVal dwProcessId As Long) As Long

    Private Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

    Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

    Private Declare Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, lpExitCode As Long) As Long

    Private Declare Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)

    Private Declare Function CopyFileEx Lib "kernel32" Alias "CopyFileExA" _
        (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, _
         ByVal lpProgressRoutine As Long, ByVal lpData As Long, _
         ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long

    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

' Konstanten fuer OpenProcess
Private Const PROCESS_ALL_ACCESS    As Long = &H1F0FFF
Private Const SYNCHRONIZE           As Long = &H100000
Private Const WAIT_TIMEOUT          As Long = 258
Private Const STILL_ACTIVE          As Long = 259


' ---------------------------------------------------------------------------
' MODUL-VARIABLEN (Test-Konfiguration)
' ---------------------------------------------------------------------------
Private m_strTempDir        As String   ' Lokaler Temp-Ordner
Private m_strNetzwerkDir    As String   ' Netzwerk-Zielordner (oder lokal als Fallback)
Private m_lngTestDateiMB    As Long     ' Testdatei-Groesse in MB
Private m_blnVerbose        As Boolean  ' Detaillierte Ausgabe


' ===========================================================================
' HAUPT-EINSTIEGSPUNKT: Alle Methoden testen
' ===========================================================================

' Alle Kopiermethoden durchlaufen und Ergebnis-Tabelle ausgeben.
' strNetzwerkPfad: Optionaler Netzlaufwerk-Pfad zum testen. Wenn leer,
'                  wird ein lokaler Ordner als Fallback verwendet.
' lngTestMB:       Groesse der Testdatei in MB (Standard: 5 MB)
Public Sub TestAlleKopiermethoden(Optional ByVal strNetzwerkPfad As String = "", _
                                   Optional ByVal lngTestMB As Long = 5)
    On Error GoTo ErrHandler

    m_blnVerbose = True

    ' Temp-Verzeichnis
    m_strTempDir = Environ("TEMP") & "\KopierTest\"
    If Dir(m_strTempDir, vbDirectory) = "" Then MkDir m_strTempDir

    ' Netzwerk-Pfad (oder Fallback)
    If strNetzwerkPfad <> "" Then
        m_strNetzwerkDir = strNetzwerkPfad
        If Right(m_strNetzwerkDir, 1) <> "\" Then m_strNetzwerkDir = m_strNetzwerkDir & "\"
        m_strNetzwerkDir = m_strNetzwerkDir & "KopierTest\"
    Else
        m_strNetzwerkDir = Environ("TEMP") & "\KopierTestZiel\"
        Debug.Print "*** HINWEIS: Kein Netzlaufwerk angegeben, teste lokal-zu-lokal."
        Debug.Print "    Fuer echten Test:  TestAlleKopiermethoden ""Z:\Pfad\"""
        Debug.Print ""
    End If
    If Dir(m_strNetzwerkDir, vbDirectory) = "" Then MkDir m_strNetzwerkDir

    m_lngTestDateiMB = lngTestMB
    If m_lngTestDateiMB < 1 Then m_lngTestDateiMB = 1
    If m_lngTestDateiMB > 100 Then m_lngTestDateiMB = 100

    ' --- Ueberschrift ---
    Debug.Print String(80, "=")
    Debug.Print "  KOPIERMETHODEN-TEST"
    Debug.Print String(80, "=")
    Debug.Print "  Lokal       : " & m_strTempDir
    Debug.Print "  Ziel        : " & m_strNetzwerkDir
    Debug.Print "  Testdatei   : " & m_lngTestDateiMB & " MB"
    Debug.Print "  Zeitpunkt   : " & Now()
    Debug.Print "  VBA7/64-Bit : " & _
#If VBA7 Then
        "Ja"
#Else
        "Nein"
#End If
    Debug.Print String(80, "=")
    Debug.Print ""

    ' --- Testdatei erstellen ---
    Debug.Print "Erstelle Testdatei (" & m_lngTestDateiMB & " MB)..."
    Dim strTestDatei As String
    strTestDatei = ErstelleTestDatei(m_lngTestDateiMB)
    If strTestDatei = "" Then
        Debug.Print "*** FEHLER: Konnte Testdatei nicht erstellen! Abbruch."
        Exit Sub
    End If
    Debug.Print "  -> " & strTestDatei & " (" & FileLen(strTestDatei) & " Bytes)"
    Debug.Print ""

    ' --- Ergebnis-Array ---
    Dim aErgebnisse(0 To 9, 0 To 4) As String  ' Name, Status, Typ, Zeit, Kommentar

    ' --- Tests durchfuehren ---
    Dim i As Integer

    i = 0: RunTest strTestDatei, "1. VBA FileCopy (Baseline)", _
            "TestFileCopyBaseline", aErgebnisse, i

    i = 1: RunTest strTestDatei, "2. Shell cmd /c copy", _
            "TestShellCmdCopy", aErgebnisse, i

    i = 2: RunTest strTestDatei, "3. Shell robocopy", _
            "TestShellRobocopy", aErgebnisse, i

    i = 3: RunTest strTestDatei, "4. Shell xcopy", _
            "TestShellXcopy", aErgebnisse, i

    i = 4: RunTest strTestDatei, "5. cmd /c start /b copy", _
            "TestShellStartCopy", aErgebnisse, i

    i = 5: RunTest strTestDatei, "6. WScript.Shell.Run", _
            "TestWScriptShellRun", aErgebnisse, i

    i = 6: RunTest strTestDatei, "7. CopyFileEx API", _
            "TestCopyFileExAPI", aErgebnisse, i

    i = 7: RunTest strTestDatei, "8. bitsadmin (BITS)", _
            "TestBITSAdmin", aErgebnisse, i

    i = 8: RunTest strTestDatei, "9. PowerShell Start-Job", _
            "TestPowerShellCopy", aErgebnisse, i

    i = 9: RunTest strTestDatei, "10. Access Timer (chunk)", _
            "TestTimerInfo", aErgebnisse, i

    ' --- ERGEBNIS-TABELLE ---
    Debug.Print ""
    Debug.Print String(80, "=")
    Debug.Print "  ERGEBNIS-UEBERSICHT"
    Debug.Print String(80, "=")
    Debug.Print ""
    Debug.Print PadR("Methode", 30) & PadR("Status", 12) & _
                PadR("Blockiert?", 12) & PadR("Zeit", 10) & "Kommentar"
    Debug.Print String(80, "-")

    For i = 0 To 9
        If aErgebnisse(i, 0) <> "" Then
            Debug.Print PadR(aErgebnisse(i, 0), 30) & _
                        PadR(aErgebnisse(i, 1), 12) & _
                        PadR(aErgebnisse(i, 2), 12) & _
                        PadR(aErgebnisse(i, 3), 10) & _
                        aErgebnisse(i, 4)
        End If
    Next i

    Debug.Print String(80, "-")
    Debug.Print ""

    ' --- BEWERTUNG ausgeben ---
    Debug.Print "LEGENDE:"
    Debug.Print "  OK       = Datei erfolgreich kopiert"
    Debug.Print "  FEHLER   = Methode nicht verfuegbar oder blockiert"
    Debug.Print "  SYNC     = Blockiert Access waehrend der Kopie"
    Debug.Print "  ASYNC    = Gibt sofort zurueck, Kopie laeuft im Hintergrund"
    Debug.Print "  SEMI     = Blockiert kurz, aber mit DoEvents moeglich"
    Debug.Print ""

    Debug.Print "EMPFEHLUNG: Die ASYNC/OK-Methoden sind fuer Hintergrund-Kopie geeignet."
    Debug.Print "            Nutze die schnellste ASYNC-Methode als primaere Strategie,"
    Debug.Print "            VBA FileCopy als Fallback."
    Debug.Print ""

    ' --- Aufraeumen ---
    AufraeumenTestdateien strTestDatei
    Debug.Print "Testdateien bereinigt."
    Debug.Print String(80, "=")
    Exit Sub

ErrHandler:
    Debug.Print "*** FEHLER in TestAlleKopiermethoden: " & Err.Description
End Sub


' ===========================================================================
' EINZELNE TEST-METHODEN
' ===========================================================================

' ---------------------------------------------------------------------------
' TEST 1: VBA FileCopy (synchron, Baseline zum Vergleich)
' ---------------------------------------------------------------------------
Public Sub TestFileCopyBaseline(Optional ByVal strQuelle As String = "", _
                                 Optional ByVal strZiel As String = "")
    On Error GoTo ErrHandler

    InitFallback strQuelle, strZiel, "test_filecopy.dat"

    Debug.Print "  [FileCopy] Start..."
    Dim t1 As Long: t1 = GetTickCount()

    FileCopy strQuelle, strZiel

    Dim t2 As Long: t2 = GetTickCount()
    Dim blnOK As Boolean: blnOK = (Dir(strZiel) <> "")

    Debug.Print "  [FileCopy] " & IIf(blnOK, "OK", "FEHLER") & _
                " - " & (t2 - t1) & " ms (SYNCHRON/blockierend)"

    CleanupZiel strZiel
    Exit Sub
ErrHandler:
    Debug.Print "  [FileCopy] FEHLER: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST 2: Shell() + cmd /c copy (Shell gibt sofort zurueck = async!)
' ---------------------------------------------------------------------------
Public Sub TestShellCmdCopy(Optional ByVal strQuelle As String = "", _
                             Optional ByVal strZiel As String = "")
    On Error GoTo ErrHandler

    InitFallback strQuelle, strZiel, "test_cmdcopy.dat"

    Debug.Print "  [cmd /c copy] Start..."
    Dim t1 As Long: t1 = GetTickCount()

    ' Shell() in VBA gibt sofort die PID zurueck und wartet NICHT
    Dim lngPID As Long
    lngPID = Shell("cmd /c copy """ & strQuelle & """ """ & strZiel & """", vbHide)

    Dim t2 As Long: t2 = GetTickCount()
    Debug.Print "  [cmd /c copy] Shell zurueck nach " & (t2 - t1) & " ms (PID=" & lngPID & ")"

    ' Warten bis Prozess fertig (max 30 Sekunden)
    Dim blnFertig As Boolean
    blnFertig = WarteAufProzess(lngPID, 30000)

    Dim t3 As Long: t3 = GetTickCount()
    Dim blnOK As Boolean: blnOK = (Dir(strZiel) <> "")

    Debug.Print "  [cmd /c copy] " & IIf(blnOK, "OK", "FEHLER") & _
                " - Shell: " & (t2 - t1) & " ms, Gesamt: " & (t3 - t1) & " ms" & _
                IIf(blnFertig, "", " (Timeout!)")
    Debug.Print "  [cmd /c copy] -> ASYNC: Shell kehrt sofort zurueck. Prozess laeuft im Hintergrund."

    CleanupZiel strZiel
    Exit Sub
ErrHandler:
    Debug.Print "  [cmd /c copy] FEHLER: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST 3: Shell() + robocopy (Windows built-in, robust, retry-faehig)
' ---------------------------------------------------------------------------
Public Sub TestShellRobocopy(Optional ByVal strQuelle As String = "", _
                              Optional ByVal strZiel As String = "")
    On Error GoTo ErrHandler

    InitFallback strQuelle, strZiel, "test_robocopy.dat"

    Debug.Print "  [robocopy] Start..."

    ' robocopy braucht Quell-DIR + Ziel-DIR + Dateiname
    Dim strSrcDir As String, strDstDir As String, strDatei As String
    strSrcDir = Left(strQuelle, InStrRev(strQuelle, "\") - 1)
    strDatei = Mid(strQuelle, InStrRev(strQuelle, "\") + 1)
    strDstDir = Left(strZiel, InStrRev(strZiel, "\") - 1)

    ' Erst pruefen ob robocopy existiert
    Dim strRobo As String
    strRobo = Environ("SystemRoot") & "\System32\robocopy.exe"
    If Dir(strRobo) = "" Then
        Debug.Print "  [robocopy] NICHT VERFUEGBAR: " & strRobo & " nicht gefunden"
        Exit Sub
    End If

    Dim t1 As Long: t1 = GetTickCount()

    ' /R:3 /W:2 = 3 Retries, 2 Sek Pause (Netzwerk-Robustheit!)
    ' /NJH /NJS = keine Header/Summary im Output
    Dim lngPID As Long
    lngPID = Shell("cmd /c robocopy """ & strSrcDir & """ """ & strDstDir & _
                    """ """ & strDatei & """ /R:3 /W:2 /NJH /NJS", vbHide)

    Dim t2 As Long: t2 = GetTickCount()
    Debug.Print "  [robocopy] Shell zurueck nach " & (t2 - t1) & " ms (PID=" & lngPID & ")"

    ' Warten bis fertig
    WarteAufProzess lngPID, 30000
    Dim t3 As Long: t3 = GetTickCount()
    Dim blnOK As Boolean: blnOK = (Dir(strZiel) <> "")

    Debug.Print "  [robocopy] " & IIf(blnOK, "OK", "FEHLER") & _
                " - Shell: " & (t2 - t1) & " ms, Gesamt: " & (t3 - t1) & " ms"
    Debug.Print "  [robocopy] -> ASYNC + eingebaute Retry-Logik + Netzwerk-Optimierung"
    If blnOK Then
        Debug.Print "  [robocopy] -> EMPFOHLEN fuer Netzwerk-Kopien (built-in, kein Admin noetig)"
    End If

    CleanupZiel strZiel
    Exit Sub
ErrHandler:
    Debug.Print "  [robocopy] FEHLER: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST 4: Shell() + xcopy
' ---------------------------------------------------------------------------
Public Sub TestShellXcopy(Optional ByVal strQuelle As String = "", _
                           Optional ByVal strZiel As String = "")
    On Error GoTo ErrHandler

    InitFallback strQuelle, strZiel, "test_xcopy.dat"

    Debug.Print "  [xcopy] Start..."
    Dim t1 As Long: t1 = GetTickCount()

    ' /Y = keine Ueberschreib-Frage, /Q = leise
    Dim lngPID As Long
    lngPID = Shell("cmd /c echo F | xcopy """ & strQuelle & """ """ & strZiel & """ /Y /Q", vbHide)

    Dim t2 As Long: t2 = GetTickCount()
    Debug.Print "  [xcopy] Shell zurueck nach " & (t2 - t1) & " ms (PID=" & lngPID & ")"

    WarteAufProzess lngPID, 30000
    Dim t3 As Long: t3 = GetTickCount()
    Dim blnOK As Boolean: blnOK = (Dir(strZiel) <> "")

    Debug.Print "  [xcopy] " & IIf(blnOK, "OK", "FEHLER") & _
                " - Shell: " & (t2 - t1) & " ms, Gesamt: " & (t3 - t1) & " ms"
    Debug.Print "  [xcopy] -> ASYNC via Shell()"

    CleanupZiel strZiel
    Exit Sub
ErrHandler:
    Debug.Print "  [xcopy] FEHLER: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST 5: cmd /c start /b copy (detached process)
' ---------------------------------------------------------------------------
Public Sub TestShellStartCopy(Optional ByVal strQuelle As String = "", _
                               Optional ByVal strZiel As String = "")
    On Error GoTo ErrHandler

    InitFallback strQuelle, strZiel, "test_startcopy.dat"

    Debug.Print "  [start /b copy] Start..."
    Dim t1 As Long: t1 = GetTickCount()

    ' start /b = ohne neues Fenster, laeuft komplett losgeloest
    Dim lngPID As Long
    lngPID = Shell("cmd /c start /b cmd /c copy """ & strQuelle & """ """ & strZiel & """", vbHide)

    Dim t2 As Long: t2 = GetTickCount()
    Debug.Print "  [start /b copy] Shell zurueck nach " & (t2 - t1) & " ms (PID=" & lngPID & ")"

    ' Hier kann es sein dass der tatsaechliche copy-Prozess ein Kind-Prozess ist
    ' -> wir warten einfach und pruefen dann die Datei
    Sleep 3000
    Dim t3 As Long: t3 = GetTickCount()
    Dim blnOK As Boolean: blnOK = (Dir(strZiel) <> "")

    Debug.Print "  [start /b copy] " & IIf(blnOK, "OK", "FEHLER") & _
                " - Shell: " & (t2 - t1) & " ms (nach 3s Pause geprueft)"
    Debug.Print "  [start /b copy] -> ASYNC: Komplett losgeloester Prozess"

    CleanupZiel strZiel
    Exit Sub
ErrHandler:
    Debug.Print "  [start /b copy] FEHLER: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST 6: WScript.Shell.Run (COM-basiert, async-Modus)
' ---------------------------------------------------------------------------
Public Sub TestWScriptShellRun(Optional ByVal strQuelle As String = "", _
                                Optional ByVal strZiel As String = "")
    On Error GoTo ErrHandler

    InitFallback strQuelle, strZiel, "test_wscriptrun.dat"

    Debug.Print "  [WScript.Shell.Run] Start..."

    ' WScript.Shell erstellen via Late Binding
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")

    Dim t1 As Long: t1 = GetTickCount()

    ' .Run mit bWaitOnReturn=False = ASYNC!
    ' Rueckgabewert 0 bei bWaitOnReturn=False
    objShell.Run "cmd /c copy """ & strQuelle & """ """ & strZiel & """", 0, False

    Dim t2 As Long: t2 = GetTickCount()
    Debug.Print "  [WScript.Shell.Run] Zurueck nach " & (t2 - t1) & " ms (ASYNC)"

    ' Warten bis Datei da ist (max 30s)
    Dim blnOK As Boolean
    blnOK = WarteAufDatei(strZiel, 30000)

    Dim t3 As Long: t3 = GetTickCount()
    Debug.Print "  [WScript.Shell.Run] " & IIf(blnOK, "OK", "FEHLER") & _
                " - Run: " & (t2 - t1) & " ms, Datei da: " & (t3 - t1) & " ms"
    Debug.Print "  [WScript.Shell.Run] -> ASYNC + bWaitOnReturn optional"

    Set objShell = Nothing
    CleanupZiel strZiel
    Exit Sub
ErrHandler:
    Debug.Print "  [WScript.Shell.Run] FEHLER: " & Err.Description
    Debug.Print "  [WScript.Shell.Run] -> Moeglicherweise durch Sicherheitsrichtlinie blockiert"
    Set objShell = Nothing
End Sub


' ---------------------------------------------------------------------------
' TEST 7: CopyFileEx Windows API (mit Progress-Callback-Moeglichkeit)
' ---------------------------------------------------------------------------
Public Sub TestCopyFileExAPI(Optional ByVal strQuelle As String = "", _
                              Optional ByVal strZiel As String = "")
    On Error GoTo ErrHandler

    InitFallback strQuelle, strZiel, "test_copyfileex.dat"

    Debug.Print "  [CopyFileEx API] Start..."
    Dim t1 As Long: t1 = GetTickCount()

    ' CopyFileEx: lpProgressRoutine=0 = kein Callback
    ' pbCancel=0 = nicht abbrechen, dwCopyFlags=0 = normal
    Dim lngCancel As Long: lngCancel = 0
    Dim lngResult As Long
    lngResult = CopyFileEx(strQuelle, strZiel, 0, 0, lngCancel, 0)

    Dim t2 As Long: t2 = GetTickCount()
    Dim blnOK As Boolean: blnOK = (lngResult <> 0 And Dir(strZiel) <> "")

    Debug.Print "  [CopyFileEx API] " & IIf(blnOK, "OK", "FEHLER (Ret=" & lngResult & ")") & _
                " - " & (t2 - t1) & " ms"
    Debug.Print "  [CopyFileEx API] -> SYNCHRON, aber: Cancel-Flag + Progress-Callback moeglich"
    Debug.Print "  [CopyFileEx API] -> Vorteil: Abbruch waehrend Kopie moeglich (pbCancel=1)"
    If blnOK Then
        Debug.Print "  [CopyFileEx API] -> Nutzbar in Kombination mit Timer fuer Nicht-Blockierung"
    End If

    CleanupZiel strZiel
    Exit Sub
ErrHandler:
    Debug.Print "  [CopyFileEx API] FEHLER: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST 8: bitsadmin (Background Intelligent Transfer Service)
' ---------------------------------------------------------------------------
Public Sub TestBITSAdmin(Optional ByVal strQuelle As String = "", _
                          Optional ByVal strZiel As String = "")
    On Error GoTo ErrHandler

    InitFallback strQuelle, strZiel, "test_bits.dat"

    Debug.Print "  [bitsadmin] Start..."

    ' Pruefen ob bitsadmin existiert
    Dim strBits As String
    strBits = Environ("SystemRoot") & "\System32\bitsadmin.exe"
    If Dir(strBits) = "" Then
        Debug.Print "  [bitsadmin] NICHT VERFUEGBAR: " & strBits & " nicht gefunden"
        Exit Sub
    End If

    ' BITS Job erstellen + starten
    ' Hinweis: bitsadmin ist deprecated zugunsten von PowerShell BITS-Cmdlets,
    ' aber die .exe ist noch ueberall vorhanden und braucht kein Admin
    Dim strJobName As String
    strJobName = "VBA_KopierTest_" & Format(Now(), "yyyymmdd_hhnnss")

    Dim t1 As Long: t1 = GetTickCount()

    ' BITS braucht file:// URLs fuer lokale Dateien
    Dim strQuellURL As String
    strQuellURL = "file:///" & Replace(strQuelle, "\", "/")

    Dim lngPID As Long
    lngPID = Shell("cmd /c bitsadmin /create " & strJobName & _
             " & bitsadmin /addfile " & strJobName & " """ & strQuellURL & _
             """ """ & strZiel & """" & _
             " & bitsadmin /resume " & strJobName & _
             " & bitsadmin /complete " & strJobName, vbHide)

    Dim t2 As Long: t2 = GetTickCount()
    Debug.Print "  [bitsadmin] Shell zurueck nach " & (t2 - t1) & " ms (PID=" & lngPID & ")"

    ' Warten bis Datei da ist
    Dim blnOK As Boolean
    blnOK = WarteAufDatei(strZiel, 30000)

    Dim t3 As Long: t3 = GetTickCount()
    Debug.Print "  [bitsadmin] " & IIf(blnOK, "OK", "FEHLER") & _
                " - Shell: " & (t2 - t1) & " ms, Datei da: " & (t3 - t1) & " ms"
    Debug.Print "  [bitsadmin] -> ASYNC + Bandbreiten-Drosselung + Resume nach Neustart"
    If Not blnOK Then
        Debug.Print "  [bitsadmin] -> Moeglicherweise durch GPO blockiert"
    End If

    ' Cleanup: Job loeschen falls noch da
    On Error Resume Next
    Shell "cmd /c bitsadmin /cancel " & strJobName, vbHide
    On Error GoTo 0

    CleanupZiel strZiel
    Exit Sub
ErrHandler:
    Debug.Print "  [bitsadmin] FEHLER: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST 9: PowerShell Start-Job (wahrscheinlich blockiert)
' ---------------------------------------------------------------------------
Public Sub TestPowerShellCopy(Optional ByVal strQuelle As String = "", _
                               Optional ByVal strZiel As String = "")
    On Error GoTo ErrHandler

    InitFallback strQuelle, strZiel, "test_powershell.dat"

    Debug.Print "  [PowerShell] Start..."

    ' Test 1: Ist powershell.exe ueberhaupt aufrufbar?
    Dim t1 As Long: t1 = GetTickCount()

    Dim lngPID As Long
    lngPID = Shell("cmd /c powershell.exe -NoProfile -ExecutionPolicy Bypass " & _
                    "-Command ""Copy-Item '" & strQuelle & "' '" & strZiel & "'""", vbHide)

    Dim t2 As Long: t2 = GetTickCount()
    Debug.Print "  [PowerShell] Shell zurueck nach " & (t2 - t1) & " ms (PID=" & lngPID & ")"

    ' Warten
    Dim blnOK As Boolean
    blnOK = WarteAufDatei(strZiel, 15000)

    Dim t3 As Long: t3 = GetTickCount()
    Debug.Print "  [PowerShell] " & IIf(blnOK, "OK", "FEHLER/BLOCKIERT") & _
                " - Shell: " & (t2 - t1) & " ms, Datei da: " & (t3 - t1) & " ms"
    If Not blnOK Then
        Debug.Print "  [PowerShell] -> Wie erwartet: Execution Policy blockiert Ausfuehrung"
        Debug.Print "  [PowerShell] -> Alternativ: -ExecutionPolicy Bypass koennte gehen"
    Else
        Debug.Print "  [PowerShell] -> ASYNC + maechtiges Scripting moeglich"
    End If

    CleanupZiel strZiel
    Exit Sub
ErrHandler:
    Debug.Print "  [PowerShell] FEHLER: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST 10: Timer-basierte Info (kein echter Test, nur Erklaerung)
' ---------------------------------------------------------------------------
Public Sub TestTimerInfo()
    Debug.Print "  [Access Timer] Kein automatischer Test moeglich (braucht Formular)."
    Debug.Print "  [Access Timer] KONZEPT:"
    Debug.Print "    1. Access-Formular mit TimerInterval = 500 (ms)"
    Debug.Print "    2. Timer-Event nimmt naechste Datei aus der Queue"
    Debug.Print "    3. Kopiert mit FileCopy EINE Datei (kleine Blockierung)"
    Debug.Print "    4. Gibt Kontrolle zurueck an Access (DoEvents eingebaut)"
    Debug.Print "    5. Naechstes Timer-Event -> naechste Datei"
    Debug.Print "  [Access Timer] -> SEMI-ASYNC: Access bleibt nutzbar zwischen Kopien"
    Debug.Print "  [Access Timer] -> Funktioniert IMMER, kein Shell/Admin/Policy noetig"
    Debug.Print "  [Access Timer] -> Ideal als Fallback wenn Shell-Methoden blockiert sind"
    Debug.Print "  [Access Timer] -> Zusaetzlich nutzbar fuer Fortschrittsanzeige"
End Sub


' ===========================================================================
' BONUS: Kombinations-Test (empfohlene Produkt-Strategie)
' ===========================================================================

' Testet die empfohlene 3-Stufen-Strategie:
'   1. Primaer: Shell robocopy (async + retry + netzwerk-optimiert)
'   2. Fallback 1: WScript.Shell.Run cmd /c copy (async, kein robocopy noetig)
'   3. Fallback 2: FileCopy (synchron, funktioniert immer)
Public Sub TestEmpfohleneStrategie(Optional ByVal strNetzwerkPfad As String = "")
    Debug.Print String(80, "=")
    Debug.Print "  TEST: Empfohlene 3-Stufen-Kopierstrategie"
    Debug.Print String(80, "=")
    Debug.Print ""

    m_strTempDir = Environ("TEMP") & "\KopierTest\"
    If Dir(m_strTempDir, vbDirectory) = "" Then MkDir m_strTempDir

    If strNetzwerkPfad <> "" Then
        m_strNetzwerkDir = strNetzwerkPfad
        If Right(m_strNetzwerkDir, 1) <> "\" Then m_strNetzwerkDir = m_strNetzwerkDir & "\"
        m_strNetzwerkDir = m_strNetzwerkDir & "KopierTest\"
    Else
        m_strNetzwerkDir = Environ("TEMP") & "\KopierTestZiel\"
    End If
    If Dir(m_strNetzwerkDir, vbDirectory) = "" Then MkDir m_strNetzwerkDir

    m_lngTestDateiMB = 5

    ' Testdatei
    Dim strTestDatei As String
    strTestDatei = ErstelleTestDatei(5)
    If strTestDatei = "" Then
        Debug.Print "*** Testdatei-Erstellung fehlgeschlagen!"
        Exit Sub
    End If

    ' --- Stufe 1: robocopy ---
    Debug.Print "STUFE 1: robocopy"
    Dim strZiel1 As String: strZiel1 = m_strNetzwerkDir & "strat_robocopy.dat"
    Dim blnRobo As Boolean

    Dim strRobo As String
    strRobo = Environ("SystemRoot") & "\System32\robocopy.exe"
    If Dir(strRobo) <> "" Then
        Dim strSrcDir As String, strDatei As String
        strSrcDir = Left(strTestDatei, InStrRev(strTestDatei, "\") - 1)
        strDatei = Mid(strTestDatei, InStrRev(strTestDatei, "\") + 1)
        Dim strDstDir As String
        strDstDir = Left(strZiel1, InStrRev(strZiel1, "\") - 1)

        Dim t1 As Long: t1 = GetTickCount()
        Shell "cmd /c robocopy """ & strSrcDir & """ """ & strDstDir & _
              """ """ & strDatei & """ /R:2 /W:1 /NJH /NJS & ren """ & _
              strDstDir & "\" & strDatei & """ strat_robocopy.dat", vbHide

        blnRobo = WarteAufDatei(strZiel1, 15000)
        Debug.Print "  -> " & IIf(blnRobo, "OK (" & (GetTickCount() - t1) & " ms)", "FEHLER")
    Else
        Debug.Print "  -> NICHT VERFUEGBAR"
        blnRobo = False
    End If
    CleanupZiel strZiel1

    ' --- Stufe 2: WScript.Shell ---
    Debug.Print "STUFE 2: WScript.Shell.Run"
    Dim strZiel2 As String: strZiel2 = m_strNetzwerkDir & "strat_wscript.dat"
    Dim blnWS As Boolean

    On Error Resume Next
    Dim objSh As Object
    Set objSh = CreateObject("WScript.Shell")
    If Err.Number = 0 Then
        Dim t2 As Long: t2 = GetTickCount()
        objSh.Run "cmd /c copy """ & strTestDatei & """ """ & strZiel2 & """", 0, False
        blnWS = WarteAufDatei(strZiel2, 15000)
        Debug.Print "  -> " & IIf(blnWS, "OK (" & (GetTickCount() - t2) & " ms)", "FEHLER")
    Else
        Debug.Print "  -> BLOCKIERT: " & Err.Description
        blnWS = False
    End If
    Err.Clear
    On Error GoTo 0
    Set objSh = Nothing
    CleanupZiel strZiel2

    ' --- Stufe 3: FileCopy ---
    Debug.Print "STUFE 3: FileCopy (Fallback)"
    Dim strZiel3 As String: strZiel3 = m_strNetzwerkDir & "strat_filecopy.dat"
    Dim blnFC As Boolean

    On Error Resume Next
    Dim t3 As Long: t3 = GetTickCount()
    FileCopy strTestDatei, strZiel3
    blnFC = (Err.Number = 0 And Dir(strZiel3) <> "")
    Debug.Print "  -> " & IIf(blnFC, "OK (" & (GetTickCount() - t3) & " ms)", "FEHLER: " & Err.Description)
    On Error GoTo 0
    CleanupZiel strZiel3

    ' --- Zusammenfassung ---
    Debug.Print ""
    Debug.Print "ERGEBNIS:"
    Debug.Print "  Robocopy       : " & IIf(blnRobo, "VERFUEGBAR (async)", "NICHT VERFUEGBAR")
    Debug.Print "  WScript.Shell  : " & IIf(blnWS, "VERFUEGBAR (async)", "BLOCKIERT")
    Debug.Print "  FileCopy       : " & IIf(blnFC, "VERFUEGBAR (sync)", "FEHLER?!")
    Debug.Print ""

    If blnRobo Then
        Debug.Print ">>> EMPFEHLUNG: robocopy als primaere Methode verwenden."
        Debug.Print "    Vorteile: Async, eingebauter Retry, Netzwerk-optimiert, kein Admin."
    ElseIf blnWS Then
        Debug.Print ">>> EMPFEHLUNG: WScript.Shell.Run als primaere Methode verwenden."
        Debug.Print "    Vorteile: Async, COM-basiert, kein externes Tool."
    Else
        Debug.Print ">>> Nur FileCopy verfuegbar. Timer-basierte Strategie empfohlen:"
        Debug.Print "    -> Access-Formular mit Timer, eine Datei pro Tick."
    End If
    Debug.Print String(80, "=")

    AufraeumenTestdateien strTestDatei
End Sub


' ===========================================================================
' HILFSFUNKTIONEN
' ===========================================================================

' Erstellt eine Testdatei mit zufaelligem Inhalt (konfigurierbare Groesse)
Private Function ErstelleTestDatei(ByVal lngMB As Long) As String
    On Error GoTo ErrHandler

    Dim strPfad As String
    strPfad = m_strTempDir & "kopiertest_" & lngMB & "mb.dat"

    ' Wenn schon da und richtige Groesse, wiederverwenden
    If Dir(strPfad) <> "" Then
        If FileLen(strPfad) >= (lngMB * 1024& * 1024& * 0.9) Then
            ErstelleTestDatei = strPfad
            Exit Function
        End If
    End If

    ' Neu erstellen: 1-MB-Bloecke schreiben
    Dim intFile As Integer
    intFile = FreeFile
    Open strPfad For Binary Access Write As #intFile

    ' 1 MB Block mit wiederholendem Text
    Dim strBlock As String
    Dim lngBlockSize As Long
    lngBlockSize = 65536  ' 64 KB Bloecke
    strBlock = String(lngBlockSize, "X")  ' Platzhalter

    Dim lngGeschrieben As Long
    Dim lngZiel As Long
    lngZiel = lngMB * 1024& * 1024&

    Do While lngGeschrieben < lngZiel
        Put #intFile, , strBlock
        lngGeschrieben = lngGeschrieben + lngBlockSize
        If lngGeschrieben Mod (1024& * 1024&) = 0 Then DoEvents
    Loop

    Close #intFile

    ErstelleTestDatei = strPfad
    Exit Function

ErrHandler:
    Close #intFile
    ErstelleTestDatei = ""
End Function


' Warte auf einen Shell-Prozess (nicht-blockierend mit DoEvents)
Private Function WarteAufProzess(ByVal lngPID As Long, _
                                  ByVal lngTimeoutMS As Long) As Boolean
    On Error Resume Next

    #If VBA7 Then
        Dim hProc As LongPtr
    #Else
        Dim hProc As Long
    #End If

    hProc = OpenProcess(SYNCHRONIZE, 0, lngPID)
    If hProc = 0 Then
        ' Prozess schon beendet oder kein Zugriff
        WarteAufProzess = True
        Exit Function
    End If

    Dim lngStart As Long
    lngStart = GetTickCount()

    Do
        Dim lngWait As Long
        lngWait = WaitForSingleObject(hProc, 100)  ' 100ms warten
        DoEvents  ' Access UI responsiv halten!

        If lngWait <> WAIT_TIMEOUT Then
            ' Prozess beendet
            WarteAufProzess = True
            CloseHandle hProc
            Exit Function
        End If

        ' Timeout pruefen
        If (GetTickCount() - lngStart) > lngTimeoutMS Then
            WarteAufProzess = False
            CloseHandle hProc
            Exit Function
        End If
    Loop

    CloseHandle hProc
    WarteAufProzess = True
End Function


' Warte bis eine Datei existiert (Polling mit DoEvents)
Private Function WarteAufDatei(ByVal strPfad As String, _
                                ByVal lngTimeoutMS As Long) As Boolean
    Dim lngStart As Long
    lngStart = GetTickCount()

    Do
        If Dir(strPfad) <> "" Then
            ' Datei existiert - noch pruefen ob Groesse > 0
            ' (koennte noch geschrieben werden)
            DoEvents
            Sleep 200
            If Dir(strPfad) <> "" And FileLen(strPfad) > 0 Then
                WarteAufDatei = True
                Exit Function
            End If
        End If

        DoEvents
        Sleep 250

        If (GetTickCount() - lngStart) > lngTimeoutMS Then
            WarteAufDatei = False
            Exit Function
        End If
    Loop
End Function


' Fallback-Initialisierung fuer Einzelaufrufe
Private Sub InitFallback(ByRef strQuelle As String, ByRef strZiel As String, _
                          ByVal strDateiname As String)
    If m_strTempDir = "" Then
        m_strTempDir = Environ("TEMP") & "\KopierTest\"
        If Dir(m_strTempDir, vbDirectory) = "" Then MkDir m_strTempDir
    End If

    If m_strNetzwerkDir = "" Then
        m_strNetzwerkDir = Environ("TEMP") & "\KopierTestZiel\"
        If Dir(m_strNetzwerkDir, vbDirectory) = "" Then MkDir m_strNetzwerkDir
    End If

    If strQuelle = "" Then
        m_lngTestDateiMB = 5
        strQuelle = ErstelleTestDatei(5)
    End If

    If strZiel = "" Then
        strZiel = m_strNetzwerkDir & strDateiname
    End If
End Sub


' Zieldatei loeschen
Private Sub CleanupZiel(ByVal strZiel As String)
    On Error Resume Next
    If Dir(strZiel) <> "" Then Kill strZiel
    On Error GoTo 0
End Sub


' Alle Testdateien aufraeumen
Private Sub AufraeumenTestdateien(ByVal strTestDatei As String)
    On Error Resume Next

    ' Testdatei loeschen
    If strTestDatei <> "" And Dir(strTestDatei) <> "" Then
        Kill strTestDatei
    End If

    ' Zielordner leeren + loeschen
    If m_strNetzwerkDir <> "" Then
        If Dir(m_strNetzwerkDir & "*.*") <> "" Then
            Kill m_strNetzwerkDir & "*.*"
        End If
        RmDir m_strNetzwerkDir
    End If

    ' Temp-Ordner leeren + loeschen
    If m_strTempDir <> "" Then
        If Dir(m_strTempDir & "*.*") <> "" Then
            Kill m_strTempDir & "*.*"
        End If
        RmDir m_strTempDir
    End If

    On Error GoTo 0
End Sub


' String rechts mit Leerzeichen auffuellen
Private Function PadR(ByVal strText As String, ByVal lngLen As Long) As String
    If Len(strText) >= lngLen Then
        PadR = Left(strText, lngLen)
    Else
        PadR = strText & Space(lngLen - Len(strText))
    End If
End Function


' Dateiendung extrahieren
Private Function HoleEndung(ByVal strDatei As String) As String
    Dim lngPos As Long
    lngPos = InStrRev(strDatei, ".")
    If lngPos > 0 Then
        HoleEndung = Mid(strDatei, lngPos + 1)
    Else
        HoleEndung = ""
    End If
End Function


' Config lesen (Standalone-Version fuer Testmodul)
Private Function LeseConfig(ByVal strKey As String, _
                             ByVal strDefault As String) As String
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Wert FROM tblConfig WHERE Schluessel='" & strKey & "'", _
        dbOpenSnapshot)
    If Not rs.EOF Then
        LeseConfig = Nz(rs!Wert, strDefault)
    Else
        LeseConfig = strDefault
    End If
    rs.Close: Set rs = Nothing

    If Err.Number <> 0 Then LeseConfig = strDefault
    Err.Clear
    On Error GoTo 0
End Function


' Test-Runner: Fuehrt einen Test durch und sammelt Ergebnis
Private Sub RunTest(ByVal strTestDatei As String, _
                     ByVal strName As String, _
                     ByVal strMethode As String, _
                     ByRef aErgebnisse() As String, _
                     ByVal intIndex As Integer)
    On Error Resume Next

    Debug.Print ""
    Debug.Print "--- " & strName & " ---"

    Dim strZiel As String
    Dim t1 As Long: t1 = GetTickCount()

    Select Case strMethode
        Case "TestFileCopyBaseline"
            strZiel = m_strNetzwerkDir & "test_filecopy.dat"
            FileCopy strTestDatei, strZiel
            Dim t2fc As Long: t2fc = GetTickCount()

            aErgebnisse(intIndex, 0) = strName
            If Err.Number = 0 And Dir(strZiel) <> "" Then
                aErgebnisse(intIndex, 1) = "OK"
                aErgebnisse(intIndex, 3) = (t2fc - t1) & " ms"
            Else
                aErgebnisse(intIndex, 1) = "FEHLER"
                aErgebnisse(intIndex, 3) = "-"
            End If
            aErgebnisse(intIndex, 2) = "SYNC"
            aErgebnisse(intIndex, 4) = "Baseline, immer verfuegbar"
            CleanupZiel strZiel
            Debug.Print "  -> " & aErgebnisse(intIndex, 1) & " " & aErgebnisse(intIndex, 3)
            Err.Clear

        Case "TestShellCmdCopy"
            strZiel = m_strNetzwerkDir & "test_cmdcopy.dat"
            Dim pidCmd As Long
            pidCmd = Shell("cmd /c copy """ & strTestDatei & """ """ & strZiel & """", vbHide)
            Dim t2cmd As Long: t2cmd = GetTickCount()
            Dim blnCmd As Boolean: blnCmd = WarteAufDatei(strZiel, 15000)

            aErgebnisse(intIndex, 0) = strName
            If blnCmd Then
                aErgebnisse(intIndex, 1) = "OK"
                aErgebnisse(intIndex, 3) = (t2cmd - t1) & "/" & (GetTickCount() - t1) & " ms"
            Else
                aErgebnisse(intIndex, 1) = IIf(Err.Number <> 0, "FEHLER", "TIMEOUT")
                aErgebnisse(intIndex, 3) = "-"
            End If
            aErgebnisse(intIndex, 2) = "ASYNC"
            aErgebnisse(intIndex, 4) = "Shell kehrt sofort zurueck"
            CleanupZiel strZiel
            Debug.Print "  -> " & aErgebnisse(intIndex, 1)
            Err.Clear

        Case "TestShellRobocopy"
            strZiel = m_strNetzwerkDir & "test_robocopy.dat"
            Dim strRoboExe As String
            strRoboExe = Environ("SystemRoot") & "\System32\robocopy.exe"

            aErgebnisse(intIndex, 0) = strName
            If Dir(strRoboExe) <> "" Then
                Dim strSD As String, strDD As String, strFN As String
                strSD = Left(strTestDatei, InStrRev(strTestDatei, "\") - 1)
                strFN = Mid(strTestDatei, InStrRev(strTestDatei, "\") + 1)
                strDD = Left(strZiel, InStrRev(strZiel, "\") - 1)

                Dim pidRobo As Long
                pidRobo = Shell("cmd /c robocopy """ & strSD & """ """ & strDD & _
                          """ """ & strFN & """ /R:1 /W:1 /NJH /NJS & ren """ & _
                          strDD & "\" & strFN & """ test_robocopy.dat", vbHide)
                Dim t2r As Long: t2r = GetTickCount()
                Dim blnRobo As Boolean: blnRobo = WarteAufDatei(strZiel, 15000)

                If blnRobo Then
                    aErgebnisse(intIndex, 1) = "OK"
                    aErgebnisse(intIndex, 3) = (t2r - t1) & "/" & (GetTickCount() - t1) & " ms"
                    aErgebnisse(intIndex, 4) = "EMPFOHLEN: async+retry+netzwerk"
                Else
                    aErgebnisse(intIndex, 1) = "TIMEOUT"
                    aErgebnisse(intIndex, 3) = "-"
                    aErgebnisse(intIndex, 4) = "Verfuegbar aber langsam?"
                End If
                aErgebnisse(intIndex, 2) = "ASYNC"
            Else
                aErgebnisse(intIndex, 1) = "N/A"
                aErgebnisse(intIndex, 2) = "-"
                aErgebnisse(intIndex, 3) = "-"
                aErgebnisse(intIndex, 4) = "robocopy.exe nicht gefunden"
            End If
            CleanupZiel strZiel
            Debug.Print "  -> " & aErgebnisse(intIndex, 1)
            Err.Clear

        Case "TestShellXcopy"
            strZiel = m_strNetzwerkDir & "test_xcopy.dat"
            Dim pidXC As Long
            pidXC = Shell("cmd /c echo F | xcopy """ & strTestDatei & """ """ & strZiel & """ /Y /Q", vbHide)
            Dim t2xc As Long: t2xc = GetTickCount()
            Dim blnXC As Boolean: blnXC = WarteAufDatei(strZiel, 15000)

            aErgebnisse(intIndex, 0) = strName
            If blnXC Then
                aErgebnisse(intIndex, 1) = "OK"
                aErgebnisse(intIndex, 3) = (t2xc - t1) & "/" & (GetTickCount() - t1) & " ms"
            Else
                aErgebnisse(intIndex, 1) = IIf(Err.Number <> 0, "FEHLER", "TIMEOUT")
                aErgebnisse(intIndex, 3) = "-"
            End If
            aErgebnisse(intIndex, 2) = "ASYNC"
            aErgebnisse(intIndex, 4) = "xcopy /Y /Q"
            CleanupZiel strZiel
            Debug.Print "  -> " & aErgebnisse(intIndex, 1)
            Err.Clear

        Case "TestShellStartCopy"
            strZiel = m_strNetzwerkDir & "test_startcopy.dat"
            Dim pidSB As Long
            pidSB = Shell("cmd /c start /b cmd /c copy """ & strTestDatei & """ """ & strZiel & """", vbHide)
            Dim t2sb As Long: t2sb = GetTickCount()
            Dim blnSB As Boolean: blnSB = WarteAufDatei(strZiel, 15000)

            aErgebnisse(intIndex, 0) = strName
            If blnSB Then
                aErgebnisse(intIndex, 1) = "OK"
                aErgebnisse(intIndex, 3) = (t2sb - t1) & "/" & (GetTickCount() - t1) & " ms"
            Else
                aErgebnisse(intIndex, 1) = IIf(Err.Number <> 0, "FEHLER", "TIMEOUT")
                aErgebnisse(intIndex, 3) = "-"
            End If
            aErgebnisse(intIndex, 2) = "ASYNC"
            aErgebnisse(intIndex, 4) = "Losgeloester Prozess"
            CleanupZiel strZiel
            Debug.Print "  -> " & aErgebnisse(intIndex, 1)
            Err.Clear

        Case "TestWScriptShellRun"
            strZiel = m_strNetzwerkDir & "test_wscriptrun.dat"

            aErgebnisse(intIndex, 0) = strName
            Dim objWS As Object
            Set objWS = CreateObject("WScript.Shell")
            If Err.Number = 0 Then
                objWS.Run "cmd /c copy """ & strTestDatei & """ """ & strZiel & """", 0, False
                Dim t2ws As Long: t2ws = GetTickCount()
                Dim blnWS2 As Boolean: blnWS2 = WarteAufDatei(strZiel, 15000)

                If blnWS2 Then
                    aErgebnisse(intIndex, 1) = "OK"
                    aErgebnisse(intIndex, 3) = (t2ws - t1) & "/" & (GetTickCount() - t1) & " ms"
                Else
                    aErgebnisse(intIndex, 1) = "TIMEOUT"
                    aErgebnisse(intIndex, 3) = "-"
                End If
                aErgebnisse(intIndex, 2) = "ASYNC"
                aErgebnisse(intIndex, 4) = "COM-basiert, bWaitOnReturn=False"
            Else
                aErgebnisse(intIndex, 1) = "BLOCKIERT"
                aErgebnisse(intIndex, 2) = "-"
                aErgebnisse(intIndex, 3) = "-"
                aErgebnisse(intIndex, 4) = "CreateObject blockiert"
            End If
            Set objWS = Nothing
            CleanupZiel strZiel
            Debug.Print "  -> " & aErgebnisse(intIndex, 1)
            Err.Clear

        Case "TestCopyFileExAPI"
            strZiel = m_strNetzwerkDir & "test_copyfileex.dat"
            Dim lngCancel As Long: lngCancel = 0
            Dim lngRet As Long
            lngRet = CopyFileEx(strTestDatei, strZiel, 0, 0, lngCancel, 0)
            Dim t2cf As Long: t2cf = GetTickCount()

            aErgebnisse(intIndex, 0) = strName
            If lngRet <> 0 And Dir(strZiel) <> "" Then
                aErgebnisse(intIndex, 1) = "OK"
                aErgebnisse(intIndex, 3) = (t2cf - t1) & " ms"
            Else
                aErgebnisse(intIndex, 1) = "FEHLER"
                aErgebnisse(intIndex, 3) = "-"
            End If
            aErgebnisse(intIndex, 2) = "SYNC"
            aErgebnisse(intIndex, 4) = "Cancel+Progress moeglich"
            CleanupZiel strZiel
            Debug.Print "  -> " & aErgebnisse(intIndex, 1)
            Err.Clear

        Case "TestBITSAdmin"
            strZiel = m_strNetzwerkDir & "test_bits.dat"
            Dim strBitsExe As String
            strBitsExe = Environ("SystemRoot") & "\System32\bitsadmin.exe"

            aErgebnisse(intIndex, 0) = strName
            If Dir(strBitsExe) <> "" Then
                Dim strJob As String
                strJob = "VBA_Test_" & Format(Now(), "hhnnss")
                Dim strURL As String
                strURL = "file:///" & Replace(strTestDatei, "\", "/")

                Shell "cmd /c bitsadmin /create " & strJob & _
                      " & bitsadmin /addfile " & strJob & " """ & strURL & _
                      """ """ & strZiel & """" & _
                      " & bitsadmin /resume " & strJob & _
                      " & bitsadmin /complete " & strJob, vbHide
                Dim t2bits As Long: t2bits = GetTickCount()
                Dim blnBits As Boolean: blnBits = WarteAufDatei(strZiel, 15000)

                If blnBits Then
                    aErgebnisse(intIndex, 1) = "OK"
                    aErgebnisse(intIndex, 3) = (t2bits - t1) & "/" & (GetTickCount() - t1) & " ms"
                    aErgebnisse(intIndex, 4) = "Bandbreiten-Drosselung moegl."
                Else
                    aErgebnisse(intIndex, 1) = "FEHLER/GPO"
                    aErgebnisse(intIndex, 3) = "-"
                    aErgebnisse(intIndex, 4) = "Evtl. GPO-blockiert"
                End If
                aErgebnisse(intIndex, 2) = "ASYNC"

                On Error Resume Next
                Shell "cmd /c bitsadmin /cancel " & strJob, vbHide
                On Error GoTo 0
            Else
                aErgebnisse(intIndex, 1) = "N/A"
                aErgebnisse(intIndex, 2) = "-"
                aErgebnisse(intIndex, 3) = "-"
                aErgebnisse(intIndex, 4) = "bitsadmin.exe nicht gefunden"
            End If
            CleanupZiel strZiel
            Debug.Print "  -> " & aErgebnisse(intIndex, 1)
            Err.Clear

        Case "TestPowerShellCopy"
            strZiel = m_strNetzwerkDir & "test_powershell.dat"
            Dim pidPS As Long
            pidPS = Shell("cmd /c powershell.exe -NoProfile -ExecutionPolicy Bypass " & _
                          "-Command ""Copy-Item '" & strTestDatei & "' '" & strZiel & "'""", vbHide)
            Dim t2ps As Long: t2ps = GetTickCount()
            Dim blnPS As Boolean: blnPS = WarteAufDatei(strZiel, 15000)

            aErgebnisse(intIndex, 0) = strName
            If blnPS Then
                aErgebnisse(intIndex, 1) = "OK"
                aErgebnisse(intIndex, 3) = (t2ps - t1) & "/" & (GetTickCount() - t1) & " ms"
                aErgebnisse(intIndex, 4) = "-ExecutionPolicy Bypass geht!"
            Else
                aErgebnisse(intIndex, 1) = "BLOCKIERT"
                aErgebnisse(intIndex, 3) = "-"
                aErgebnisse(intIndex, 4) = "Execution Policy / GPO"
            End If
            aErgebnisse(intIndex, 2) = "ASYNC"
            CleanupZiel strZiel
            Debug.Print "  -> " & aErgebnisse(intIndex, 1)
            Err.Clear

        Case "TestTimerInfo"
            aErgebnisse(intIndex, 0) = strName
            aErgebnisse(intIndex, 1) = "MOEGLICH"
            aErgebnisse(intIndex, 2) = "SEMI"
            aErgebnisse(intIndex, 3) = "-"
            aErgebnisse(intIndex, 4) = "Braucht Form, immer verfuegbar"
            Debug.Print "  -> Konzept (braucht Access-Formular)"

    End Select
    On Error GoTo 0
End Sub
