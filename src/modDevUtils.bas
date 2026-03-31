Option Compare Database
Option Explicit

' ===========================================================================
' modDevUtils - Entwickler-Utilities, Fehlerbehandlung, Diagnose
' ===========================================================================
' v0.5: Zentrales Modul fuer Dev-Mode, Error-Handling, Timing, Diagnose
'
' ERROR-HANDLING:
'   HandleError        - Standardisierte Fehlerbehandlung (ersetzt LogVBAError-Muster)
'   HandleErrorSilent  - Wie HandleError, aber ohne Re-Raise (fuer Resume Next-Stellen)
'   RaiseDevError      - Fehler nur im DevMode eskalieren
'
' DEV-MODE:
'   DevModeEin / DevModeAus / IstDevMode
'   DevInfo            - Diagnoseinformationen im Direktbereich
'   DevAssert          - Assertions (nur im DevMode aktiv)
'
' TIMING:
'   TimerStart / TimerStop / TimerReport
'   (mehrere benannte Timer gleichzeitig moeglich)
'
' DIAGNOSE:
'   ProjektStatus      - Uebersicht: Tabellen, Config, Backend
'   TabellenStatus     - Alle Tabellen mit Datensatzanzahl
'   ConfigDump         - Alle Config-Werte ausgeben
'   SchemaCheck        - Schema-Version pruefen + Warnung
'
' Abhaengigkeiten: modGlobals (Konstanten), modLogging (Log-Funktionen),
'                  modSchemaTools (TabelleExistiert), modCache (CacheGetConfig)
' ===========================================================================


' ---------------------------------------------------------------------------
' DEV-MODE FLAG
' ---------------------------------------------------------------------------
Private m_blnDevMode As Boolean

Private m_dictTimer As Object

' ===========================================================================
' ERROR-HANDLING
' ===========================================================================

' Standardisierte Fehlerbehandlung - in jedem ErrHandler aufrufen.
' Loggt den Fehler, gibt im DevMode zusaetzliche Details aus,
' und gibt die Fehlernummer zurueck (0 = kein Fehler).
'
' Aufruf:
'   ErrHandler:
'       HandleError "modDAO", "SpeichereEmail"
'
Public Function HandleError(ByVal strModul As String, _
                             ByVal strProzedur As String, _
                             Optional ByVal strKontext As String = "") As Long
    Dim lngErr As Long
    Dim strDesc As String
    Dim strSrc As String

    lngErr = Err.Number
    strDesc = Err.Description
    strSrc = Err.Source

    If lngErr = 0 Then HandleError = 0: Exit Function

    ' Fehler loggen (nutzt bestehendes Logging-System)
    If strKontext <> "" Then
        LogError strModul & "." & strProzedur & ": [" & lngErr & "] " & _
                 strDesc & " | " & strKontext, strModul
    Else
        LogError strModul & "." & strProzedur & ": [" & lngErr & "] " & strDesc, strModul
    End If

    ' Im DevMode: zusaetzliche Details im Direktbereich
    If m_blnDevMode Then
        Debug.Print String(50, "-")
        Debug.Print "[DEV-ERROR] " & strModul & "." & strProzedur
        Debug.Print "  Err.Number: " & lngErr
        Debug.Print "  Err.Desc:   " & strDesc
        Debug.Print "  Err.Source: " & strSrc
        If strKontext <> "" Then Debug.Print "  Kontext:    " & strKontext
        Debug.Print "  Zeit:       " & Now()
        Debug.Print String(50, "-")
    End If

    HandleError = lngErr
End Function


' Wie HandleError, aber fuer Stellen mit Resume Next:
' Loggt nur wenn tatsaechlich ein Fehler vorliegt, cleared danach.
Public Sub HandleErrorSilent(ByVal strModul As String, _
                              ByVal strProzedur As String)
    If Err.Number = 0 Then Exit Sub

    HandleError strModul, strProzedur
    Err.Clear
End Sub


' ===========================================================================
' DEV-MODE STEUERUNG
' ===========================================================================

Public Sub DevModeEin()
    m_blnDevMode = True
    LogInfo "DevMode AKTIVIERT", "DEV"
    Debug.Print ">>> DevMode EIN <<<"
End Sub

Public Sub DevModeAus()
    m_blnDevMode = False
    LogInfo "DevMode DEAKTIVIERT", "DEV"
    Debug.Print ">>> DevMode AUS <<<"
End Sub

Public Function IstDevMode() As Boolean
    IstDevMode = m_blnDevMode
End Function

' Assertion - nur im DevMode aktiv (zero-cost in Produktion)
Public Sub DevAssert(ByVal blnBedingung As Boolean, _
                      Optional ByVal strMsg As String = "Assertion fehlgeschlagen")
    If Not m_blnDevMode Then Exit Sub
    If blnBedingung Then Exit Sub

    Debug.Print "[ASSERT FAILED] " & strMsg
    LogWarn "DevAssert: " & strMsg, "DEV"
End Sub


' Vergangene Sekunden robust auch ueber Mitternacht (Timer-Reset).
Public Function SekundenDiff(ByVal dblStart As Double, ByVal dblEnd As Double) As Double
    If dblEnd >= dblStart Then
        SekundenDiff = dblEnd - dblStart
    Else
        SekundenDiff = (86400# - dblStart) + dblEnd
    End If
End Function


' ===========================================================================
' TIMER (Performance-Messung)
' ===========================================================================

Public Sub TimerStart(Optional ByVal strName As String = "default")
    If m_dictTimer Is Nothing Then Set m_dictTimer = CreateObject("Scripting.Dictionary")
    m_dictTimer(strName) = Timer
    If m_blnDevMode Then Debug.Print "[TIMER] Start: " & strName
End Sub

Public Function TimerStop(Optional ByVal strName As String = "default") As Double
    If m_dictTimer Is Nothing Then TimerStop = 0: Exit Function
    If Not m_dictTimer.Exists(strName) Then TimerStop = 0: Exit Function

    TimerStop = Timer - CDbl(m_dictTimer(strName))
    m_dictTimer.Remove strName

    If m_blnDevMode Then
        Debug.Print "[TIMER] " & strName & ": " & Format(TimerStop, "0.000") & "s"
    End If
End Function

Public Sub TimerReport()
    If m_dictTimer Is Nothing Then
        Debug.Print "[TIMER] Keine aktiven Timer."
        Exit Sub
    End If

    Dim k As Variant
    Debug.Print String(50, "-")
    Debug.Print "[TIMER] Aktive Timer:"
    For Each k In m_dictTimer.Keys
        Debug.Print "  " & k & ": laeuft seit " & _
                    Format(Timer - CDbl(m_dictTimer(k)), "0.0") & "s"
    Next k
    Debug.Print String(50, "-")
End Sub


' ===========================================================================
' DIAGNOSE
' ===========================================================================

' Gesamtstatus des Projekts ausgeben
Public Sub ProjektStatus()
    Debug.Print String(70, "=")
    Debug.Print "=== PROJEKT-STATUS ==="
    Debug.Print "  Schema-Version : " & CacheGetConfig(CFG_SCHEMA_VERSION, "???")
    Debug.Print "  Backend        : " & CacheGetConfig(CFG_BACKEND_PFAD, "(lokal)")
    Debug.Print "  LogLevel       : " & g_intLogLevel & " (" & LevelText(g_intLogLevel) & ")"
    Debug.Print "  DevMode        : " & IIf(m_blnDevMode, "EIN", "AUS")
    Debug.Print "  Export-Pfad    : " & CacheGetConfig(CFG_EXPORT_PFAD, "(Standard)")
    Debug.Print "  Temp-Pfad      : " & CacheGetConfig(CFG_TEMP_PFAD, "(Standard)")
    Debug.Print "  Buffer-Groesse : " & CacheGetConfig(CFG_BUFFER_GROESSE, "25")
    Debug.Print String(70, "-")
    Call TabellenStatus
    Debug.Print String(70, "=")
End Sub


' Alle Tabellen mit Datensatzanzahl auflisten
Public Sub TabellenStatus()
    Dim arrTabellen As Variant
    arrTabellen = Array(TBL_CONFIG, TBL_SYNC_LAUF, TBL_KONTAKTE, _
                        TBL_OUTLOOK_ORDNER, TBL_EMAIL_THREADS, TBL_EMAILS, _
                        TBL_EMAIL_CONTENT, TBL_EMAIL_EMPFAENGER, _
                        TBL_EMAIL_ANHAENGE, TBL_EMAIL_STATUS, _
                        TBL_SYNC_PROFIL, TBL_SYNC_PROFIL_ORDNER, TBL_LOG)

    Dim i As Long
    Dim strTbl As String
    Dim lngCount As Long

    Debug.Print "  TABELLEN:"
    For i = LBound(arrTabellen) To UBound(arrTabellen)
        strTbl = CStr(arrTabellen(i))
        If TabelleExistiert(strTbl) Then
            On Error Resume Next
            lngCount = DCount("*", strTbl)
            If Err.Number <> 0 Then lngCount = -1: Err.Clear
            On Error GoTo 0
            Debug.Print "    " & PadRight(strTbl, 25) & lngCount & " Datensaetze"
        Else
            Debug.Print "    " & PadRight(strTbl, 25) & "[FEHLT]"
        End If
    Next i
End Sub


' Alle Config-Werte ausgeben
Public Sub ConfigDump()
    On Error GoTo ErrHandler

    If Not TabelleExistiert(TBL_CONFIG) Then
        Debug.Print "[WARN] " & TBL_CONFIG & " existiert nicht."
        Exit Sub
    End If

    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Schluessel, Wert, Beschreibung FROM [" & TBL_CONFIG & "] ORDER BY Schluessel", dbOpenSnapshot)

    Debug.Print String(70, "=")
    Debug.Print "=== KONFIGURATION ==="
    Do While Not rs.EOF
        Debug.Print "  " & PadRight(Nz(rs!Schluessel, ""), 25) & _
                    Nz(rs!Wert, "(leer)")
        rs.MoveNext
    Loop
    Debug.Print String(70, "=")

    rs.Close: Set rs = Nothing: Set db = Nothing
    Exit Sub

ErrHandler:
    HandleError "modDevUtils", "ConfigDump"
End Sub


' Schema-Version pruefen und warnen falls veraltet
Public Sub SchemaCheck()
    Dim strDB As String
    strDB = CacheGetConfig(CFG_SCHEMA_VERSION, "0.0")

    Debug.Print "  Schema DB : " & strDB
    Debug.Print "  Schema App: 0.5"

    If strDB <> "0.5" Then
        Debug.Print "  [WARNUNG] Schema-Version stimmt nicht ueberein!"
        Debug.Print "  -> ErstelleAlleTabellen ausfuehren fuer Update"
    Else
        Debug.Print "  [OK] Schema ist aktuell."
    End If
End Sub


' Pfad-Diagnose fuer Temp/Export/Backend/Worker.
' Gibt effektive Speicherorte inkl. Netzwerk/Lokal-Klassifikation aus.
Public Sub DevPfadDiagnose(Optional ByVal strKontext As String = "", _
                           Optional ByVal blnNurWennDevMode As Boolean = False, _
                           Optional ByVal strWorkerDbPfad As String = "")
    On Error GoTo ErrHandler

    If blnNurWennDevMode Then
        If Not m_blnDevMode Then Exit Sub
    End If

    Dim strCfgExport As String
    Dim strEffExport As String
    Dim strCfgTemp As String
    Dim strEffTemp As String
    Dim strBackend As String
    Dim strWorkerEff As String

    strCfgExport = Nz(CacheGetConfig(CFG_EXPORT_PFAD, ""), "")
    strEffExport = NormalisierePfad(CacheGetConfig(CFG_EXPORT_PFAD, Environ("USERPROFILE") & PATH_DEFAULT_FALLBACK))

    strCfgTemp = Nz(CacheGetConfig(CFG_TEMP_PFAD, ""), "")
    strEffTemp = HoleTempPfad()

    strBackend = Nz(GetBackendPfad(), "")

    If Nz(strWorkerDbPfad, "") = "" Then
        strWorkerEff = CurrentDb.Name
    Else
        strWorkerEff = strWorkerDbPfad
    End If

    Debug.Print String(70, "-")
    If strKontext <> "" Then
        Debug.Print "[PFAD-DIAGNOSE] " & strKontext
    Else
        Debug.Print "[PFAD-DIAGNOSE]"
    End If

    Debug.Print "  UserProfile        : " & Environ("USERPROFILE")
    Debug.Print "  TEMP (Env)         : " & Environ("TEMP")
    Debug.Print "  TMP  (Env)         : " & Environ("TMP")

    Debug.Print "  Export cfg         : " & IIf(strCfgExport = "", "(leer -> Fallback)", strCfgExport)
    Debug.Print "  Export effektiv    : " & strEffExport
    Debug.Print "  Export Typ         : " & PfadTypText(strEffExport)

    Debug.Print "  Temp cfg           : " & IIf(strCfgTemp = "", "(leer -> %TEMP%\\OutlookSync\\)", strCfgTemp)
    Debug.Print "  Temp effektiv      : " & strEffTemp
    Debug.Print "  Temp Typ           : " & PfadTypText(strEffTemp)

    Debug.Print "  Backend cfg        : " & IIf(strBackend = "", "(lokal, kein Backend)", strBackend)
    If strBackend <> "" Then
        Debug.Print "  Backend Typ        : " & PfadTypText(strBackend)
    End If

    Debug.Print "  Worker DB effektiv : " & strWorkerEff
    Debug.Print "  Worker DB Typ      : " & PfadTypText(strWorkerEff)
    Debug.Print String(70, "-")
    Exit Sub

ErrHandler:
    HandleError "modDevUtils", "DevPfadDiagnose", strKontext
End Sub


' ===========================================================================
' PRIVATE HELPER
' ===========================================================================

Private Function PadRight(ByVal s As String, ByVal l As Long) As String
    If Len(s) >= l Then
        PadRight = Left(s, l)
    Else
        PadRight = s & Space(l - Len(s))
    End If
End Function

Private Function LevelText(ByVal i As Integer) As String
    Select Case i
        Case LOG_NONE:  LevelText = "NONE"
        Case LOG_ERROR: LevelText = "ERROR"
        Case LOG_WARN:  LevelText = "WARN"
        Case LOG_INFO:  LevelText = "INFO"
        Case LOG_DEBUG: LevelText = "DEBUG"
        Case LOG_TRACE: LevelText = "TRACE"
        Case Else:      LevelText = "?"
    End Select
End Function

Private Function PfadTypText(ByVal strPfad As String) As String
    Dim strNorm As String
    strNorm = Nz(strPfad, "")

    If strNorm = "" Then
        PfadTypText = "(leer)"
        Exit Function
    End If

    If IstNetzwerkPfad(strNorm) Then
        PfadTypText = "Netzwerk"
    ElseIf InStr(1, strNorm, Environ("USERPROFILE"), vbTextCompare) = 1 Then
        PfadTypText = "Lokal (UserProfile)"
    Else
        PfadTypText = "Lokal"
    End If
End Function


