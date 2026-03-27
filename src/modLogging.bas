Option Compare Database
Option Explicit

' ===========================================================================
' modLogging v2.0 - Vereinheitlichte Logging-Infrastruktur
' ===========================================================================
' VERSION:  2.0 (v0.4.4)
' ZWECK:    Multi-Sink Logging fuer Entwicklung, Debug und Produktivbetrieb
'
' ARCHITEKTUR (4 Ausgabekanaele / Sinks):
'   Sink 1: Debug.Print      VBA-Direktbereich              (Default: ON)
'   Sink 2: Logfile          Tagesbasiert, Logfile_yyyymmdd  (Default: ON)
'   Sink 3: DB-Tabelle       tblLog, persistentes Logging    (Default: OFF)
'   Sink 4: StatusForm       Hook fuer frm_StatusLog         (Default: OFF)
'
' ABWAERTSKOMPATIBEL - alle bestehenden Aufrufe funktionieren 1:1:
'   LogInfo, LogWarn, LogError, LogDebug, LogTrace, LogVBAError
'
' NEU:
'   LogEvent    Strukturiertes Logging (Modul.Prozedur + Kontext)
'   LogDev      Nur im DevMode aktiv (zero-cost in Produktion)
'   LogSQL      SQL-Statements im DevMode tracen
'   Batches     StartLogBatch / EndLogBatch / GetLogBatchID
'   Statistik   GetLogStats / GetLogErrorCount / ResetLogStats
'   Fortschritt LogProgress_Start / _Step / _Status / _IsCancelled
'   DevMode     LogSetDevMode / LogStatus (Immediate-Window Diagnose)
'   DB-Setup    ErstelleLogTabelle (einmalig aus Direktbereich)
'
' LOG-LEVELS (aus modGlobals, unveraendert):
'   0=NONE  1=ERROR  2=WARN  3=INFO  4=DEBUG  5=TRACE
'
' STATUSFORM-INTERFACE (fuer zukuenftiges frm_StatusLog):
'   Das Formular muss folgende Public-Methoden bereitstellen:
'     Sub AddLog(strMsg As String, strLevel As String)
'     Sub StartProgress(strMsg As String)
'     Sub SetProgressSteps(lCurrent As Long, lTotal As Long, [strMsg])
'     Sub SetStatus(strMsg As String)
'     Property Get IsCancelled() As Boolean
'
' QUICKSTART (im Direktbereich, Strg+G):
'   LogSetDevMode True       ' DevMode ein (TRACE Level)
'   LogSetDevMode False      ' DevMode aus (INFO Level)
'   LogStatus                ' Aktuellen Status anzeigen
'   ErstelleLogTabelle       ' DB-Sink Tabelle erstellen
'   LogEnableDBSink True     ' DB-Logging aktivieren
'   LogSetStatusFormName "frm_StatusLog"  ' Form-Hook setzen
' ===========================================================================


' ---------------------------------------------------------------------------
' KONSTANTEN
' ---------------------------------------------------------------------------
Private Const LOG_FILE_PREFIX   As String = "Logfile_"
' LOG_DB_TABLE: Nutzt TBL_LOG aus modGlobals (keine lokale Kopie mehr)


' ---------------------------------------------------------------------------
' SINK-STEUERUNG (Ausgabekanaele)
' ---------------------------------------------------------------------------
Private m_blnSinkDebug      As Boolean  ' Debug.Print
Private m_blnSinkFile       As Boolean  ' Logfile
Private m_blnSinkDB         As Boolean  ' DB-Tabelle
Private m_blnSinkForm       As Boolean  ' StatusForm
Private m_blnInitialized    As Boolean  ' Init gelaufen?


' ---------------------------------------------------------------------------
' BATCH-VERWALTUNG
' ---------------------------------------------------------------------------
Private m_strBatchID        As String   ' Aktive Batch-ID (leer = kein Batch)


' ---------------------------------------------------------------------------
' SESSION-STATISTIK
' ---------------------------------------------------------------------------
Private m_lngErrors         As Long
Private m_lngWarnings       As Long
Private m_lngTotal          As Long


' ---------------------------------------------------------------------------
' DEV-MODE & STATUSFORM
' ---------------------------------------------------------------------------
Private m_blnDevMode        As Boolean
Private m_objStatusForm     As Object   ' Direkte Formular-Referenz
Private m_strFormName       As String   ' Formularname (fuer Lazy-Binding)
Private m_blnFormAutoOpen   As Boolean  ' Form automatisch oeffnen?


' ---------------------------------------------------------------------------
' DB-SINK ZUSTAND
' ---------------------------------------------------------------------------
Private m_blnDBTableChecked As Boolean  ' Einmal-Check ob tblLog existiert
Private m_blnDBTableExists  As Boolean


' ===========================================================================
' 1. INITIALISIERUNG
' ===========================================================================

Public Sub InitLogging(Optional ByVal blnDevMode As Boolean = False)
    If m_blnInitialized Then Exit Sub

    m_blnSinkDebug = True
    m_blnSinkFile = True
    m_blnSinkDB = False
    m_blnSinkForm = False
    m_blnDevMode = blnDevMode
    m_blnFormAutoOpen = False
    m_lngErrors = 0
    m_lngWarnings = 0
    m_lngTotal = 0

    m_blnInitialized = True
End Sub


' ===========================================================================
' 2. ABWAERTSKOMPATIBLE API (Signaturen 1:1 unveraendert)
' ===========================================================================

Public Sub LogInfo(ByVal strMsg As String, Optional ByVal strKat As String = "")
    SchreibeLog strMsg, LOG_INFO, strKat
End Sub

Public Sub LogWarn(ByVal strMsg As String, Optional ByVal strKat As String = "")
    SchreibeLog strMsg, LOG_WARN, strKat
End Sub

Public Sub LogError(ByVal strMsg As String, Optional ByVal strKat As String = "")
    SchreibeLog strMsg, LOG_ERROR, strKat
End Sub

Public Sub LogDebug(ByVal strMsg As String, Optional ByVal strKat As String = "")
    SchreibeLog strMsg, LOG_DEBUG, strKat
End Sub

Public Sub LogTrace(ByVal strMsg As String, Optional ByVal strKat As String = "")
    SchreibeLog strMsg, LOG_TRACE, strKat
End Sub

Public Sub LogVBAError(ByVal strKontext As String, _
                       Optional ByVal lngErrNum As Long = 0, _
                       Optional ByVal strErrDesc As String = "")
    If lngErrNum = 0 Then lngErrNum = Err.Number
    If strErrDesc = "" Then strErrDesc = Err.Description

    SchreibeLog strKontext & " - Err " & lngErrNum & ": " & strErrDesc, LOG_ERROR, ""
End Sub


' ===========================================================================
' 3. ZENTRALE LOG-FUNKTION (erweitert, abwaertskompatibel)
' ===========================================================================

Public Sub SchreibeLog(ByVal strNachricht As String, _
                       Optional ByVal iLevel As Integer = 3, _
                       Optional ByVal strKategorie As String = "")

    ' Auto-Init bei erstem Aufruf
    If Not m_blnInitialized Then InitLogging

    ' Log-Level pruefen
    If iLevel > g_intLogLevel Then Exit Sub

    ' Statistik
    m_lngTotal = m_lngTotal + 1
    If iLevel = LOG_ERROR Then m_lngErrors = m_lngErrors + 1
    If iLevel = LOG_WARN Then m_lngWarnings = m_lngWarnings + 1

    ' Eintrag formatieren
    Dim strEintrag As String
    If strKategorie <> "" Then
        strEintrag = Format$(Now, "hh:nn:ss") & " " & LevelTag(iLevel) & _
                     " [" & strKategorie & "] " & strNachricht
    Else
        strEintrag = Format$(Now, "hh:nn:ss") & " " & LevelTag(iLevel) & _
                     " " & strNachricht
    End If

    ' Sink 1: Debug.Print
    If m_blnSinkDebug Then Debug.Print strEintrag

    ' Sink 2: Logfile
    If m_blnSinkFile Then SinkDatei strEintrag

    ' Sink 3: DB (wenn aktiv)
    If m_blnSinkDB Then SinkDB "", "", strNachricht, 0, strKategorie, iLevel

    ' Sink 4: StatusForm (wenn aktiv)
    If m_blnSinkForm Then SinkForm strEintrag, iLevel
End Sub


' ===========================================================================
' 4. STRUKTURIERTE API (fuer neuen / refaktorierten Code)
' ===========================================================================

' LogEvent - Strukturierter Logeintrag mit vollem Kontext
' Fuer neuen Code: Modul + Prozedur + ggf. Fehler-Nr + Kontext
Public Sub LogEvent(ByVal strModul As String, _
                    ByVal strProzedur As String, _
                    ByVal strText As String, _
                    Optional ByVal lngFehlerNr As Long = 0, _
                    Optional ByVal strKontext As String = "", _
                    Optional ByVal iLevel As Integer = 3)

    If Not m_blnInitialized Then InitLogging
    If iLevel > g_intLogLevel Then Exit Sub

    ' Statistik
    m_lngTotal = m_lngTotal + 1
    If iLevel = LOG_ERROR Then m_lngErrors = m_lngErrors + 1
    If iLevel = LOG_WARN Then m_lngWarnings = m_lngWarnings + 1

    ' Formatieren: "hh:nn:ss [LEVEL] [Modul.Proc] Text (Err X) | Kontext"
    Dim strEintrag As String
    strEintrag = Format$(Now, "hh:nn:ss") & " " & LevelTag(iLevel) & _
                 " [" & strModul & "." & strProzedur & "] " & strText
    If lngFehlerNr <> 0 Then strEintrag = strEintrag & " (Err " & lngFehlerNr & ")"
    If strKontext <> "" Then strEintrag = strEintrag & " | " & strKontext

    ' Sinks
    If m_blnSinkDebug Then Debug.Print strEintrag
    If m_blnSinkFile Then SinkDatei strEintrag
    If m_blnSinkDB Then SinkDB strModul, strProzedur, strText, lngFehlerNr, strKontext, iLevel
    If m_blnSinkForm Then SinkForm strEintrag, iLevel
End Sub


' Nur im DevMode aktiv — zero-cost in Produktion
Public Sub LogDev(ByVal strMsg As String, Optional ByVal strKat As String = "")
    If Not m_blnDevMode Then Exit Sub
    SchreibeLog strMsg, LOG_DEBUG, strKat
End Sub


' SQL-Statements im DevMode tracen
Public Sub LogSQL(ByVal strSql As String, Optional ByVal strKontext As String = "SQL")
    If Not m_blnDevMode Then Exit Sub
    SchreibeLog "SQL: " & Left$(strSql, 500), LOG_TRACE, strKontext
End Sub


' ===========================================================================
' 5. BATCH-VERWALTUNG (gruppierte Operationen)
' ===========================================================================

Public Function StartLogBatch(Optional ByVal strPrefix As String = "BATCH") As String
    m_strBatchID = UCase$(strPrefix) & "_" & Format$(Now, "yyyymmdd_hhnnss")
    LogInfo "Batch gestartet: " & m_strBatchID
    StartLogBatch = m_strBatchID
End Function

Public Sub EndLogBatch()
    If m_strBatchID <> "" Then
        LogInfo "Batch beendet: " & m_strBatchID & " | " & GetLogStats()
        m_strBatchID = ""
    End If
End Sub

Public Function GetLogBatchID() As String
    GetLogBatchID = m_strBatchID
End Function


' ===========================================================================
' 6. SESSION-STATISTIK
' ===========================================================================

Public Function GetLogStats() As String
    GetLogStats = "E:" & m_lngErrors & _
                  " W:" & m_lngWarnings & _
                  " Total:" & m_lngTotal
End Function

Public Function GetLogErrorCount() As Long
    GetLogErrorCount = m_lngErrors
End Function

Public Function GetLogWarningCount() As Long
    GetLogWarningCount = m_lngWarnings
End Function

Public Sub ResetLogStats()
    m_lngErrors = 0
    m_lngWarnings = 0
    m_lngTotal = 0
End Sub


' ===========================================================================
' 7. KONFIGURATION & DEV-MODE
' ===========================================================================

Public Sub LogSetLevel(ByVal iLevel As Integer)
    g_intLogLevel = iLevel
End Sub

Public Sub LogSetDevMode(ByVal blnEnabled As Boolean)
    m_blnDevMode = blnEnabled
    If blnEnabled Then
        g_intLogLevel = LOG_TRACE
        LogInfo "=== DEV-MODE ON (TRACE) ==="
    Else
        g_intLogLevel = LOG_INFO
        LogInfo "=== DEV-MODE OFF (INFO) ==="
    End If
End Sub

Public Function IsDevMode() As Boolean
    IsDevMode = m_blnDevMode
End Function

Public Sub LogEnableDBSink(ByVal blnEnabled As Boolean)
    m_blnSinkDB = blnEnabled
End Sub

Public Sub LogEnableFileSink(ByVal blnEnabled As Boolean)
    m_blnSinkFile = blnEnabled
End Sub

Public Sub LogEnableDebugSink(ByVal blnEnabled As Boolean)
    m_blnSinkDebug = blnEnabled
End Sub

' StatusForm: Name setzen (Lazy-Binding ueber Forms()-Collection)
Public Sub LogSetStatusFormName(ByVal strName As String, _
                                Optional ByVal blnAutoOpen As Boolean = False)
    m_strFormName = strName
    m_blnFormAutoOpen = blnAutoOpen
    m_blnSinkForm = (Len(strName) > 0)
End Sub

' StatusForm: Direkte Objektreferenz setzen
Public Sub LogSetStatusForm(ByVal objForm As Object)
    Set m_objStatusForm = objForm
    m_blnSinkForm = Not (objForm Is Nothing)
End Sub

' Diagnose: Aktuellen Status im Direktbereich anzeigen (Strg+G ? LogStatus)
Public Sub LogStatus()
    Debug.Print String$(50, "-")
    Debug.Print "=== LOGGING STATUS ==="
    Debug.Print "Level:      " & g_intLogLevel & " (" & LevelName(g_intLogLevel) & ")"
    Debug.Print "DevMode:    " & IIf(m_blnDevMode, "ON", "OFF")
    Debug.Print "Sink Debug: " & IIf(m_blnSinkDebug, "ON", "OFF")
    Debug.Print "Sink File:  " & IIf(m_blnSinkFile, "ON", "OFF")
    Debug.Print "Sink DB:    " & IIf(m_blnSinkDB, "ON", "OFF") & _
                IIf(m_blnSinkDB And Not m_blnDBTableExists, " (Tabelle fehlt!)", "")
    Debug.Print "Sink Form:  " & IIf(m_blnSinkForm, "ON", "OFF") & _
                IIf(m_strFormName <> "", " (" & m_strFormName & ")", "")
    Debug.Print "Batch:      " & IIf(m_strBatchID <> "", m_strBatchID, "(keiner)")
    Debug.Print "Statistik:  " & GetLogStats()
    Debug.Print String$(50, "-")
End Sub


' ===========================================================================
' 8. FORTSCHRITTS-STEUERUNG (Vorbereitung fuer StatusForm)
' ===========================================================================

Public Function LogProgress_Start(Optional ByVal strMsg As String = "Start") As Boolean
    On Error Resume Next
    Dim frm As Object
    Set frm = GetFormRef()
    If frm Is Nothing Then Exit Function
    frm.StartProgress strMsg
    LogProgress_Start = (Err.Number = 0)
End Function

Public Function LogProgress_Step(ByVal lngCurrent As Long, ByVal lngTotal As Long, _
                                 Optional ByVal strMsg As String = "") As Boolean
    On Error Resume Next
    Dim frm As Object
    Set frm = GetFormRef()
    If frm Is Nothing Then Exit Function
    frm.SetProgressSteps lngCurrent, lngTotal, strMsg
    LogProgress_Step = (Err.Number = 0)
End Function

Public Function LogProgress_Status(ByVal strMsg As String) As Boolean
    On Error Resume Next
    Dim frm As Object
    Set frm = GetFormRef()
    If frm Is Nothing Then Exit Function
    frm.SetStatus strMsg
    LogProgress_Status = (Err.Number = 0)
End Function

Public Function LogProgress_IsCancelled() As Boolean
    On Error Resume Next
    Dim frm As Object
    Set frm = GetFormRef()
    If frm Is Nothing Then Exit Function
    LogProgress_IsCancelled = frm.IsCancelled
End Function


' ===========================================================================
' 9. INTERNE SINKS (Private)
' ===========================================================================

Private Sub SinkDatei(ByVal strText As String)
    On Error Resume Next
    Dim strPath As String
    Dim f       As Integer

    strPath = CurrentProject.Path & "\" & LOG_FILE_PREFIX & _
              Format$(Date, "yyyymmdd") & ".txt"
    f = FreeFile
    Open strPath For Append As #f
    Print #f, strText
    Close #f
End Sub


Private Sub SinkDB(ByVal strModul As String, _
                   ByVal strProzedur As String, _
                   ByVal strText As String, _
                   ByVal lngFehlerNr As Long, _
                   ByVal strKontext As String, _
                   ByVal iLevel As Integer)
    On Error Resume Next

    ' Einmal pruefen ob Tabelle existiert (ueber Cache)
    If Not m_blnDBTableChecked Then
        m_blnDBTableExists = CacheTabelleExistiert(TBL_LOG)
        m_blnDBTableChecked = True
        If Not m_blnDBTableExists Then Exit Sub
    End If
    If Not m_blnDBTableExists Then Exit Sub

    Dim sSQL As String
    sSQL = "INSERT INTO [" & TBL_LOG & "] " & _
           "(LogZeit, LogLevel, Modul, Prozedur, FehlerNr, " & _
           "Nachricht, Kontext, BatchID, UserName) VALUES (" & _
           "Now(), " & iLevel & ", " & _
           "'" & Left$(Esc(strModul), 100) & "', " & _
           "'" & Left$(Esc(strProzedur), 100) & "', " & _
           lngFehlerNr & ", " & _
           "'" & Left$(Esc(strText), 255) & "', " & _
           "'" & Left$(Esc(strKontext), 255) & "', " & _
           "'" & Left$(Esc(m_strBatchID), 50) & "', " & _
           "'" & Left$(Esc(Environ$("USERNAME")), 50) & "')"

    CurrentDb.Execute sSQL, dbFailOnError
End Sub


Private Sub SinkForm(ByVal strEintrag As String, ByVal iLevel As Integer)
    On Error Resume Next

    Dim frm As Object
    Set frm = GetFormRef()
    If frm Is Nothing Then Exit Sub

    frm.AddLog strEintrag, LevelName(iLevel)

    ' Bei Fehler/Warnung auch Statuszeile aktualisieren
    If iLevel <= LOG_WARN Then
        frm.SetStatus strEintrag
    End If
End Sub


' ===========================================================================
' 10. HELPER
' ===========================================================================

Private Function LevelTag(ByVal iLevel As Integer) As String
    Select Case iLevel
        Case LOG_ERROR: LevelTag = "[ERROR]"
        Case LOG_WARN:  LevelTag = "[WARN ]"
        Case LOG_INFO:  LevelTag = "[INFO ]"
        Case LOG_DEBUG: LevelTag = "[DEBUG]"
        Case LOG_TRACE: LevelTag = "[TRACE]"
        Case Else:      LevelTag = "[     ]"
    End Select
End Function

Private Function LevelName(ByVal iLevel As Integer) As String
    Select Case iLevel
        Case LOG_ERROR: LevelName = "ERROR"
        Case LOG_WARN:  LevelName = "WARN"
        Case LOG_INFO:  LevelName = "INFO"
        Case LOG_DEBUG: LevelName = "DEBUG"
        Case LOG_TRACE: LevelName = "TRACE"
        Case Else:      LevelName = "LOG"
    End Select
End Function

Private Function Esc(ByVal s As String) As String
    Esc = Replace(s, "'", "''")
End Function

' StatusForm-Referenz aufloesen (Lazy-Binding)
Private Function GetFormRef() As Object
    On Error Resume Next

    ' Direkte Referenz bevorzugen
    If Not m_objStatusForm Is Nothing Then
        Set GetFormRef = m_objStatusForm
        Exit Function
    End If

    ' Per Formularname
    If m_strFormName = "" Then Exit Function

    ' Pruefen ob geladen
    If (SysCmd(acSysCmdGetObjectState, acForm, m_strFormName) _
        And acObjStateOpen) <> 0 Then
        Set GetFormRef = Forms(m_strFormName)
        Exit Function
    End If

    ' AutoOpen wenn konfiguriert
    If m_blnFormAutoOpen Then
        DoCmd.OpenForm m_strFormName, , , , , acWindowNormal
        If Err.Number = 0 Then
            Set GetFormRef = Forms(m_strFormName)
        End If
    End If
End Function


' ===========================================================================
' 11. DB-TABELLE ERSTELLEN (einmalig aus Direktbereich: ErstelleLogTabelle)
' ===========================================================================

Public Sub ErstelleLogTabelle()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Set db = CurrentDb

    ' Pruefen ob schon vorhanden
    On Error Resume Next
    Dim strTest As String
    strTest = db.TableDefs(TBL_LOG).Name
    If Err.Number = 0 Then
        Debug.Print TBL_LOG & " existiert bereits."
        Exit Sub
    End If
    Err.Clear
    On Error GoTo ErrHandler

    db.Execute "CREATE TABLE [" & TBL_LOG & "] (" & _
               "[ID] AUTOINCREMENT PRIMARY KEY, " & _
               "[LogZeit] DATETIME NOT NULL, " & _
               "[LogLevel] INTEGER NOT NULL, " & _
               "[Modul] TEXT(100), " & _
               "[Prozedur] TEXT(100), " & _
               "[FehlerNr] LONG, " & _
               "[Nachricht] TEXT(255), " & _
               "[Kontext] TEXT(255), " & _
               "[BatchID] TEXT(50), " & _
               "[UserName] TEXT(50)" & _
               ")", dbFailOnError

    ' Indizes fuer schnelle Abfragen
    db.Execute "CREATE INDEX IX_LogZeit ON [" & TBL_LOG & _
               "] (LogZeit)", dbFailOnError
    db.Execute "CREATE INDEX IX_LogLevel ON [" & TBL_LOG & _
               "] (LogLevel)", dbFailOnError

    m_blnDBTableChecked = True
    m_blnDBTableExists = True
    Debug.Print TBL_LOG & " erfolgreich erstellt."
    Exit Sub

ErrHandler:
    Debug.Print "FEHLER ErstelleLogTabelle: " & Err.Number & " - " & Err.Description
End Sub


