Option Compare Database
Option Explicit

' ===========================================================================
' modDualAccessWorker - No-Admin Dual-Access Queue/Worker Runtime
' ===========================================================================
' Enthaltet die Basiskomponenten fuer die 2-Prozess-Strategie:
'   - SetupDualAccessNoAdmin   -> Tabellen/Formulare fuer Queue-Betrieb sicherstellen
'   - WorkerStarten            -> Zweite Access-Instanz als Worker starten
'   - ClaimNextJob             -> Atomarer Job-Claim (queued -> running)
'   - HeartbeatUpdate          -> Progress/Lease aktualisieren
'   - HandlePauseCancel        -> Pause/Cancel-Steuerung pro Job
'   - FinalizeJob              -> Job sauber abschliessen
' ===========================================================================

Private Const MODUL_NAME As String = "modDualAccessWorker"
Private Const FORM_SYNC_JOB As String = "frmSyncJobQueue"
Private Const FORM_SYNC_HEARTBEAT As String = "frmSyncHeartbeat"
Private Const FORM_WORKER_LEASE As String = "frmWorkerLease"

Private m_strAktiverWorkerId As String

#If VBA7 Then
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
#Else
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If

Private Const SW_HIDE As Long = 0
Private Const SW_MINIMIZE As Long = 6

Private Const WORKER_SYNC_LOOKUP_RETRIES As Long = 10
Private Const WORKER_SYNC_LOOKUP_SLEEP_MS As Long = 300

Private m_blnWorkerTraceChecked As Boolean
Private m_blnWorkerTraceAvailable As Boolean


' ---------------------------------------------------------------------------
' SETUP: Dual-Access No-Admin vorbereiten
' ---------------------------------------------------------------------------
Public Function SetupDualAccessNoAdmin(Optional ByVal blnCreateForms As Boolean = False, _
                                       Optional ByVal blnRequireWorkerMacro As Boolean = False) As Boolean
    On Error GoTo ErrHandler

    Dim blnOK As Boolean
    Dim strSetupDetail As String

    If DualAccessBereitOhneMigration() Then
        InitStandardConfig

        If blnCreateForms Then
            If Not EnsureDualAccessForms() Then
                LogWarn "Dual-Access Formulare konnten nicht komplett erstellt werden", MODUL_NAME
            End If
        End If

        SetupDualAccessNoAdmin = True
        LogInfo "Dual-Access Setup bereits bereit", MODUL_NAME
        Exit Function
    End If

    ' Standard-Config sicherstellen (inkl. Worker-Keys).
    ' Ist idempotent und setzt fehlende Schluessel nach.
    InitStandardConfig

    blnOK = EnsureDualAccessBasis(strSetupDetail)
    If Not blnOK And Trim$(strSetupDetail) <> "" Then
        LogWarn "Dual-Access Basis-Setup unvollstaendig: " & strSetupDetail, MODUL_NAME
        Debug.Print "[SETUP] Dual-Access Basis FEHLER | " & strSetupDetail
    End If

    If Not MakroExistiert("AutoStartWorker") Then
        LogWarn "Makro 'AutoStartWorker' fehlt (WorkerStarten via /x noch nicht verfuegbar)", MODUL_NAME
        If blnRequireWorkerMacro Then
            blnOK = False
        End If
    End If

    If blnCreateForms Then
        If Not EnsureDualAccessForms() Then
            LogWarn "Dual-Access Formulare konnten nicht komplett erstellt werden", MODUL_NAME
        End If
    End If

    SetupDualAccessNoAdmin = blnOK
    If blnOK Then
        LogInfo "Dual-Access Setup bereit", MODUL_NAME
    Else
        LogWarn "Dual-Access Setup unvollstaendig (Details siehe Warnungen)", MODUL_NAME
    End If
    Exit Function

ErrHandler:
    HandleError MODUL_NAME, "SetupDualAccessNoAdmin"
    SetupDualAccessNoAdmin = False
End Function

Private Function EnsureDualAccessBasis(Optional ByRef strDetail As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim strBackendPfad As String

    strDetail = ""
    strBackendPfad = Trim$(CacheGetConfig(CFG_BACKEND_PFAD, ""))

    If strBackendPfad <> "" Then
        If Not EnsureDualAccessBackendBereit(strBackendPfad, strDetail) Then
            EnsureDualAccessBasis = False
            Exit Function
        End If
    Else
        If Not EnsureDualAccessLokaleTabellen(strDetail) Then
            EnsureDualAccessBasis = False
            Exit Function
        End If
    End If

    EnsureDualAccessBasis = EnsureDualAccessPflichttabellen(strDetail)
    Exit Function

ErrHandler:
    strDetail = "EnsureDualAccessBasis: Err " & Err.Number & " - " & Err.Description
    EnsureDualAccessBasis = False
End Function

Private Function EnsureDualAccessBackendBereit(ByVal strBackendPfad As String, _
                                              Optional ByRef strDetail As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim blnOK As Boolean

    strDetail = ""
    Set db = DBEngine.OpenDatabase(strBackendPfad)
    blnOK = ErstelleBackendTabellenInDB(db)
    db.Close
    Set db = Nothing

    If Not blnOK Then
        strDetail = "Backend-Tabellen konnten nicht sichergestellt werden: " & strBackendPfad
        EnsureDualAccessBackendBereit = False
        Exit Function
    End If

    blnOK = EnsureDualAccessLink(TBL_SYNC_JOB, strBackendPfad, strDetail)
    blnOK = blnOK And EnsureDualAccessLink(TBL_SYNC_HEARTBEAT, strBackendPfad, strDetail)
    blnOK = blnOK And EnsureDualAccessLink(TBL_SYNC_CONTROL, strBackendPfad, strDetail)
    blnOK = blnOK And EnsureDualAccessLink(TBL_WORKER_LEASE, strBackendPfad, strDetail)
    blnOK = blnOK And EnsureDualAccessLink(TBL_WORKER_TRACE, strBackendPfad, strDetail)

    m_blnWorkerTraceChecked = False
    m_blnWorkerTraceAvailable = False
    EnsureDualAccessBackendBereit = blnOK
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    strDetail = "EnsureDualAccessBackendBereit: Err " & Err.Number & " - " & Err.Description
    EnsureDualAccessBackendBereit = False
End Function

Private Function EnsureDualAccessLokaleTabellen(Optional ByRef strDetail As String = "") As Boolean
    On Error GoTo ErrHandler

    strDetail = ""

    If Not TabelleExistiert(TBL_SYNC_JOB) Then
        CurrentDb.Execute "CREATE TABLE [" & TBL_SYNC_JOB & "] (" & _
                          "JobID AUTOINCREMENT CONSTRAINT PK_SyncJob PRIMARY KEY, " & _
                          "CreatedAt DATETIME, CreatedBy TEXT(100), RequestedFolderPath TEXT(255), " & _
                          "RequestedMaxMails LONG, RequestedSubfolders YESNO, Status TEXT(30), " & _
                          "WorkerId TEXT(100), StartedAt DATETIME, FinishedAt DATETIME, LastError MEMO, Priority INTEGER)"
        CurrentDb.Execute "CREATE INDEX idx_SyncJob_Status ON [" & TBL_SYNC_JOB & "] (Status)"
        CurrentDb.Execute "CREATE INDEX idx_SyncJob_CreatedAt ON [" & TBL_SYNC_JOB & "] (CreatedAt)"
        CurrentDb.Execute "CREATE INDEX idx_SyncJob_PrioCreated ON [" & TBL_SYNC_JOB & "] (Priority, CreatedAt)"
    End If

    If Not TabelleExistiert(TBL_SYNC_HEARTBEAT) Then
        CurrentDb.Execute "CREATE TABLE [" & TBL_SYNC_HEARTBEAT & "] (" & _
                          "WorkerId TEXT(100) CONSTRAINT PK_SyncHeartbeat PRIMARY KEY, JobID LONG, Stage TEXT(50), " & _
                          "CurrentItem LONG, TotalItems LONG, COMRetries LONG, COMReconnects LONG, UpdatedAt DATETIME, LastMessage MEMO)"
        CurrentDb.Execute "CREATE INDEX idx_SyncHB_UpdatedAt ON [" & TBL_SYNC_HEARTBEAT & "] (UpdatedAt)"
        CurrentDb.Execute "CREATE INDEX idx_SyncHB_JobID ON [" & TBL_SYNC_HEARTBEAT & "] (JobID)"
    End If

    If Not TabelleExistiert(TBL_SYNC_CONTROL) Then
        CurrentDb.Execute "CREATE TABLE [" & TBL_SYNC_CONTROL & "] (" & _
                          "JobID LONG CONSTRAINT PK_SyncControl PRIMARY KEY, PauseRequested YESNO, CancelRequested YESNO, UpdatedAt DATETIME)"
    End If

    If Not TabelleExistiert(TBL_WORKER_LEASE) Then
        CurrentDb.Execute "CREATE TABLE [" & TBL_WORKER_LEASE & "] (" & _
                          "WorkerId TEXT(100) CONSTRAINT PK_WorkerLease PRIMARY KEY, LeaseUntil DATETIME, UpdatedAt DATETIME, HostName TEXT(100), SessionUser TEXT(100))"
        CurrentDb.Execute "CREATE INDEX idx_WorkerLease_Until ON [" & TBL_WORKER_LEASE & "] (LeaseUntil)"
    End If

    EnsureDualAccessLokaleTabellen = WorkerTraceTabelleSicher() And EnsureDualAccessPflichttabellen(strDetail)
    Exit Function

ErrHandler:
    strDetail = "EnsureDualAccessLokaleTabellen: Err " & Err.Number & " - " & Err.Description
    EnsureDualAccessLokaleTabellen = False
End Function

Private Function EnsureDualAccessPflichttabellen(Optional ByRef strDetail As String = "") As Boolean
    Dim strFehlt As String

    strDetail = ""
    If Not TabelleExistiert(TBL_SYNC_JOB) Then strFehlt = strFehlt & TBL_SYNC_JOB & ", "
    If Not TabelleExistiert(TBL_SYNC_HEARTBEAT) Then strFehlt = strFehlt & TBL_SYNC_HEARTBEAT & ", "
    If Not TabelleExistiert(TBL_SYNC_CONTROL) Then strFehlt = strFehlt & TBL_SYNC_CONTROL & ", "
    If Not TabelleExistiert(TBL_WORKER_LEASE) Then strFehlt = strFehlt & TBL_WORKER_LEASE & ", "
    If Not WorkerTraceTabelleSicher() Then strFehlt = strFehlt & TBL_WORKER_TRACE & ", "
    If Len(Nz(CacheGetConfig(CFG_WORKER_POLL_MS, ""), "")) = 0 Then strFehlt = strFehlt & CFG_WORKER_POLL_MS & ", "
    If Len(Nz(CacheGetConfig(CFG_WORKER_HB_S, ""), "")) = 0 Then strFehlt = strFehlt & CFG_WORKER_HB_S & ", "
    If Len(Nz(CacheGetConfig(CFG_WORKER_STALE_S, ""), "")) = 0 Then strFehlt = strFehlt & CFG_WORKER_STALE_S & ", "

    If strFehlt <> "" Then
        strDetail = "Fehlt: " & Left$(strFehlt, Len(strFehlt) - 2)
        EnsureDualAccessPflichttabellen = False
    Else
        EnsureDualAccessPflichttabellen = True
    End If
End Function

Private Function EnsureDualAccessLink(ByVal strTabelle As String, _
                                      ByVal strBackendPfad As String, _
                                      Optional ByRef strDetail As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim strConnect As String

    strDetail = ""
    If Trim$(strBackendPfad) = "" Then Exit Function
    If Not TabelleExistiertInBackend(strTabelle, strBackendPfad) Then
        strDetail = "Backend-Tabelle fehlt: " & strTabelle
        EnsureDualAccessLink = False
        Exit Function
    End If

    strConnect = ";DATABASE=" & strBackendPfad
    Set db = CurrentDb

    If TabelleExistiert(strTabelle) Then
        If IstLinkedTable(strTabelle) Then
            If StrComp(Nz(GetConnectString(strTabelle), ""), strConnect, vbTextCompare) = 0 Then
                On Error Resume Next
                db.TableDefs(strTabelle).RefreshLink
                EnsureDualAccessLink = (Err.Number = 0)
                Err.Clear
                On Error GoTo ErrHandler
                If EnsureDualAccessLink Then GoTo CleanExit
            End If

            On Error Resume Next
            db.TableDefs.Delete strTabelle
            db.TableDefs.Refresh
            Err.Clear
            On Error GoTo ErrHandler
        Else
            On Error Resume Next
            db.Execute "INSERT INTO [" & strTabelle & "] IN '" & Replace(strBackendPfad, "'", "''") & "' SELECT * FROM [" & strTabelle & "]"
            Err.Clear
            db.Execute "DROP TABLE [" & strTabelle & "]"
            db.TableDefs.Refresh
            Err.Clear
            On Error GoTo ErrHandler
        End If
    End If

    Set td = db.CreateTableDef(strTabelle)
    td.Connect = strConnect
    td.SourceTableName = strTabelle
    db.TableDefs.Append td
    db.TableDefs.Refresh

    EnsureDualAccessLink = TabelleExistiert(strTabelle) And IstLinkedTable(strTabelle)

CleanExit:
    Set td = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    strDetail = "EnsureDualAccessLink(" & strTabelle & "): Err " & Err.Number & " - " & Err.Description
    On Error Resume Next
    Set td = Nothing
    Set db = Nothing
    EnsureDualAccessLink = False
End Function

Private Function TabelleExistiertInBackend(ByVal strTabelle As String, ByVal strBackendPfad As String) As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef

    Set db = DBEngine.OpenDatabase(strBackendPfad)
    For Each td In db.TableDefs
        If StrComp(td.Name, strTabelle, vbTextCompare) = 0 Then
            TabelleExistiertInBackend = True
            Exit For
        End If
    Next td

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set td = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    TabelleExistiertInBackend = False
    Resume CleanExit
End Function


' ---------------------------------------------------------------------------
' UI: Worker-Prozess starten (ohne Admin/PowerShell)
' ---------------------------------------------------------------------------
Public Function WorkerStarten(Optional ByVal strWorkerDbPath As String = "", _
                              Optional ByVal strWorkerId As String = "", _
                              Optional ByVal lngWaitSek As Long = 10, _
                              Optional ByRef lngOutTaskID As Long = 0, _
                              Optional ByVal blnRuntimeMode As Boolean = True, _
                              Optional ByVal blnHideWindow As Boolean = True) As Boolean
    On Error GoTo ErrHandler

    Dim strAccessExe As String
    Dim strCmd As String
    Dim lngTaskID As Long
    Dim strLaunchDbPath As String
    Dim intWindowStyle As VbAppWinStyle
    Dim blnLaunchArtefaktAngelegt As Boolean

    If strWorkerDbPath = "" Then strWorkerDbPath = CurrentDb.Name
    If strWorkerId = "" Then strWorkerId = WorkerIdStandard()

    If WorkerIstOnline(strWorkerId) Then
        LogWarn "Worker bereits online, kein zweiter Start: " & strWorkerId, MODUL_NAME
        Debug.Print "[WORKER] Start uebersprungen | Grund=bereits online | WorkerId=" & strWorkerId
        lngOutTaskID = 0
        WorkerStarten = False
        Exit Function
    End If

    ' Alte Stop-Signale wegraeumen, damit ein neuer Start nicht sofort stoppt.
    WorkerStopSignalLoeschen strWorkerId

    Dim strGrund As String
    If Not WorkerStartVoraussetzungenOK(strGrund, strWorkerDbPath) Then
        LogError "Worker-Start Voraussetzungen fehlen: " & strGrund, MODUL_NAME
        Debug.Print "[WORKER] Start abgebrochen | Grund=" & strGrund
        WorkerStarten = False
        Exit Function
    End If

    If Not WorkerLaunchDbVorbereiten(strWorkerDbPath, strWorkerId, strLaunchDbPath, strGrund) Then
        LogError "Worker-Start Vorbereitung fehlgeschlagen: " & strGrund, MODUL_NAME
        Debug.Print "[WORKER] Start abgebrochen | Grund=" & strGrund
        WorkerStarten = False
        Exit Function
    End If
    blnLaunchArtefaktAngelegt = (Trim$(strLaunchDbPath) <> "")

    strAccessExe = HoleAccessExePfad()
    If strAccessExe = "" Then
        LogError "MSACCESS.EXE nicht gefunden", MODUL_NAME
        WorkerStarten = False
        Exit Function
    End If

    ' Transparenz: effektive Speicher-/Prozesspfade vor Worker-Start ausgeben.
    DevPfadDiagnose "WorkerStarten", False, strLaunchDbPath

    strCmd = Quote(strAccessExe) & " " & Quote(strLaunchDbPath) & _
             " /x AutoStartWorker /cmd " & Quote(strWorkerId)

    If blnRuntimeMode Then
        strCmd = strCmd & " /runtime"
    End If

    ' Kein Startup-UI im Worker: reduziert Fokuswechsel/Fensterflackern.
    strCmd = strCmd & " /nostartup"

    If blnHideWindow Then
        intWindowStyle = vbHide
    Else
        intWindowStyle = vbMinimizedNoFocus
    End If

    Debug.Print "[WORKER] Startmodus | WorkerId=" & strWorkerId & _
                " | WindowMode=" & IIf(blnHideWindow, "hide", "minimized") & _
                " | Runtime=" & IIf(blnRuntimeMode, "on", "off") & _
                " | NoStartup=on"

    lngTaskID = Shell(strCmd, intWindowStyle)
    If lngTaskID <= 0 Then
        LogError "Worker-Shell lieferte ungueltige TaskID (<=0): " & strWorkerId, MODUL_NAME
        Debug.Print "[WORKER] Shell-Start FEHLER | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID
        lngOutTaskID = 0
        WorkerStarten = False
        Exit Function
    End If

    lngOutTaskID = lngTaskID
    LogInfo "Worker-Shell gestartet: " & strWorkerId & " (TaskID=" & lngTaskID & ") | DB=" & strLaunchDbPath, MODUL_NAME
    Debug.Print "[WORKER] Shell gestartet | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID & " | DB=" & strLaunchDbPath

    If lngWaitSek > 0 Then
        If WaitForWorkerOnline(strWorkerId, lngWaitSek) Then
            LogInfo "Worker online bestaetigt: " & strWorkerId, MODUL_NAME
            Debug.Print "[WORKER] ONLINE bestaetigt | WorkerId=" & strWorkerId
            WorkerStarten = True
        Else
            Dim lngGraceSek As Long
            lngGraceSek = 5

            ' Bei UNC/Netz-Latenz kann der erste Heartbeat knapp nach dem Timeout eintreffen.
            If WaitForWorkerOnline(strWorkerId, lngGraceSek) Then
                LogWarn "Worker online erst im Grace-Fenster bestaetigt: " & strWorkerId, MODUL_NAME
                Debug.Print "[WORKER] ONLINE spaet bestaetigt (+" & lngGraceSek & "s) | WorkerId=" & strWorkerId
                WorkerStarten = True
            Else
                LogWarn "Worker-Start nicht bestaetigt innerhalb " & lngWaitSek & "s (+" & lngGraceSek & "s Grace): " & strWorkerId, MODUL_NAME
                Debug.Print "[WORKER] Keine Online-Bestaetigung innerhalb " & lngWaitSek & "s (+" & lngGraceSek & "s Grace) | WorkerId=" & strWorkerId
                Debug.Print "[WORKER] Fehlstart-Cleanup startet | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID
                Call WorkerStoppen(strWorkerId, lngTaskID, 2)
                WorkerStarten = False
            End If
        End If
    Else
        WorkerStarten = True
    End If

    Exit Function

ErrHandler:
    HandleError MODUL_NAME, "WorkerStarten"
    If blnLaunchArtefaktAngelegt And lngTaskID <= 0 Then
        Debug.Print "[WORKER] Start-Fehler Cleanup | WorkerId=" & strWorkerId & " | DB=" & strLaunchDbPath
        Call WorkerLoescheDateiSicher(strWorkerId, strLaunchDbPath)
        Call WorkerLoescheDateiSicher(strWorkerId, WorkerLockDateiPfad(strLaunchDbPath))
    End If
    lngOutTaskID = 0
    WorkerStarten = False
End Function

Public Function WorkerStoppen(Optional ByVal strWorkerId As String = "", _
                              Optional ByVal lngTaskID As Long = 0, _
                              Optional ByVal lngGraceSek As Long = 5) As Boolean
    On Error GoTo ErrHandler

    Dim strLaunchDbPath As String
    Dim strGrund As String

    If Trim$(strWorkerId) <> "" Then
        strLaunchDbPath = WorkerLaunchDateiPfad(strWorkerId)
    Else
        strLaunchDbPath = ""
    End If

    If lngTaskID <= 0 And Trim$(strWorkerId) = "" Then
        Debug.Print "[WORKER] Stop uebersprungen | Grund=keine TaskID und keine WorkerId"
        WorkerStoppen = False
        Exit Function
    End If

    If lngGraceSek < 1 Then lngGraceSek = 1
    If lngGraceSek > 30 Then lngGraceSek = 30

    LogInfo "Worker-Stop angefordert: Worker=" & strWorkerId & " TaskID=" & lngTaskID, MODUL_NAME
    Debug.Print "[WORKER] Stop angefordert | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID

    If Trim$(strWorkerId) <> "" Then
        WorkerStopSignalSetzen strWorkerId
        Debug.Print "[WORKER] Stop-Signal gesetzt | WorkerId=" & strWorkerId
    End If

    If lngTaskID > 0 Then
        Debug.Print "[WORKER] Graceful-Stop wartet | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID & " | GraceSek=" & lngGraceSek
        If WorkerWarteAufExit(lngTaskID, lngGraceSek * 1000) Then
            WorkerBereinigeLaunchArtefakte strWorkerId, True
            Debug.Print "[WORKER] Stop bestaetigt (graceful) | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID
            WorkerStoppen = True
            Exit Function
        End If
    End If

    Debug.Print "[WORKER] Force-Stop eskaliert | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID & " | DB=" & strLaunchDbPath

    If lngTaskID > 0 Then WorkerStopTaskKill lngTaskID, True
    Call WorkerBeendeAccessProzesseFuerDatei(strWorkerId, strLaunchDbPath, lngTaskID, True)

    If (lngTaskID <= 0 Or WorkerWarteAufExit(lngTaskID, 3000)) And _
       WorkerStelleDateiFrei(strWorkerId, strLaunchDbPath, True, strGrund) Then
        WorkerBereinigeLaunchArtefakte strWorkerId, True
        Debug.Print "[WORKER] Stop bestaetigt (force) | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID
        WorkerStoppen = True
    Else
        Debug.Print "[WORKER] Stop FEHLER | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID & " | Grund=" & strGrund
        WorkerStoppen = False
    End If

    Exit Function

ErrHandler:
    HandleError MODUL_NAME, "WorkerStoppen", "WorkerId=" & strWorkerId & "|TaskID=" & lngTaskID
    WorkerStoppen = False
End Function

Public Function WorkerStartVoraussetzungenOK(Optional ByRef strGrund As String = "", _
                                            Optional ByVal strWorkerDbPath As String = "") As Boolean
    strGrund = ""

    If strWorkerDbPath = "" Then strWorkerDbPath = CurrentDb.Name

    If HoleAccessExePfad() = "" Then
        strGrund = "MSACCESS.EXE nicht gefunden"
        WorkerStartVoraussetzungenOK = False
        Exit Function
    End If

    If Not MakroExistiert("AutoStartWorker") Then
        strGrund = "Makro 'AutoStartWorker' fehlt"
        WorkerStartVoraussetzungenOK = False
        Exit Function
    End If

    If Not WorkerQuelleStartbereit(strWorkerDbPath, strGrund) Then
        WorkerStartVoraussetzungenOK = False
        Exit Function
    End If

    If Not WorkerTempOrdnerBereit(strGrund) Then
        WorkerStartVoraussetzungenOK = False
        Exit Function
    End If

    WorkerStartVoraussetzungenOK = True
End Function

Public Sub WorkerStartVoraussetzungenReport(Optional ByVal strWorkerDbPath As String = "")
    Dim strGrund As String
    Dim strAccessExe As String
    Dim strDbPruefung As String
    Dim strStartOrdnerPruefung As String
    strAccessExe = HoleAccessExePfad()

    If strWorkerDbPath = "" Then strWorkerDbPath = CurrentDb.Name
    If WorkerQuelleStartbereit(strWorkerDbPath, strDbPruefung) Then
        strDbPruefung = "OK"
    End If
    If WorkerTempOrdnerBereit(strStartOrdnerPruefung) Then
        strStartOrdnerPruefung = "OK"
    End If

    Debug.Print String(70, "-")
    Debug.Print "[WORKER-START-VORAUSSETZUNGEN]"
    Debug.Print "  Access EXE     : " & IIf(strAccessExe = "", "FEHLT", strAccessExe)
    Debug.Print "  Makro vorhanden: " & IIf(MakroExistiert("AutoStartWorker"), "JA", "Nein")
    Debug.Print "  Worker-Quelle  : " & strWorkerDbPath
    Debug.Print "  Quelle lesbar  : " & strDbPruefung
    Debug.Print "  Worker-Start   : " & WorkerLaunchOrdnerPfad()
    Debug.Print "  Startordner OK : " & strStartOrdnerPruefung
    Debug.Print "  Startmodus     : stabile Worker-FE pro Worker im Benutzerordner"

    If WorkerStartVoraussetzungenOK(strGrund, strWorkerDbPath) Then
        Debug.Print "  Ergebnis       : OK"
    Else
        Debug.Print "  Ergebnis       : FEHLT - " & strGrund
        If InStr(1, strGrund, "AutoStartWorker", vbTextCompare) > 0 Then
            Debug.Print ""
            Debug.Print "  Einrichtung in Access:"
            Debug.Print "  1. Erstellen -> Makro"
            Debug.Print "  2. Aktionszeile: AusfuehrenCode"
            Debug.Print "  3. Funktionsname: =AutoStartWorker()"
            Debug.Print "  4. Speichern als: AutoStartWorker"
            Debug.Print "  5. Smoke-Test erneut starten"
        ElseIf InStr(1, strGrund, "kopie", vbTextCompare) > 0 Or InStr(1, strGrund, "temp", vbTextCompare) > 0 Or InStr(1, strGrund, "quelle", vbTextCompare) > 0 Or InStr(1, strGrund, "dokument", vbTextCompare) > 0 Then
            Debug.Print ""
            Debug.Print "  Hinweis:"
            Debug.Print "  - Worker startet jetzt aus einer stabilen Worker-FE pro Worker im Benutzerordner Dokumente."
            Debug.Print "  - Wenn das scheitert, ist meist die Live-FE nicht lesbar oder der Zielordner nicht beschreibbar."
            Debug.Print "  - Offene Access-Dialoge schliessen und Smoke-Test erneut starten."
            Debug.Print "  - Danach Smoke-Test erneut starten."
        End If
    End If
    Debug.Print String(70, "-")
End Sub


' ---------------------------------------------------------------------------
' Worker-Entry (fuer Access-Makro /x AutoStartWorker via AusfuehrenCode)
' ---------------------------------------------------------------------------
Public Function AutoStartWorker() As Boolean
    On Error GoTo ErrHandler

    Dim strWorkerId As String
    strWorkerId = Trim$(Command$)
    If strWorkerId = "" Then strWorkerId = WorkerIdStandard()

    WorkerApplyBackgroundUiMode True
    WorkerLogTrace MODUL_NAME, "AutoStartWorker", "Worker-Autostart initialisiert", strWorkerId, 0, "Command=" & Trim$(Command$)

    AutoStartWorker = True
    ' Kein internes Quit aus laufendem VBA-Kontext, da das Access-Dialoge
    ' im Vordergrund triggern kann ("Sie koennen Access jetzt nicht beenden").
    WorkerPollLoop strWorkerId, 0, False
    Exit Function

ErrHandler:
    HandleError MODUL_NAME, "AutoStartWorker"
    AutoStartWorker = False
End Function


' ---------------------------------------------------------------------------
' Worker-Hauptschleife
' ---------------------------------------------------------------------------
Public Sub WorkerPollLoop(Optional ByVal strWorkerId As String = "", _
                          Optional ByVal lngMaxCycles As Long = 0, _
                          Optional ByVal blnSelfTerminate As Boolean = False)
    On Error GoTo ErrHandler

    If strWorkerId = "" Then strWorkerId = WorkerIdStandard()
    m_strAktiverWorkerId = strWorkerId

    Dim lngPollMs As Long
    lngPollMs = CLng(CacheGetConfig(CFG_WORKER_POLL_MS, "3000"))
    If lngPollMs < 500 Then lngPollMs = 500
    If lngPollMs > 60000 Then lngPollMs = 60000

    LogInfo "WorkerPollLoop gestartet: " & strWorkerId, MODUL_NAME
    Debug.Print "[WORKER] PollLoop gestartet | WorkerId=" & strWorkerId
    HeartbeatUpdate strWorkerId, 0, "boot", 0, 0, "worker boot"

    Dim lngCycle As Long
    Dim lngJobID As Long

    Do
        lngCycle = lngCycle + 1

        If lngCycle = 1 Or (lngCycle Mod 10) = 0 Then
            WorkerDebugPrint "Poll-Zyklus", strWorkerId, 0, "cycle=" & lngCycle & " | max=" & lngMaxCycles & " | abort=" & IIf(g_blnAbbrechen, "1", "0")
        End If

        If WorkerStopSignalAktiv(strWorkerId) Then
            Debug.Print "[WORKER] Shutdown-Signal erkannt | WorkerId=" & strWorkerId
            Exit Do
        End If

        lngJobID = ClaimNextJob(strWorkerId)
        If lngJobID > 0 Then
            WorkerDebugPrint "Job geclaimt, starte Ausfuehrung", strWorkerId, lngJobID
            ExecuteQueuedJob lngJobID, strWorkerId
        Else
            HeartbeatUpdate strWorkerId, 0, "idle", 0, 0, "idle"
            If WorkerSleepMitStopSignal(strWorkerId, lngPollMs, 200) Then
                Exit Do
            End If
        End If

        If lngMaxCycles > 0 Then
            If lngCycle >= lngMaxCycles Then Exit Do
        End If

        If g_blnAbbrechen Then Exit Do
    Loop

    On Error Resume Next
    CurrentDb.Execute "UPDATE [" & TBL_SYNC_HEARTBEAT & "] SET Stage='shutdown', LastMessage='stop signal', UpdatedAt=" & SQLJetzt() & " WHERE WorkerId='" & SQLSafe(strWorkerId) & "'"
    On Error GoTo ErrHandler

    ReleaseWorkerLease strWorkerId

    If Trim$(strWorkerId) <> "" Then
        WorkerStopSignalLoeschen strWorkerId
    End If

    If blnSelfTerminate Then
        WorkerSelfTerminate
    End If
    WorkerCleanupRuntimeState strWorkerId, "WorkerPollLoop Exit", 0, True, True
    m_strAktiverWorkerId = ""
    Exit Sub

ErrHandler:
    HandleError MODUL_NAME, "WorkerPollLoop"
    ReleaseWorkerLease strWorkerId
    If Trim$(strWorkerId) <> "" Then
        WorkerStopSignalLoeschen strWorkerId
    End If
    WorkerCleanupRuntimeState strWorkerId, "WorkerPollLoop Error", 0, True, True
    If blnSelfTerminate Then
        WorkerSelfTerminate
    End If
    m_strAktiverWorkerId = ""
End Sub


' ---------------------------------------------------------------------------
' Job einreihen (UI-Helfer)
' ---------------------------------------------------------------------------
Public Function EnqueueSyncJob(ByVal strFolderPath As String, _
                               Optional ByVal lngMaxMails As Long = 500, _
                               Optional ByVal blnSubfolders As Boolean = False, _
                               Optional ByVal strCreatedBy As String = "", _
                               Optional ByVal intPriority As Integer = 100) As Long
    On Error GoTo ErrHandler

    Dim strBereinigterPfad As String
    Dim lngTry As Long
    Dim lngMaxTry As Long
    Dim blnInserted As Boolean

    If strCreatedBy = "" Then strCreatedBy = Environ("USERNAME")
    strBereinigterPfad = BereinigeOutlookPfad(strFolderPath)
    lngMaxTry = 4

    Dim db As DAO.Database
    Set db = CurrentDb

    For lngTry = 1 To lngMaxTry
        On Error GoTo InsertErr
        db.Execute "INSERT INTO [" & TBL_SYNC_JOB & "] (CreatedAt, CreatedBy, RequestedFolderPath, RequestedMaxMails, RequestedSubfolders, Status, Priority) VALUES (" & _
                   SQLJetzt() & ", '" & SQLSafe(strCreatedBy) & "', '" & SQLSafe(strBereinigterPfad) & "', " & lngMaxMails & ", " & BoolSQL(blnSubfolders) & ", '" & JOB_STATUS_QUEUED & "', " & intPriority & ")", dbFailOnError
        blnInserted = True
        Exit For

InsertErr:
    If IsLockError(Err.Number) And lngTry < lngMaxTry Then
            Debug.Print "[QUEUE] Insert Retry | Versuch=" & lngTry & " | Err=" & Err.Number & " - " & Err.Description
            Err.Clear
            Sleep 150 * lngTry
            DoEvents
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Next lngTry

    If Not blnInserted Then
        EnqueueSyncJob = 0
        Exit Function
    End If

    EnqueueSyncJob = ErmittleNeuesteJobID(strCreatedBy, strBereinigterPfad)
    If EnqueueSyncJob > 0 Then
        EnsureControlRow EnqueueSyncJob
    Else
        Debug.Print "[QUEUE] WARN: Job angelegt, aber JobID nicht aufloesbar | CreatedBy=" & strCreatedBy & " | Pfad=" & strBereinigterPfad
    End If

    LogInfo "Job eingereiht: ID=" & EnqueueSyncJob & " Pfad=" & strFolderPath, MODUL_NAME
    Exit Function

ErrHandler:
    Debug.Print "[QUEUE] FEHLER | CreatedBy=" & strCreatedBy & " | Pfad=" & strBereinigterPfad & " | Err=" & Err.Number & " - " & Err.Description
    HandleError MODUL_NAME, "EnqueueSyncJob", "CreatedBy=" & strCreatedBy & "|Pfad=" & strBereinigterPfad
    EnqueueSyncJob = 0
End Function

Private Function ErmittleNeuesteJobID(ByVal strCreatedBy As String, ByVal strFolderPath As String) As Long
    On Error GoTo ErrHandler

    Dim lngTry As Long

    For lngTry = 1 To 6
        On Error GoTo LookupErr
        ErmittleNeuesteJobID = CLng(Nz(DMax("JobID", TBL_SYNC_JOB, _
            "CreatedBy='" & SQLSafe(strCreatedBy) & "' AND RequestedFolderPath='" & SQLSafe(strFolderPath) & "'"), 0))
        If ErmittleNeuesteJobID > 0 Then Exit Function

        Sleep 120 * lngTry
        DoEvents
    Next lngTry

    ErmittleNeuesteJobID = 0
    Exit Function

LookupErr:
    If IsLockError(Err.Number) And lngTry < 6 Then
        Err.Clear
        Sleep 120 * lngTry
        DoEvents
        Resume Next
    End If
    Err.Raise Err.Number, Err.Source, Err.Description

ErrHandler:
    ErmittleNeuesteJobID = 0
End Function


' ---------------------------------------------------------------------------
' Atomarer Job-Claim: queued -> running
' ---------------------------------------------------------------------------
Public Function ClaimNextJob(ByVal strWorkerId As String) As Long
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim lngJobID As Long

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT TOP 1 JobID FROM [" & TBL_SYNC_JOB & "] " & _
        "WHERE Status='" & JOB_STATUS_QUEUED & "' " & _
        "ORDER BY Priority ASC, CreatedAt ASC", dbOpenSnapshot)

    If rs.EOF Then
        WorkerDebugPrint "Keine queued Jobs gefunden", strWorkerId
        ClaimNextJob = 0
        GoTo CleanExit
    End If

    lngJobID = CLng(Nz(rs!JobID, 0))

    db.Execute "UPDATE [" & TBL_SYNC_JOB & "] SET " & _
               "Status='" & JOB_STATUS_RUNNING & "', " & _
               "WorkerId='" & SQLSafe(strWorkerId) & "', " & _
               "StartedAt=" & SQLJetzt() & " " & _
               "WHERE JobID=" & lngJobID & " AND Status='" & JOB_STATUS_QUEUED & "'", dbFailOnError

    If db.RecordsAffected = 1 Then
        EnsureControlRow lngJobID
        HeartbeatUpdate strWorkerId, lngJobID, "init", 0, 0, "claimed"
        ClaimNextJob = lngJobID
        LogInfo "Job geclaimt: ID=" & lngJobID & " Worker=" & strWorkerId, MODUL_NAME
        WorkerDebugPrint "Claim erfolgreich", strWorkerId, lngJobID, WorkerHoleJobSnapshot(lngJobID)
    Else
        WorkerDebugPrint "Claim verloren - Job inzwischen belegt", strWorkerId, lngJobID
        ClaimNextJob = 0
    End If

CleanExit:
    WorkerCloseRecordset rs, "ClaimNextJob"
    WorkerCloseDatabase db, "ClaimNextJob"
    Exit Function

ErrHandler:
    HandleError MODUL_NAME, "ClaimNextJob"
    ClaimNextJob = 0
    Resume CleanExit
End Function


' ---------------------------------------------------------------------------
' Heartbeat + Lease aktualisieren
' ---------------------------------------------------------------------------
Public Sub HeartbeatUpdate(ByVal strWorkerId As String, _
                           ByVal lngJobID As Long, _
                           ByVal strStage As String, _
                           ByVal lngCurrent As Long, _
                           ByVal lngTotal As Long, _
                           Optional ByVal strLastMessage As String = "")
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim lngRetry As Long
    Dim lngRec As Long
    lngRetry = COMRetryZaehler()
    lngRec = COMReconnectZaehler()

    If DCount("*", TBL_SYNC_HEARTBEAT, "WorkerId='" & SQLSafe(strWorkerId) & "'") = 0 Then
        db.Execute "INSERT INTO [" & TBL_SYNC_HEARTBEAT & "] (WorkerId, JobID, Stage, CurrentItem, TotalItems, COMRetries, COMReconnects, UpdatedAt, LastMessage) VALUES (" & _
                   "'" & SQLSafe(strWorkerId) & "', " & lngJobID & ", '" & SQLSafe(strStage) & "', " & lngCurrent & ", " & lngTotal & ", " & lngRetry & ", " & lngRec & ", " & SQLJetzt() & ", '" & SQLSafe(strLastMessage) & "')", dbFailOnError
    Else
        db.Execute "UPDATE [" & TBL_SYNC_HEARTBEAT & "] SET " & _
                   "JobID=" & lngJobID & ", " & _
                   "Stage='" & SQLSafe(strStage) & "', " & _
                   "CurrentItem=" & lngCurrent & ", " & _
                   "TotalItems=" & lngTotal & ", " & _
                   "COMRetries=" & lngRetry & ", " & _
                   "COMReconnects=" & lngRec & ", " & _
                   "UpdatedAt=" & SQLJetzt() & ", " & _
                   "LastMessage='" & SQLSafe(strLastMessage) & "' " & _
                   "WHERE WorkerId='" & SQLSafe(strWorkerId) & "'", dbFailOnError
    End If

    RefreshWorkerLease strWorkerId
    WorkerLogTrace MODUL_NAME, "HeartbeatUpdate", "Heartbeat geschrieben", strWorkerId, lngJobID, "Stage=" & strStage & " | Current=" & lngCurrent & "/" & lngTotal & " | Msg=" & Left$(strLastMessage, 120)
    Exit Sub

ErrHandler:
    HandleError MODUL_NAME, "HeartbeatUpdate"
End Sub


' ---------------------------------------------------------------------------
' Pause/Cancel auswerten
' Rueckgabe: "continue" | "paused" | "canceled"
' ---------------------------------------------------------------------------
Public Function HandlePauseCancel(ByVal lngJobID As Long, _
                                  ByVal strWorkerId As String) As String
    On Error GoTo ErrHandler

    HandlePauseCancel = "continue"

    EnsureControlRow lngJobID

    If LeseFlag(lngJobID, "CancelRequested") Then
        UpdateJobStatus lngJobID, JOB_STATUS_CANCEL_REQUESTED
        HandlePauseCancel = "canceled"
        Exit Function
    End If

    If LeseFlag(lngJobID, "PauseRequested") Then
        UpdateJobStatus lngJobID, JOB_STATUS_PAUSED
        HandlePauseCancel = "paused"

        Do
            HeartbeatUpdate strWorkerId, lngJobID, "paused", 0, 0, "pause"

            If WorkerStopSignalAktiv(strWorkerId) Then
                HandlePauseCancel = "canceled"
                Exit Function
            End If

            If WorkerSleepMitStopSignal(strWorkerId, 1000, 200) Then
                HandlePauseCancel = "canceled"
                Exit Function
            End If

            If LeseFlag(lngJobID, "CancelRequested") Then
                UpdateJobStatus lngJobID, JOB_STATUS_CANCEL_REQUESTED
                HandlePauseCancel = "canceled"
                Exit Function
            End If

            If Not LeseFlag(lngJobID, "PauseRequested") Then
                UpdateJobStatus lngJobID, JOB_STATUS_RUNNING
                HandlePauseCancel = "continue"
                Exit Function
            End If
        Loop
    End If

    Exit Function

ErrHandler:
    HandleError MODUL_NAME, "HandlePauseCancel"
    HandlePauseCancel = "canceled"
End Function


' ---------------------------------------------------------------------------
' Job finalisieren
' ---------------------------------------------------------------------------
Public Sub FinalizeJob(ByVal lngJobID As Long, _
                       ByVal strFinalStatus As String, _
                       Optional ByVal strLastError As String = "", _
                       Optional ByVal strWorkerId As String = "")
    On Error GoTo ErrHandler

    Dim strSetErr As String
    strSetErr = "LastError='" & SQLSafe(strLastError) & "'"

    WorkerDebugPrint "Finalize gestartet", strWorkerId, lngJobID, "status=" & strFinalStatus & IIf(Trim$(strLastError) <> "", " | detail=" & Left$(strLastError, 200), "")

    CurrentDb.Execute "UPDATE [" & TBL_SYNC_JOB & "] SET " & _
                      "Status='" & SQLSafe(strFinalStatus) & "', " & _
                      "FinishedAt=" & SQLJetzt() & ", " & strSetErr & " " & _
                      "WHERE JobID=" & lngJobID

    WorkerDebugPrint "Finalize geschrieben", strWorkerId, lngJobID, WorkerHoleJobSnapshot(lngJobID)

    If strWorkerId <> "" Then
        HeartbeatUpdate strWorkerId, lngJobID, "finalize", 0, 0, strFinalStatus
    End If

    Exit Sub

ErrHandler:
    HandleError MODUL_NAME, "FinalizeJob"
End Sub


' ---------------------------------------------------------------------------
' Formulare fuer Queue-Betrieb erstellen
' ---------------------------------------------------------------------------
Public Function EnsureDualAccessForms() As Boolean
    On Error GoTo ErrHandler

    EnsureDualAccessForms = True

    EnsureDualAccessForms = EnsureDualAccessForms And EnsureFormViaFactory(TBL_SYNC_JOB, FORM_SYNC_JOB, True)
    EnsureDualAccessForms = EnsureDualAccessForms And EnsureFormViaFactory(TBL_SYNC_HEARTBEAT, FORM_SYNC_HEARTBEAT, True)
    EnsureDualAccessForms = EnsureDualAccessForms And EnsureFormViaFactory(TBL_WORKER_LEASE, FORM_WORKER_LEASE, True)
    Exit Function

ErrHandler:
    HandleError MODUL_NAME, "EnsureDualAccessForms"
    EnsureDualAccessForms = False
End Function


' ---------------------------------------------------------------------------
' PRIVATE: Job ausfuehren
' ---------------------------------------------------------------------------
Private Sub ExecuteQueuedJob(ByVal lngJobID As Long, ByVal strWorkerId As String)
    On Error GoTo ErrHandler

    Dim rs As DAO.Recordset
    Dim strPfad As String
    Dim lngMax As Long
    Dim blnSub As Boolean
    Dim objFolder As Object
    Dim strFlow As String
    Dim strSyncStatus As String
    Dim lngSyncFehler As Long
    Dim lngSyncLaufID As Long

    WorkerDebugPrint "ExecuteQueuedJob gestartet", strWorkerId, lngJobID

    Set rs = CurrentDb.OpenRecordset("SELECT RequestedFolderPath, RequestedMaxMails, RequestedSubfolders FROM [" & TBL_SYNC_JOB & "] WHERE JobID=" & lngJobID, dbOpenSnapshot)
    If rs.EOF Then
        WorkerDebugPrint "Jobdatensatz nicht gefunden", strWorkerId, lngJobID
        FinalizeJob lngJobID, JOB_STATUS_FAILED, "Job nicht gefunden", strWorkerId
        GoTo CleanExit
    End If

    strPfad = Nz(rs!RequestedFolderPath, "")
    lngMax = CLng(Nz(rs!RequestedMaxMails, 500))
    blnSub = CBool(Nz(rs!RequestedSubfolders, False))

    WorkerDebugPrint "Jobparameter geladen", strWorkerId, lngJobID, "Pfad=" & strPfad & " | MaxMails=" & lngMax & " | Subfolders=" & IIf(blnSub, "True", "False")

    HeartbeatUpdate strWorkerId, lngJobID, "resolve", 0, 0, "resolve folder"
    WorkerDebugPrint "Ordner wird aufgeloest", strWorkerId, lngJobID, strPfad
    Set objFolder = OeffneOrdner(strPfad)
    If objFolder Is Nothing Then
        WorkerDebugPrint "Ordneraufloesung fehlgeschlagen", strWorkerId, lngJobID, strPfad
        FinalizeJob lngJobID, JOB_STATUS_FAILED, "Ordner nicht gefunden: " & strPfad, strWorkerId
        GoTo CleanExit
    End If

    On Error Resume Next
    WorkerDebugPrint "Ordner aufgeloest", strWorkerId, lngJobID, "Name=" & Nz(objFolder.Name, "") & " | Items=" & Nz(objFolder.Items.Count, 0) & " | FolderPath=" & Nz(objFolder.FolderPath, "")
    On Error GoTo ErrHandler

    strFlow = HandlePauseCancel(lngJobID, strWorkerId)
    If strFlow = "canceled" Then
        WorkerDebugPrint "Job vor Sync abgebrochen", strWorkerId, lngJobID
        FinalizeJob lngJobID, JOB_STATUS_CANCELED, "Abgebrochen vor Sync", strWorkerId
        GoTo CleanExit
    End If

    HeartbeatUpdate strWorkerId, lngJobID, "extract", 0, 0, "start sync"
    WorkerDebugPrint "SyncFolder startet", strWorkerId, lngJobID, "Projekt=DualAccess | Phase=Job" & Format$(lngJobID, "000000")
    SyncFolder objFolder, "DualAccess", "Job" & Format$(lngJobID, "000000"), lngMax, blnSub
    WorkerDebugPrint "SyncFolder beendet", strWorkerId, lngJobID, WorkerHoleJobSnapshot(lngJobID)

    strFlow = HandlePauseCancel(lngJobID, strWorkerId)
    If strFlow = "canceled" Then
        WorkerDebugPrint "Job nach Sync als canceled markiert", strWorkerId, lngJobID
        FinalizeJob lngJobID, JOB_STATUS_CANCELED, "Abgebrochen nach Sync", strWorkerId
    ElseIf Not HoleSyncErgebnisFuerJob(lngJobID, strSyncStatus, lngSyncFehler, lngSyncLaufID, strWorkerId) Then
        WorkerDebugPrint "Kein Sync-Ergebnis auffindbar", strWorkerId, lngJobID, WorkerHoleJobSnapshot(lngJobID)
        FinalizeJob lngJobID, JOB_STATUS_FAILED, "Kein passender SyncLauf gefunden", strWorkerId
    ElseIf StrComp(strSyncStatus, "Abgeschlossen", vbTextCompare) <> 0 Or lngSyncFehler > 0 Then
        WorkerDebugPrint "Sync-Ergebnis ungleich completed", strWorkerId, lngJobID, "SyncLaufID=" & lngSyncLaufID & " | SyncStatus=" & strSyncStatus & " | Fehler=" & lngSyncFehler
        FinalizeJob lngJobID, JOB_STATUS_FAILED, "SyncStatus=" & strSyncStatus & ", Fehler=" & lngSyncFehler, strWorkerId
    Else
        WorkerDebugPrint "Sync-Ergebnis erfolgreich", strWorkerId, lngJobID, "SyncLaufID=" & lngSyncLaufID
        FinalizeJob lngJobID, JOB_STATUS_COMPLETED, "", strWorkerId
    End If

CleanExit:
    WorkerCloseRecordset rs, "ExecuteQueuedJob"
    WorkerReleaseObject objFolder, "OutlookFolder", "ExecuteQueuedJob"
    WorkerCleanupRuntimeState strWorkerId, "ExecuteQueuedJob", lngJobID, True, True
    Exit Sub

ErrHandler:
    WorkerDebugPrint "ExecuteQueuedJob FEHLER", strWorkerId, lngJobID, "Err=" & Err.Number & " - " & Err.Description
    FinalizeJob lngJobID, JOB_STATUS_FAILED, Err.Number & " - " & Err.Description, strWorkerId
    Resume CleanExit
End Sub


Private Sub EnsureControlRow(ByVal lngJobID As Long)
    If DCount("*", TBL_SYNC_CONTROL, "JobID=" & lngJobID) = 0 Then
        CurrentDb.Execute "INSERT INTO [" & TBL_SYNC_CONTROL & "] (JobID, PauseRequested, CancelRequested, UpdatedAt) VALUES (" & _
                          lngJobID & ", False, False, " & SQLJetzt() & ")"
        WorkerLogTrace MODUL_NAME, "EnsureControlRow", "ControlRow erstellt", m_strAktiverWorkerId, lngJobID
    End If
End Sub

Private Function LeseFlag(ByVal lngJobID As Long, ByVal strField As String) As Boolean
    On Error Resume Next
    LeseFlag = CBool(Nz(DLookup(strField, TBL_SYNC_CONTROL, "JobID=" & lngJobID), False))
    On Error GoTo 0
End Function

Private Sub UpdateJobStatus(ByVal lngJobID As Long, ByVal strStatus As String)
    CurrentDb.Execute "UPDATE [" & TBL_SYNC_JOB & "] SET Status='" & SQLSafe(strStatus) & "' WHERE JobID=" & lngJobID
    WorkerLogTrace MODUL_NAME, "UpdateJobStatus", "Jobstatus aktualisiert", m_strAktiverWorkerId, lngJobID, "Status=" & strStatus
End Sub

Private Sub RefreshWorkerLease(ByVal strWorkerId As String)
    On Error GoTo ErrHandler

    Dim lngStaleS As Long
    Dim lngTry As Long
    Dim lngMaxTry As Long
    Dim blnDone As Boolean
    lngStaleS = CLng(CacheGetConfig(CFG_WORKER_STALE_S, "90"))
    If lngStaleS < 30 Then lngStaleS = 30
    lngMaxTry = 4

    For lngTry = 1 To lngMaxTry
        On Error GoTo WriteErr

        If DCount("*", TBL_WORKER_LEASE, "WorkerId='" & SQLSafe(strWorkerId) & "'") = 0 Then
            CurrentDb.Execute "INSERT INTO [" & TBL_WORKER_LEASE & "] (WorkerId, LeaseUntil, UpdatedAt, HostName, SessionUser) VALUES (" & _
                              "'" & SQLSafe(strWorkerId) & "', DateAdd('s', " & lngStaleS & ", " & SQLJetzt() & "), " & SQLJetzt() & ", '" & SQLSafe(Environ("COMPUTERNAME")) & "', '" & SQLSafe(Environ("USERNAME")) & "')", dbFailOnError
        Else
            CurrentDb.Execute "UPDATE [" & TBL_WORKER_LEASE & "] SET " & _
                              "LeaseUntil=DateAdd('s', " & lngStaleS & ", " & SQLJetzt() & "), " & _
                              "UpdatedAt=" & SQLJetzt() & ", " & _
                              "HostName='" & SQLSafe(Environ("COMPUTERNAME")) & "', " & _
                              "SessionUser='" & SQLSafe(Environ("USERNAME")) & "' " & _
                              "WHERE WorkerId='" & SQLSafe(strWorkerId) & "'", dbFailOnError
        End If

        blnDone = True
        WorkerLogTrace MODUL_NAME, "RefreshWorkerLease", "Lease aktualisiert", strWorkerId, 0, "Versuch=" & lngTry & " | StaleSek=" & lngStaleS
        Exit For

WriteErr:
        ' Typische Lock-/Concurrency-Probleme kurz aussitzen und erneut versuchen.
        If (Err.Number = 3188 Or Err.Number = 3197 Or Err.Number = 3218 Or Err.Number = 3260) And lngTry < lngMaxTry Then
            Err.Clear
            Sleep 200 * lngTry
            DoEvents
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Next lngTry

    If Not blnDone Then
        LogWarn "WorkerLease Refresh ohne Erfolg nach Retries: " & strWorkerId, MODUL_NAME
    End If

    If DCount("*", TBL_WORKER_LEASE, "WorkerId='" & SQLSafe(strWorkerId) & "'") = 0 Then
        LogWarn "WorkerLease nach Refresh nicht vorhanden: " & strWorkerId, MODUL_NAME
        Debug.Print "[WORKER] Lease fehlt nach Refresh | WorkerId=" & strWorkerId
    End If
    Exit Sub

ErrHandler:
    LogError "WorkerLease Refresh fehlgeschlagen: Worker=" & strWorkerId & " | Err " & Err.Number & " - " & Err.Description, MODUL_NAME
    Debug.Print "[WORKER] Lease Refresh FEHLER | WorkerId=" & strWorkerId & " | Err=" & Err.Number & " - " & Err.Description
End Sub

Private Sub ReleaseWorkerLease(ByVal strWorkerId As String)
    On Error GoTo ErrHandler
    If strWorkerId <> "" Then
        CurrentDb.Execute "DELETE FROM [" & TBL_WORKER_LEASE & "] WHERE WorkerId='" & SQLSafe(strWorkerId) & "'"
        WorkerLogTrace MODUL_NAME, "ReleaseWorkerLease", "Lease geloescht", strWorkerId
    End If
    Exit Sub

ErrHandler:
    LogWarn "WorkerLease Release fehlgeschlagen: Worker=" & strWorkerId & " | Err " & Err.Number & " - " & Err.Description, MODUL_NAME
    Debug.Print "[WORKER] Lease Release FEHLER | WorkerId=" & strWorkerId & " | Err=" & Err.Number & " - " & Err.Description
End Sub

Private Function EnsureFormViaFactory(ByVal strTable As String, ByVal strForm As String, ByVal blnContinuous As Boolean) As Boolean
    On Error GoTo Fallback

    If FormExists(strForm) Then
        If FormRecordSourceMatches(strForm, strTable) Then
            EnsureFormViaFactory = True
            Exit Function
        End If

        LogWarn "Formular wird neu erstellt (RecordSource passt nicht): " & strForm & " -> " & strTable, MODUL_NAME
        DeleteFormSafe strForm
    End If

    ' Versucht die bestehende frm_factory-Basis zu nutzen.
    ' Parameter: sourceTable, targetFormName, layout=1, view=1/2, footer=1
    Application.Run "AutoGenerateForm", strTable, strForm, 1, IIf(blnContinuous, 1, 2), 1, False, "", "", ""
    EnsureFormViaFactory = (FormExists(strForm) And FormRecordSourceMatches(strForm, strTable))

    If EnsureFormViaFactory Then
        Exit Function
    End If

    DeleteFormSafe strForm
    EnsureFormViaFactory = CreateSimpleForm(strTable, strForm)
    Exit Function

Fallback:
    LogWarn "AutoGenerateForm fehlgeschlagen fuer " & strForm & ": " & Err.Number & " - " & Err.Description, MODUL_NAME
    DeleteFormSafe strForm
    EnsureFormViaFactory = CreateSimpleForm(strTable, strForm)
End Function

Private Function CreateSimpleForm(ByVal strTable As String, ByVal strForm As String) As Boolean
    On Error GoTo ErrHandler

    If FormExists(strForm) Then
        CreateSimpleForm = True
        Exit Function
    End If

    Dim frm As Form
    Dim strTempName As String
    Set frm = CreateForm
    strTempName = frm.Name
    frm.RecordSource = strTable
    frm.Caption = strForm
    frm.DefaultView = 2   ' Datasheet
    frm.NavigationButtons = True
    frm.RecordSelectors = True

    DoCmd.Save acForm, strTempName
    DoCmd.Close acForm, strTempName, acSaveYes
    DoCmd.Rename strForm, acForm, strTempName

    CreateSimpleForm = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Len(strTempName) > 0 Then
        DoCmd.Close acForm, strTempName, acSaveNo
        DoCmd.DeleteObject acForm, strTempName
    End If
    On Error GoTo 0
    HandleError MODUL_NAME, "CreateSimpleForm", strForm
    CreateSimpleForm = False
End Function

Private Function FormExists(ByVal strForm As String) As Boolean
    On Error GoTo ErrHandler

    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If StrComp(obj.Name, strForm, vbTextCompare) = 0 Then
            FormExists = True
            Exit Function
        End If
    Next obj

    FormExists = False
    Exit Function

ErrHandler:
    FormExists = False
End Function

Private Sub DeleteFormSafe(ByVal strForm As String)
    On Error Resume Next

    If Len(Trim$(strForm)) = 0 Then Exit Sub

    If CurrentProject.AllForms(strForm).IsLoaded Then
        DoCmd.Close acForm, strForm, acSaveNo
    End If

    DoCmd.DeleteObject acForm, strForm
    Err.Clear
    On Error GoTo 0
End Sub

Private Function FormRecordSourceMatches(ByVal strForm As String, ByVal strTable As String) As Boolean
    On Error GoTo ErrHandler

    If Not FormExists(strForm) Then
        FormRecordSourceMatches = False
        Exit Function
    End If

    DoCmd.OpenForm strForm, acDesign

    Dim frm As Form
    Set frm = Forms(strForm)

    FormRecordSourceMatches = (StrComp(NormalisiereObjektName(Nz(frm.RecordSource, "")), _
                                       NormalisiereObjektName(strTable), _
                                       vbTextCompare) = 0)

    DoCmd.Close acForm, strForm, acSaveNo
    Set frm = Nothing
    Exit Function

ErrHandler:
    On Error Resume Next
    If CurrentProject.AllForms(strForm).IsLoaded Then
        DoCmd.Close acForm, strForm, acSaveNo
    End If
    On Error GoTo 0
    FormRecordSourceMatches = False
End Function

Private Function NormalisiereObjektName(ByVal strName As String) As String
    Dim s As String
    s = Trim$(strName)
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    NormalisiereObjektName = s
End Function

Private Function MakroExistiert(ByVal strMacro As String) As Boolean
    On Error GoTo ErrHandler

    Dim obj As AccessObject
    For Each obj In CurrentProject.AllMacros
        If StrComp(obj.Name, strMacro, vbTextCompare) = 0 Then
            MakroExistiert = True
            Exit Function
        End If
    Next obj

    MakroExistiert = False
    Exit Function

ErrHandler:
    MakroExistiert = False
End Function

Private Function HoleSyncErgebnisFuerJob(ByVal lngJobID As Long, _
                                         ByRef strStatus As String, _
                                         ByRef lngAnzahlFehler As Long, _
                                         Optional ByRef lngSyncLaufID As Long = 0, _
                                         Optional ByVal strWorkerId As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim strPhase As String
    Dim strCriteria As String
    Dim varSyncLaufID As Variant
    Dim lngTry As Long

    strPhase = "Job" & Format$(lngJobID, "000000")
    strCriteria = "Projekt='DualAccess' AND Phase='" & SQLSafe(strPhase) & "'"

    WorkerDebugPrint "Suche Sync-Ergebnis", strWorkerId, lngJobID, strCriteria

    For lngTry = 1 To WORKER_SYNC_LOOKUP_RETRIES
        On Error GoTo LookupErr

        varSyncLaufID = DMax("SyncLaufID", TBL_SYNC_LAUF, strCriteria)
        If Not IsNull(varSyncLaufID) And CLng(Nz(varSyncLaufID, 0)) > 0 Then
            lngSyncLaufID = CLng(varSyncLaufID)
            strStatus = Nz(DLookup("Status", TBL_SYNC_LAUF, "SyncLaufID=" & lngSyncLaufID), "")
            lngAnzahlFehler = CLng(Nz(DLookup("AnzahlFehler", TBL_SYNC_LAUF, "SyncLaufID=" & lngSyncLaufID), 0))
            WorkerDebugPrint "Sync-Ergebnis gefunden", strWorkerId, lngJobID, "Versuch=" & lngTry & " | SyncLaufID=" & lngSyncLaufID & " | Status=" & strStatus & " | Fehler=" & lngAnzahlFehler
            HoleSyncErgebnisFuerJob = True
            Exit Function
        End If

        WorkerDebugPrint "Sync-Ergebnis noch nicht sichtbar", strWorkerId, lngJobID, "Versuch=" & lngTry & " | Phase=" & strPhase
        Sleep WORKER_SYNC_LOOKUP_SLEEP_MS * lngTry
        DoEvents
    Next lngTry

    WorkerDebugPrint "Sync-Ergebnis nicht gefunden nach Retries", strWorkerId, lngJobID, "Phase=" & strPhase
    HoleSyncErgebnisFuerJob = False
    Exit Function

LookupErr:
    If (Err.Number = 3704 Or Err.Number = 3188 Or Err.Number = 3197 Or Err.Number = 3218 Or Err.Number = 3260) And lngTry < WORKER_SYNC_LOOKUP_RETRIES Then
        WorkerDebugPrint "Sync-Ergebnis Lookup Retry", strWorkerId, lngJobID, "Versuch=" & lngTry & " | Err=" & Err.Number & " - " & Err.Description
        Err.Clear
        Sleep WORKER_SYNC_LOOKUP_SLEEP_MS * lngTry
        DoEvents
        Resume Next
    End If
    Err.Raise Err.Number, Err.Source, Err.Description

ErrHandler:
    strStatus = ""
    lngAnzahlFehler = 0
    lngSyncLaufID = 0
    WorkerDebugPrint "Sync-Ergebnis FEHLER", strWorkerId, lngJobID, "Err=" & Err.Number & " - " & Err.Description
    HoleSyncErgebnisFuerJob = False
End Function

Private Sub WorkerDebugPrint(ByVal strMessage As String, _
                             Optional ByVal strWorkerId As String = "", _
                             Optional ByVal lngJobID As Long = 0, _
                             Optional ByVal strDetails As String = "")
    WorkerLogTrace MODUL_NAME, "", strMessage, strWorkerId, lngJobID, strDetails
End Sub

Public Sub WorkerLogTrace(ByVal strModul As String, _
                          ByVal strProzedur As String, _
                          ByVal strMessage As String, _
                          Optional ByVal strWorkerId As String = "", _
                          Optional ByVal lngJobID As Long = 0, _
                          Optional ByVal strDetails As String = "")
    WorkerLogEvent LOG_TRACE, strModul, strProzedur, strMessage, strWorkerId, lngJobID, strDetails
End Sub

Public Sub WorkerLogDebug(ByVal strModul As String, _
                          ByVal strProzedur As String, _
                          ByVal strMessage As String, _
                          Optional ByVal strWorkerId As String = "", _
                          Optional ByVal lngJobID As Long = 0, _
                          Optional ByVal strDetails As String = "")
    WorkerLogEvent LOG_DEBUG, strModul, strProzedur, strMessage, strWorkerId, lngJobID, strDetails
End Sub

Public Sub WorkerLogInfo(ByVal strModul As String, _
                         ByVal strProzedur As String, _
                         ByVal strMessage As String, _
                         Optional ByVal strWorkerId As String = "", _
                         Optional ByVal lngJobID As Long = 0, _
                         Optional ByVal strDetails As String = "")
    WorkerLogEvent LOG_INFO, strModul, strProzedur, strMessage, strWorkerId, lngJobID, strDetails
End Sub

Public Sub WorkerLogWarn(ByVal strModul As String, _
                         ByVal strProzedur As String, _
                         ByVal strMessage As String, _
                         Optional ByVal strWorkerId As String = "", _
                         Optional ByVal lngJobID As Long = 0, _
                         Optional ByVal strDetails As String = "")
    WorkerLogEvent LOG_WARN, strModul, strProzedur, strMessage, strWorkerId, lngJobID, strDetails
End Sub

Public Sub WorkerLogError(ByVal strModul As String, _
                          ByVal strProzedur As String, _
                          ByVal strMessage As String, _
                          Optional ByVal strWorkerId As String = "", _
                          Optional ByVal lngJobID As Long = 0, _
                          Optional ByVal strDetails As String = "")
    WorkerLogEvent LOG_ERROR, strModul, strProzedur, strMessage, strWorkerId, lngJobID, strDetails
End Sub

Public Sub WorkerLogEvent(ByVal lngLevel As Integer, _
                          ByVal strModul As String, _
                          ByVal strProzedur As String, _
                          ByVal strMessage As String, _
                          Optional ByVal strWorkerId As String = "", _
                          Optional ByVal lngJobID As Long = 0, _
                          Optional ByVal strDetails As String = "")
    Dim strLine As String

    If lngLevel > WorkerAktiverLogLevel() Then Exit Sub

    strLine = "[WORKER-" & WorkerTraceLevelName(lngLevel) & "] " & Format$(Now, "hh:nn:ss") & " | " & strMessage
    If Trim$(strWorkerId) <> "" Then strLine = strLine & " | WorkerId=" & strWorkerId
    If lngJobID > 0 Then strLine = strLine & " | JobID=" & lngJobID
    If Trim$(strDetails) <> "" Then strLine = strLine & " | " & strDetails

    Debug.Print strLine
    WorkerTracePersist lngLevel, strModul, strProzedur, strMessage, strWorkerId, lngJobID, strDetails
End Sub

Public Sub WorkerTraceReport(Optional ByVal strWorkerId As String = "", _
                             Optional ByVal lngJobID As Long = 0, _
                             Optional ByVal lngTop As Long = 50, _
                             Optional ByVal lngMinLevel As Integer = LOG_TRACE)
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSql As String
    Dim strWhere As String

    If lngTop < 1 Then lngTop = 1
    If lngTop > 500 Then lngTop = 500

    Debug.Print String(70, "-")
    Debug.Print "[WORKER-TRACE-REPORT] WorkerId=" & IIf(Trim$(strWorkerId) = "", "(alle)", strWorkerId) & _
                " | JobID=" & IIf(lngJobID > 0, CStr(lngJobID), "(alle)") & _
                " | Top=" & lngTop

    If Not WorkerTraceTabelleSicher() Then
        Debug.Print "[WORKER-TRACE-REPORT] nicht verfuegbar | Tabelle/Link fehlt"
        Debug.Print String(70, "-")
        Exit Sub
    End If

    Set db = GetSafeDB()

    strWhere = " WHERE LogLevel<=" & lngMinLevel
    If Trim$(strWorkerId) <> "" Then strWhere = strWhere & " AND WorkerId='" & SQLSafe(strWorkerId) & "'"
    If lngJobID > 0 Then strWhere = strWhere & " AND JobID=" & lngJobID

    strSql = "SELECT TOP " & lngTop & " LoggedAt, LevelName, WorkerId, JobID, Modul, Prozedur, Nachricht, Details " & _
             "FROM [" & TBL_WORKER_TRACE & "]" & strWhere & " ORDER BY TraceID DESC"

    Set rs = db.OpenRecordset(strSql, dbOpenSnapshot)
    If rs.EOF Then
        Debug.Print "  (keine Worker-Trace-Eintraege gefunden)"
    Else
        Do While Not rs.EOF
            Debug.Print "  " & Nz(rs!LoggedAt, "") & " | " & Nz(rs!LevelName, "") & _
                        " | WorkerId=" & Nz(rs!WorkerId, "") & _
                        " | JobID=" & Nz(rs!JobID, 0) & _
                        " | " & Nz(rs!Modul, "") & IIf(Nz(rs!Prozedur, "") <> "", "." & Nz(rs!Prozedur, ""), "") & _
                        " | " & Nz(rs!Nachricht, "")
            If Nz(rs!Details, "") <> "" Then Debug.Print "    -> " & Nz(rs!Details, "")
            rs.MoveNext
        Loop
    End If

CleanExit:
    WorkerCloseRecordset rs, "WorkerTraceReport"
    WorkerCloseDatabase db, "WorkerTraceReport"
    Debug.Print String(70, "-")
    Exit Sub

ErrHandler:
    Debug.Print "[WORKER-TRACE-REPORT] FEHLER | Err=" & Err.Number & " - " & Err.Description
    Resume CleanExit
End Sub

Private Sub WorkerTracePersist(ByVal lngLevel As Integer, _
                               ByVal strModul As String, _
                               ByVal strProzedur As String, _
                               ByVal strMessage As String, _
                               ByVal strWorkerId As String, _
                               ByVal lngJobID As Long, _
                               ByVal strDetails As String)
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim strSql As String
    Dim lngTry As Long

    If Not WorkerTraceTabelleSicher() Then Exit Sub

    Set db = GetSafeDB()
    strSql = "INSERT INTO [" & TBL_WORKER_TRACE & "] (LoggedAt, LogLevel, LevelName, WorkerId, JobID, Modul, Prozedur, Nachricht, Details, HostName, SessionUser) VALUES (" & _
             SQLJetzt() & ", " & lngLevel & ", '" & SQLSafe(WorkerTraceLevelName(lngLevel)) & "', '" & SQLSafe(strWorkerId) & "', " & lngJobID & ", '" & _
             SQLSafe(Left$(strModul, 100)) & "', '" & SQLSafe(Left$(strProzedur, 100)) & "', '" & SQLSafe(Left$(strMessage, 255)) & "', '" & _
             SQLSafe(Left$(strDetails, 2000)) & "', '" & SQLSafe(Environ$("COMPUTERNAME")) & "', '" & SQLSafe(Environ$("USERNAME")) & "')"

    For lngTry = 1 To 4
        On Error GoTo WriteErr
        db.Execute strSql, dbFailOnError
        GoTo CleanExit

WriteErr:
        If IsLockError(Err.Number) And lngTry < 4 Then
            Err.Clear
            Sleep 100 * lngTry
            DoEvents
        Else
            Exit For
        End If
    Next lngTry

CleanExit:
    WorkerCloseDatabase db, "WorkerTracePersist"
    Exit Sub

ErrHandler:
    Resume CleanExit
End Sub

Private Function WorkerTraceTabelleSicher() As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim strBackendPfad As String

    strBackendPfad = Trim$(CacheGetConfig(CFG_BACKEND_PFAD, ""))

    If strBackendPfad = "" And m_blnWorkerTraceChecked Then
        WorkerTraceTabelleSicher = m_blnWorkerTraceAvailable
        Exit Function
    End If

    m_blnWorkerTraceChecked = True
    If strBackendPfad = "" Then
        m_blnWorkerTraceAvailable = TabelleExistiert(TBL_WORKER_TRACE)
    Else
        m_blnWorkerTraceAvailable = WorkerTraceTabelleImBackend(strBackendPfad)
    End If

    If m_blnWorkerTraceAvailable And (strBackendPfad = "" Or TabelleExistiert(TBL_WORKER_TRACE)) Then
        WorkerTraceTabelleSicher = True
        Exit Function
    End If

    If strBackendPfad <> "" Then
        Set db = DBEngine.OpenDatabase(strBackendPfad)
        If ErstelleBackendTabellenInDB(db) Then
            db.Close
            Set db = Nothing
            m_blnWorkerTraceAvailable = EnsureWorkerTraceLink(strBackendPfad)
            WorkerTraceTabelleSicher = m_blnWorkerTraceAvailable
            Exit Function
        Else
            db.Close
            Set db = Nothing
        End If

        m_blnWorkerTraceAvailable = WorkerTraceTabelleImBackend(strBackendPfad) And TabelleExistiert(TBL_WORKER_TRACE)
        WorkerTraceTabelleSicher = m_blnWorkerTraceAvailable
        Exit Function
    End If

    Set db = CurrentDb
    db.Execute "CREATE TABLE [" & TBL_WORKER_TRACE & "] (" & _
               "TraceID AUTOINCREMENT CONSTRAINT PK_WorkerTrace PRIMARY KEY, " & _
               "LoggedAt DATETIME, " & _
               "LogLevel INTEGER, " & _
               "LevelName TEXT(10), " & _
               "WorkerId TEXT(100), " & _
               "JobID LONG, " & _
               "Modul TEXT(100), " & _
               "Prozedur TEXT(100), " & _
               "Nachricht TEXT(255), " & _
               "Details LONGTEXT, " & _
               "HostName TEXT(100), " & _
               "SessionUser TEXT(100))"
    db.Execute "CREATE INDEX idx_WorkerTrace_LoggedAt ON [" & TBL_WORKER_TRACE & "] (LoggedAt)"
    db.Execute "CREATE INDEX idx_WorkerTrace_Worker ON [" & TBL_WORKER_TRACE & "] (WorkerId, LoggedAt)"
    db.Execute "CREATE INDEX idx_WorkerTrace_Job ON [" & TBL_WORKER_TRACE & "] (JobID, LoggedAt)"
    db.Execute "CREATE INDEX idx_WorkerTrace_Level ON [" & TBL_WORKER_TRACE & "] (LogLevel, LoggedAt)"

    m_blnWorkerTraceAvailable = True
    WorkerTraceTabelleSicher = True
    Set db = Nothing
    Exit Function

ErrHandler:
    If strBackendPfad <> "" Then
        m_blnWorkerTraceAvailable = WorkerTraceTabelleImBackend(strBackendPfad) And TabelleExistiert(TBL_WORKER_TRACE)
    Else
        m_blnWorkerTraceAvailable = TabelleExistiert(TBL_WORKER_TRACE)
    End If
    WorkerTraceTabelleSicher = m_blnWorkerTraceAvailable
End Function

Private Function EnsureWorkerTraceLink(ByVal strBackendPfad As String) As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim strConnect As String

    If Trim$(strBackendPfad) = "" Then Exit Function
    If Not WorkerTraceTabelleImBackend(strBackendPfad) Then Exit Function

    strConnect = ";DATABASE=" & strBackendPfad
    Set db = CurrentDb

    If TabelleExistiert(TBL_WORKER_TRACE) Then
        If IstLinkedTable(TBL_WORKER_TRACE) Then
            If StrComp(Nz(GetConnectString(TBL_WORKER_TRACE), ""), strConnect, vbTextCompare) = 0 Then
                On Error Resume Next
                db.TableDefs(TBL_WORKER_TRACE).RefreshLink
                On Error GoTo ErrHandler
                EnsureWorkerTraceLink = (Err.Number = 0)
                If EnsureWorkerTraceLink Then Exit Function
                Err.Clear
            End If

            On Error Resume Next
            db.TableDefs.Delete TBL_WORKER_TRACE
            db.TableDefs.Refresh
            Err.Clear
            On Error GoTo ErrHandler
        Else
            On Error Resume Next
            db.Execute "INSERT INTO [" & TBL_WORKER_TRACE & "] IN '" & Replace(strBackendPfad, "'", "''") & "' SELECT * FROM [" & TBL_WORKER_TRACE & "]"
            Err.Clear
            db.Execute "DROP TABLE [" & TBL_WORKER_TRACE & "]"
            db.TableDefs.Refresh
            Err.Clear
            On Error GoTo ErrHandler
        End If
    End If

    Set td = db.CreateTableDef(TBL_WORKER_TRACE)
    td.Connect = strConnect
    td.SourceTableName = TBL_WORKER_TRACE
    db.TableDefs.Append td
    db.TableDefs.Refresh

    EnsureWorkerTraceLink = TabelleExistiert(TBL_WORKER_TRACE) And IstLinkedTable(TBL_WORKER_TRACE)
    Set td = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    On Error Resume Next
    Set td = Nothing
    Set db = Nothing
    EnsureWorkerTraceLink = False
End Function

Private Function WorkerTraceTabelleImBackend(ByVal strBackendPfad As String) As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef

    If Trim$(strBackendPfad) = "" Then Exit Function

    Set db = DBEngine.OpenDatabase(strBackendPfad)
    For Each td In db.TableDefs
        If td.Name = TBL_WORKER_TRACE Then
            WorkerTraceTabelleImBackend = True
            Exit For
        End If
    Next td

    db.Close
    Set db = Nothing
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    WorkerTraceTabelleImBackend = False
End Function

Private Function WorkerTraceLevelName(ByVal lngLevel As Integer) As String
    Select Case lngLevel
        Case LOG_ERROR: WorkerTraceLevelName = "ERROR"
        Case LOG_WARN: WorkerTraceLevelName = "WARN"
        Case LOG_INFO: WorkerTraceLevelName = "INFO"
        Case LOG_DEBUG: WorkerTraceLevelName = "DEBUG"
        Case LOG_TRACE: WorkerTraceLevelName = "TRACE"
        Case Else: WorkerTraceLevelName = "LOG"
    End Select
End Function

Private Function WorkerHoleJobSnapshot(ByVal lngJobID As Long) As String
    On Error GoTo ErrHandler

    Dim strStatus As String
    Dim strWorkerId As String
    Dim varStartedAt As Variant
    Dim varFinishedAt As Variant
    Dim strLastError As String

    If lngJobID <= 0 Then Exit Function

    strStatus = Nz(DLookup("Status", TBL_SYNC_JOB, "JobID=" & lngJobID), "")
    strWorkerId = Nz(DLookup("WorkerId", TBL_SYNC_JOB, "JobID=" & lngJobID), "")
    varStartedAt = DLookup("StartedAt", TBL_SYNC_JOB, "JobID=" & lngJobID)
    varFinishedAt = DLookup("FinishedAt", TBL_SYNC_JOB, "JobID=" & lngJobID)
    strLastError = Nz(DLookup("LastError", TBL_SYNC_JOB, "JobID=" & lngJobID), "")

    WorkerHoleJobSnapshot = "Status=" & strStatus & _
                            " | WorkerId=" & strWorkerId & _
                            " | StartedAt=" & Nz(varStartedAt, "") & _
                            " | FinishedAt=" & Nz(varFinishedAt, "")
    If Trim$(strLastError) <> "" Then
        WorkerHoleJobSnapshot = WorkerHoleJobSnapshot & " | LastError=" & Left$(strLastError, 200)
    End If
    Exit Function

ErrHandler:
    WorkerHoleJobSnapshot = "Snapshot-Fehler: " & Err.Number & " - " & Err.Description
End Function

Private Function WorkerIdStandard() As String
    WorkerIdStandard = Environ("COMPUTERNAME") & "_" & Environ("USERNAME")
End Function

Public Function WorkerIstOnline(ByVal strWorkerId As String, _
                                Optional ByVal lngMaxAlterSek As Long = 30) As Boolean
    On Error GoTo ErrHandler

    Dim strCrit As String
    strCrit = "WorkerId='" & SQLSafe(strWorkerId) & "'"

    Dim varLeaseUntil As Variant
    varLeaseUntil = DLookup("LeaseUntil", TBL_WORKER_LEASE, strCrit)
    If Not IsNull(varLeaseUntil) Then
        If CDate(varLeaseUntil) >= DateAdd("s", -2, Now) Then
            WorkerIstOnline = True
            Exit Function
        End If
    End If

    Dim varUpdated As Variant
    varUpdated = DLookup("UpdatedAt", TBL_SYNC_HEARTBEAT, strCrit)
    If IsNull(varUpdated) Then
        WorkerIstOnline = False
    Else
        WorkerIstOnline = (DateDiff("s", CDate(varUpdated), Now) <= lngMaxAlterSek)
    End If
    Exit Function

ErrHandler:
    WorkerIstOnline = False
End Function

Private Function DualAccessBereitOhneMigration() As Boolean
    On Error GoTo ErrHandler

    DualAccessBereitOhneMigration = TabelleExistiert(TBL_SYNC_JOB) _
        And TabelleExistiert(TBL_SYNC_HEARTBEAT) _
        And TabelleExistiert(TBL_SYNC_CONTROL) _
        And TabelleExistiert(TBL_WORKER_LEASE) _
        And WorkerTraceBereitOhneSetup() _
        And Len(Nz(CacheGetConfig(CFG_WORKER_POLL_MS, ""), "")) > 0 _
        And Len(Nz(CacheGetConfig(CFG_WORKER_HB_S, ""), "")) > 0 _
        And Len(Nz(CacheGetConfig(CFG_WORKER_STALE_S, ""), "")) > 0
    Exit Function

ErrHandler:
    DualAccessBereitOhneMigration = False
End Function

Private Function WorkerTraceBereitOhneSetup() As Boolean
    On Error GoTo ErrHandler

    Dim strBackendPfad As String

    strBackendPfad = Trim$(CacheGetConfig(CFG_BACKEND_PFAD, ""))
    If strBackendPfad = "" Then
        WorkerTraceBereitOhneSetup = TabelleExistiert(TBL_WORKER_TRACE)
    Else
        WorkerTraceBereitOhneSetup = TabelleExistiert(TBL_WORKER_TRACE) And TabelleExistiertInBackend(TBL_WORKER_TRACE, strBackendPfad)
    End If
    Exit Function

ErrHandler:
    WorkerTraceBereitOhneSetup = False
End Function

Private Function WorkerAktiverLogLevel() As Integer
    On Error GoTo ErrHandler

    Dim lngLevel As Long

    lngLevel = CLng(Val(CacheGetConfig(CFG_LOG_LEVEL, CStr(IIf(g_intLogLevel > 0, g_intLogLevel, LOG_INFO)))))
    If lngLevel < LOG_NONE Then lngLevel = LOG_NONE
    If lngLevel > LOG_TRACE Then lngLevel = LOG_TRACE

    WorkerAktiverLogLevel = CInt(lngLevel)
    Exit Function

ErrHandler:
    WorkerAktiverLogLevel = IIf(g_intLogLevel > 0, g_intLogLevel, LOG_INFO)
End Function

Public Sub WorkerStatusReport(Optional ByVal strWorkerId As String = "")
    On Error GoTo ErrHandler

    Dim strCrit As String
    If Trim$(strWorkerId) <> "" Then
        strCrit = " WHERE WorkerId='" & SQLSafe(strWorkerId) & "'"
    Else
        strCrit = ""
    End If

    Debug.Print String(70, "-")
    Debug.Print "[WORKER-STATUS] " & IIf(strWorkerId = "", "alle Worker", strWorkerId)

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT hb.WorkerId, hb.JobID, hb.Stage, hb.UpdatedAt, wl.LeaseUntil " & _
        "FROM [" & TBL_SYNC_HEARTBEAT & "] hb " & _
        "LEFT JOIN [" & TBL_WORKER_LEASE & "] wl ON hb.WorkerId = wl.WorkerId" & strCrit & _
        " ORDER BY hb.UpdatedAt DESC", dbOpenSnapshot)

    If rs.EOF Then
        Debug.Print "  (keine Heartbeats/Leases gefunden)"
    Else
        Do While Not rs.EOF
            Debug.Print "  WorkerId   : " & Nz(rs!WorkerId, "")
            Debug.Print "  JobID      : " & Nz(rs!JobID, 0)
            Debug.Print "  Stage      : " & Nz(rs!Stage, "")
            Debug.Print "  UpdatedAt  : " & Nz(rs!UpdatedAt, "")
            Debug.Print "  LeaseUntil : " & Nz(rs!LeaseUntil, "")
            Debug.Print "  Online     : " & IIf(WorkerIstOnline(Nz(rs!WorkerId, "")), "JA", "Nein")
            Debug.Print String(40, ".")
            rs.MoveNext
        Loop
    End If

CleanExit:
    WorkerCloseRecordset rs, "WorkerStatusReport"
    Debug.Print String(70, "-")
    Exit Sub

ErrHandler:
    HandleError MODUL_NAME, "WorkerStatusReport", strWorkerId
    Resume CleanExit
End Sub

Private Function WaitForWorkerOnline(ByVal strWorkerId As String, ByVal lngWaitSek As Long) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    dblStart = Timer

    Do
        If WorkerIstOnline(strWorkerId) Then
            WaitForWorkerOnline = True
            Exit Function
        End If

        Sleep 250
        DoEvents
    Loop While SekundenDiff(dblStart, Timer) < lngWaitSek

    WaitForWorkerOnline = False
    Exit Function

ErrHandler:
    WaitForWorkerOnline = False
End Function

Private Function WorkerQuelleStartbereit(ByVal strWorkerDbPath As String, _
                                         Optional ByRef strGrund As String = "") As Boolean
    On Error GoTo ErrHandler

    strGrund = ""
    If Trim$(strWorkerDbPath) = "" Then
        strGrund = "Worker-Quelle ist leer"
        WorkerQuelleStartbereit = False
        Exit Function
    End If

    If Dir(strWorkerDbPath) = "" Then
        strGrund = "Worker-Quelle nicht gefunden: " & strWorkerDbPath
        WorkerQuelleStartbereit = False
        Exit Function
    End If

    WorkerQuelleStartbereit = True
    Exit Function

ErrHandler:
    strGrund = "Worker-Quelle nicht lesbar: Err " & Err.Number & " - " & Err.Description
    WorkerQuelleStartbereit = False
End Function

Private Function WorkerTempOrdnerBereit(Optional ByRef strGrund As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim strOrdner As String
    strOrdner = WorkerLaunchOrdnerPfad()
    ErstelleOrdner strOrdner

    If Dir(strOrdner, vbDirectory) = "" Then
        strGrund = "Worker-Temp-Ordner konnte nicht erstellt werden: " & strOrdner
        WorkerTempOrdnerBereit = False
        Exit Function
    End If

    WorkerTempOrdnerBereit = True
    Exit Function

ErrHandler:
    strGrund = "Worker-Temp-Ordner nicht verfuegbar: Err " & Err.Number & " - " & Err.Description
    WorkerTempOrdnerBereit = False
End Function

Private Function WorkerLaunchDbVorbereiten(ByVal strWorkerDbPath As String, _
                                           ByVal strWorkerId As String, _
                                           ByRef strLaunchDbPath As String, _
                                           Optional ByRef strGrund As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim strMethode As String

    strGrund = ""
    strLaunchDbPath = ""

    If Not WorkerQuelleStartbereit(strWorkerDbPath, strGrund) Then
        WorkerLaunchDbVorbereiten = False
        Exit Function
    End If

    If Not WorkerTempOrdnerBereit(strGrund) Then
        WorkerLaunchDbVorbereiten = False
        Exit Function
    End If

    strLaunchDbPath = WorkerLaunchDateiPfad(strWorkerId)
    WorkerBereinigeLaunchArtefakte strWorkerId, False

    If Not WorkerStelleDateiFrei(strWorkerId, strLaunchDbPath, True, strGrund) Then
        WorkerLaunchDbVorbereiten = False
        Exit Function
    End If

    If Dir(strLaunchDbPath) <> "" Then
        If Not WorkerLoescheDateiSicher(strWorkerId, strLaunchDbPath, strGrund) Then
            WorkerLaunchDbVorbereiten = False
            Exit Function
        End If
        Call WorkerLoescheDateiSicher(strWorkerId, WorkerLockDateiPfad(strLaunchDbPath), strGrund)
    End If

    If Not WorkerDateiKopieren(strWorkerDbPath, strLaunchDbPath, strMethode, strGrund) Then
        WorkerLaunchDbVorbereiten = False
        Exit Function
    End If

    LogInfo "Worker-FE bereitgestellt via " & strMethode & ": " & strLaunchDbPath, MODUL_NAME
    Debug.Print "[WORKER] FE bereitgestellt | Methode=" & strMethode & " | DB=" & strLaunchDbPath

    WorkerLaunchDbVorbereiten = True
    Exit Function

ErrHandler:
    strGrund = "Worker-FE-Bereitstellung fehlgeschlagen: Err " & Err.Number & " - " & Err.Description
    WorkerLaunchDbVorbereiten = False
End Function

Private Function WorkerLaunchOrdnerPfad() As String
    WorkerLaunchOrdnerPfad = WorkerDokumentePfad() & "OutlookSync\WorkerFE\"
End Function

Private Function WorkerLaunchDateiPfad(ByVal strWorkerId As String) As String
    WorkerLaunchDateiPfad = WorkerLaunchOrdnerPfad() & "WorkerFE_" & BereinigeDateiname(strWorkerId, 48) & ".accdb"
End Function

Private Function WorkerDokumentePfad() As String
    On Error GoTo Fallback

    Dim objShell As Object
    Dim strPfad As String

    Set objShell = CreateObject("WScript.Shell")
    strPfad = Nz(objShell.SpecialFolders("MyDocuments"), "")
    Set objShell = Nothing

    If Trim$(strPfad) = "" Then GoTo Fallback
    WorkerDokumentePfad = NormalisierePfad(strPfad)
    Exit Function

Fallback:
    On Error Resume Next
    Set objShell = Nothing
    WorkerDokumentePfad = NormalisierePfad(Environ("USERPROFILE") & "\Documents")
    On Error GoTo 0
End Function

Private Function WorkerDateiKopieren(ByVal strQuelle As String, _
                                     ByVal strZiel As String, _
                                     ByRef strMethode As String, _
                                     Optional ByRef strGrund As String = "") As Boolean
    On Error GoTo ErrHandler

    strMethode = ""
    strGrund = ""

    On Error Resume Next
    FileCopy strQuelle, strZiel
    If Err.Number = 0 And Dir(strZiel) <> "" Then
        On Error GoTo 0
        strMethode = "VBA FileCopy"
        WorkerDateiKopieren = True
        Exit Function
    End If

    strGrund = "VBA FileCopy: Err " & Err.Number & " - " & Err.Description
    Err.Clear
    On Error GoTo ErrHandler

    If WorkerDateiKopierenViaCmdCopy(strQuelle, strZiel) Then
        strMethode = "WScript.Shell cmd copy"
        WorkerDateiKopieren = True
        Exit Function
    End If

    If WorkerDateiKopierenViaRobocopy(strQuelle, strZiel) Then
        strMethode = "robocopy"
        WorkerDateiKopieren = True
        Exit Function
    End If

    strGrund = strGrund & " | Externe Kopie ebenfalls fehlgeschlagen"
    WorkerDateiKopieren = False
    Exit Function

ErrHandler:
    strGrund = "Worker-Dateikopie fehlgeschlagen: Err " & Err.Number & " - " & Err.Description
    WorkerDateiKopieren = False
End Function

Private Function WorkerDateiKopierenViaCmdCopy(ByVal strQuelle As String, ByVal strZiel As String) As Boolean
    On Error GoTo ErrHandler

    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd /c copy /y " & Quote(strQuelle) & " " & Quote(strZiel), 0, True
    Set objShell = Nothing

    WorkerDateiKopierenViaCmdCopy = (Dir(strZiel) <> "")
    Exit Function

ErrHandler:
    On Error Resume Next
    Set objShell = Nothing
    On Error GoTo 0
    WorkerDateiKopierenViaCmdCopy = False
End Function

Private Function WorkerDateiKopierenViaRobocopy(ByVal strQuelle As String, ByVal strZiel As String) As Boolean
    On Error GoTo ErrHandler

    Dim strRobo As String
    Dim strSrcDir As String
    Dim strDstDir As String
    Dim strDatei As String
    Dim objShell As Object

    strRobo = Environ("SystemRoot") & "\System32\robocopy.exe"
    If Dir(strRobo) = "" Then
        WorkerDateiKopierenViaRobocopy = False
        Exit Function
    End If

    strSrcDir = Left$(strQuelle, InStrRev(strQuelle, "\") - 1)
    strDstDir = Left$(strZiel, InStrRev(strZiel, "\") - 1)
    strDatei = Mid$(strQuelle, InStrRev(strQuelle, "\") + 1)

    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd /c robocopy " & Quote(strSrcDir) & " " & Quote(strDstDir) & " " & Quote(strDatei) & " /R:1 /W:1 /NJH /NJS /NFL /NDL", 0, True
    Set objShell = Nothing

    WorkerDateiKopierenViaRobocopy = (Dir(strZiel) <> "")
    Exit Function

ErrHandler:
    On Error Resume Next
    Set objShell = Nothing
    On Error GoTo 0
    WorkerDateiKopierenViaRobocopy = False
End Function

Private Function HoleAccessExePfad() As String
    On Error Resume Next

    Dim strDir As String
    strDir = SysCmd(acSysCmdAccessDir)
    If strDir = "" Then
        HoleAccessExePfad = ""
        Exit Function
    End If

    If Right$(strDir, 1) <> "\" Then strDir = strDir & "\"
    HoleAccessExePfad = strDir & "MSACCESS.EXE"

    If Dir(HoleAccessExePfad) = "" Then HoleAccessExePfad = ""
    On Error GoTo 0
End Function

Private Function BoolSQL(ByVal blnValue As Boolean) As String
    If blnValue Then
        BoolSQL = "True"
    Else
        BoolSQL = "False"
    End If
End Function

Private Function Quote(ByVal s As String) As String
    Quote = Chr$(34) & s & Chr$(34)
End Function

Private Sub WorkerSelfTerminate()
    On Error Resume Next
    DoCmd.Quit acQuitSaveNone
    Application.Quit acQuitSaveNone
    On Error GoTo 0
End Sub

Private Sub WorkerApplyBackgroundUiMode(Optional ByVal blnHide As Boolean = True)
    On Error Resume Next

    DoCmd.Echo False
    DoCmd.RunCommand acCmdAppMinimize

    If blnHide Then
        ShowWindow Application.hWndAccessApp, SW_HIDE
    Else
        ShowWindow Application.hWndAccessApp, SW_MINIMIZE
    End If

    On Error GoTo 0
End Sub

Private Function WorkerStopSignalPfad(ByVal strWorkerId As String) As String
    WorkerStopSignalPfad = WorkerLaunchOrdnerPfad() & "stop_" & BereinigeDateiname(strWorkerId, 48) & ".flag"
End Function

Private Function WorkerLockDateiPfad(ByVal strDbPfad As String) As String
    If Len(strDbPfad) >= 6 And LCase$(Right$(strDbPfad, 6)) = ".accdb" Then
        WorkerLockDateiPfad = Left$(strDbPfad, Len(strDbPfad) - 6) & ".laccdb"
    Else
        WorkerLockDateiPfad = strDbPfad & ".laccdb"
    End If
End Function

Private Sub WorkerBereinigeLaunchArtefakte(ByVal strWorkerId As String, Optional ByVal blnAuchStabileDatei As Boolean = False)
    On Error GoTo CleanExit

    Dim strOrdner As String
    Dim strPrefix As String
    Dim strDatei As String
    Dim strPfad As String
    Dim strStabileDatei As String

    strOrdner = WorkerLaunchOrdnerPfad()
    strPrefix = "WorkerFE_" & BereinigeDateiname(strWorkerId, 48)
    strStabileDatei = WorkerLaunchDateiPfad(strWorkerId)

    strDatei = Dir(strOrdner & strPrefix & "_*.accdb")
    Do While strDatei <> ""
        strPfad = strOrdner & strDatei
        Call WorkerLoescheDateiSicher(strWorkerId, strPfad)
        Call WorkerLoescheDateiSicher(strWorkerId, WorkerLockDateiPfad(strPfad))
        strDatei = Dir()
    Loop

    If blnAuchStabileDatei Then
        If Dir(strStabileDatei) <> "" Then
            Call WorkerLoescheDateiSicher(strWorkerId, strStabileDatei)
            Call WorkerLoescheDateiSicher(strWorkerId, WorkerLockDateiPfad(strStabileDatei))
        End If
    End If

CleanExit:
    On Error Resume Next
End Sub

Private Function WorkerDateiIstInBenutzung(ByVal strPfad As String) As Boolean
    On Error GoTo InUse

    Dim intFile As Integer

    If Trim$(strPfad) = "" Or Dir(strPfad) = "" Then
        WorkerDateiIstInBenutzung = False
        Exit Function
    End If

    intFile = FreeFile
    Open strPfad For Binary Access Read Write Lock Read Write As #intFile
    Close #intFile
    WorkerDateiIstInBenutzung = False
    Exit Function

InUse:
    On Error Resume Next
    If intFile > 0 Then Close #intFile
    WorkerDateiIstInBenutzung = True
End Function

Private Function WorkerStelleDateiFrei(ByVal strWorkerId As String, _
                                       ByVal strPfad As String, _
                                       Optional ByVal blnProzessBereinigung As Boolean = True, _
                                       Optional ByRef strGrund As String = "") As Boolean
    On Error GoTo ErrHandler

    strGrund = ""
    If Trim$(strPfad) = "" Then
        WorkerStelleDateiFrei = True
        Exit Function
    End If

    If Dir(strPfad) = "" Then
        WorkerStelleDateiFrei = True
        Exit Function
    End If

    If Not WorkerDateiIstInBenutzung(strPfad) Then
        Debug.Print "[WORKER] Datei frei | WorkerId=" & strWorkerId & " | Datei=" & strPfad
        WorkerStelleDateiFrei = True
        Exit Function
    End If

    Debug.Print "[WORKER] Datei gesperrt | WorkerId=" & strWorkerId & " | Datei=" & strPfad

    If blnProzessBereinigung Then
        Call WorkerBeendeAccessProzesseFuerDatei(strWorkerId, strPfad, 0, True)
    End If

    If WorkerWarteAufDateiFreigabe(strPfad, 5000) Then
        Debug.Print "[WORKER] Datei-Freigabe bestaetigt | WorkerId=" & strWorkerId & " | Datei=" & strPfad
        WorkerStelleDateiFrei = True
    Else
        strGrund = "Datei bleibt gesperrt: " & strPfad
        Debug.Print "[WORKER] Datei-Freigabe FEHLER | WorkerId=" & strWorkerId & " | Datei=" & strPfad
        WorkerStelleDateiFrei = False
    End If
    Exit Function

ErrHandler:
    strGrund = "Datei-Freigabe fehlgeschlagen: Err " & Err.Number & " - " & Err.Description
    WorkerStelleDateiFrei = False
End Function

Private Function WorkerLoescheDateiSicher(ByVal strWorkerId As String, _
                                          ByVal strPfad As String, _
                                          Optional ByRef strGrund As String = "") As Boolean
    On Error GoTo ErrHandler

    strGrund = ""
    If Trim$(strPfad) = "" Then
        WorkerLoescheDateiSicher = True
        Exit Function
    End If

    If Dir(strPfad) = "" Then
        WorkerLoescheDateiSicher = True
        Exit Function
    End If

    If Not WorkerStelleDateiFrei(strWorkerId, strPfad, True, strGrund) Then
        WorkerLoescheDateiSicher = False
        Exit Function
    End If

    Kill strPfad
    WorkerLoescheDateiSicher = (Dir(strPfad) = "")
    Debug.Print "[WORKER] Datei geloescht | WorkerId=" & strWorkerId & " | Datei=" & strPfad & " | OK=" & IIf(WorkerLoescheDateiSicher, "1", "0")
    Exit Function

ErrHandler:
    strGrund = "Datei-Loeschen fehlgeschlagen: Err " & Err.Number & " - " & Err.Description & " | Datei=" & strPfad
    Debug.Print "[WORKER] Datei-Loeschen FEHLER | WorkerId=" & strWorkerId & " | Datei=" & strPfad & " | Err=" & Err.Number & " - " & Err.Description
    WorkerLoescheDateiSicher = False
End Function

Private Sub WorkerStopSignalSetzen(ByVal strWorkerId As String)
    On Error GoTo ErrHandler

    If Trim$(strWorkerId) = "" Then Exit Sub

    ErstelleOrdner WorkerLaunchOrdnerPfad()

    Dim strPfad As String
    Dim intFile As Integer
    strPfad = WorkerStopSignalPfad(strWorkerId)
    intFile = FreeFile
    Open strPfad For Output As #intFile
    Print #intFile, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #intFile
    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #intFile
    On Error GoTo 0
End Sub

Private Sub WorkerStopSignalLoeschen(ByVal strWorkerId As String)
    On Error Resume Next
    If Trim$(strWorkerId) = "" Then Exit Sub
    Kill WorkerStopSignalPfad(strWorkerId)
    On Error GoTo 0
End Sub

Private Function WorkerStopSignalAktiv(ByVal strWorkerId As String) As Boolean
    On Error GoTo ErrHandler
    If Trim$(strWorkerId) = "" Then
        WorkerStopSignalAktiv = False
    Else
        WorkerStopSignalAktiv = (Dir(WorkerStopSignalPfad(strWorkerId)) <> "")
    End If
    Exit Function

ErrHandler:
    WorkerStopSignalAktiv = False
End Function

Public Function WorkerAktiverStopAngefordert() As Boolean
    On Error GoTo ErrHandler

    If Trim$(m_strAktiverWorkerId) = "" Then
        WorkerAktiverStopAngefordert = False
    Else
        WorkerAktiverStopAngefordert = WorkerStopSignalAktiv(m_strAktiverWorkerId)
    End If
    Exit Function

ErrHandler:
    WorkerAktiverStopAngefordert = False
End Function

Private Function WorkerSleepMitStopSignal(ByVal strWorkerId As String, _
                                          ByVal lngSleepMs As Long, _
                                          Optional ByVal lngSliceMs As Long = 200) As Boolean
    On Error GoTo ErrHandler

    Dim lngRest As Long
    Dim lngNow As Long

    If lngSleepMs <= 0 Then
        WorkerSleepMitStopSignal = WorkerStopSignalAktiv(strWorkerId)
        Exit Function
    End If

    If lngSliceMs < 50 Then lngSliceMs = 50
    If lngSliceMs > 1000 Then lngSliceMs = 1000

    lngRest = lngSleepMs
    Do While lngRest > 0
        If WorkerStopSignalAktiv(strWorkerId) Then
            WorkerSleepMitStopSignal = True
            Exit Function
        End If

        lngNow = lngSliceMs
        If lngNow > lngRest Then lngNow = lngRest

        Sleep lngNow
        DoEvents
        lngRest = lngRest - lngNow
    Loop

    WorkerSleepMitStopSignal = WorkerStopSignalAktiv(strWorkerId)
    Exit Function

ErrHandler:
    WorkerSleepMitStopSignal = False
End Function

Private Sub WorkerStopTaskKill(ByVal lngTaskID As Long, ByVal blnForce As Boolean)
    On Error GoTo ErrHandler

    If lngTaskID <= 0 Then Exit Sub

    Dim objShell As Object
    Dim strCmd As String

    Set objShell = CreateObject("WScript.Shell")
    strCmd = "cmd /c taskkill /PID " & CStr(lngTaskID)
    If blnForce Then strCmd = strCmd & " /F"
    strCmd = strCmd & " >nul 2>&1"
    Debug.Print "[WORKER] taskkill | PID=" & lngTaskID & " | Force=" & IIf(blnForce, "1", "0")
    objShell.Run strCmd, 0, True

CleanExit:
    On Error Resume Next
    Set objShell = Nothing
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Resume CleanExit
End Sub

Private Function WorkerWarteAufExit(ByVal lngTaskID As Long, ByVal lngTimeoutMs As Long) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    dblStart = Timer

    If lngTimeoutMs < 0 Then lngTimeoutMs = 0

    Do
        If Not WorkerProzessLaeuft(lngTaskID) Then
            WorkerWarteAufExit = True
            Exit Function
        End If

        Sleep 200
        DoEvents
    Loop While (SekundenDiff(dblStart, Timer) * 1000#) < lngTimeoutMs

    WorkerWarteAufExit = (Not WorkerProzessLaeuft(lngTaskID))
    Exit Function

ErrHandler:
    WorkerWarteAufExit = False
End Function

Private Function WorkerWarteAufDateiFreigabe(ByVal strPfad As String, ByVal lngTimeoutMs As Long) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    dblStart = Timer

    If lngTimeoutMs < 0 Then lngTimeoutMs = 0

    Do
        If Not WorkerDateiIstInBenutzung(strPfad) Then
            WorkerWarteAufDateiFreigabe = True
            Exit Function
        End If

        Sleep 200
        DoEvents
    Loop While (SekundenDiff(dblStart, Timer) * 1000#) < lngTimeoutMs

    WorkerWarteAufDateiFreigabe = (Not WorkerDateiIstInBenutzung(strPfad))
    Exit Function

ErrHandler:
    WorkerWarteAufDateiFreigabe = False
End Function

Private Function WorkerBeendeAccessProzesseFuerDatei(ByVal strWorkerId As String, _
                                                     ByVal strDbPfad As String, _
                                                     Optional ByVal lngPreferTaskID As Long = 0, _
                                                     Optional ByVal blnForce As Boolean = True) As Long
    On Error GoTo ErrHandler

    Dim objWmi As Object
    Dim colProc As Object
    Dim objProc As Object
    Dim strCmdLine As String
    Dim lngPid As Long

    If Trim$(strDbPfad) = "" Then Exit Function

    Set objWmi = GetObject("winmgmts:\\.\root\cimv2")
    Set colProc = objWmi.ExecQuery("SELECT ProcessId, Name, CommandLine FROM Win32_Process WHERE Name='MSACCESS.EXE'")

    For Each objProc In colProc
        lngPid = CLng(Nz(objProc.ProcessId, 0))
        strCmdLine = Nz(objProc.CommandLine, "")

        If lngPid > 0 Then
            If (lngPreferTaskID > 0 And lngPid = lngPreferTaskID) Or WorkerCommandLineVerweistAufDatei(strCmdLine, strDbPfad) Then
                Debug.Print "[WORKER] Access-Prozess-Treffer | WorkerId=" & strWorkerId & " | PID=" & lngPid & " | Cmd=" & Left$(strCmdLine, 220)
                WorkerStopTaskKill lngPid, blnForce
                If WorkerWarteAufExit(lngPid, 4000) Then
                    WorkerBeendeAccessProzesseFuerDatei = WorkerBeendeAccessProzesseFuerDatei + 1
                    Debug.Print "[WORKER] Access-Prozess beendet | WorkerId=" & strWorkerId & " | PID=" & lngPid
                Else
                    Debug.Print "[WORKER] Access-Prozess haengt weiter | WorkerId=" & strWorkerId & " | PID=" & lngPid
                End If
            End If
        End If
    Next objProc

    If WorkerBeendeAccessProzesseFuerDatei = 0 Then
        Debug.Print "[WORKER] Keine passenden Access-Prozesse gefunden | WorkerId=" & strWorkerId & " | Datei=" & strDbPfad
    End If

CleanExit:
    On Error Resume Next
    Set objProc = Nothing
    Set colProc = Nothing
    Set objWmi = Nothing
    On Error GoTo 0
    Exit Function

ErrHandler:
    Debug.Print "[WORKER] Prozesssuche FEHLER | WorkerId=" & strWorkerId & " | Datei=" & strDbPfad & " | Err=" & Err.Number & " - " & Err.Description
    Resume CleanExit
End Function

Private Function WorkerCommandLineVerweistAufDatei(ByVal strCommandLine As String, ByVal strDbPfad As String) As Boolean
    If Trim$(strCommandLine) = "" Or Trim$(strDbPfad) = "" Then Exit Function
    WorkerCommandLineVerweistAufDatei = (InStr(1, strCommandLine, strDbPfad, vbTextCompare) > 0)
End Function

Private Function WorkerProzessLaeuft(ByVal lngTaskID As Long) As Boolean
    On Error GoTo ErrHandler

    If lngTaskID <= 0 Then
        WorkerProzessLaeuft = False
        Exit Function
    End If

    Dim objWmi As Object
    Dim colProc As Object
    Dim objProc As Object

    Set objWmi = GetObject("winmgmts:\\.\root\cimv2")
    Set colProc = objWmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE ProcessId=" & lngTaskID)

    WorkerProzessLaeuft = False
    For Each objProc In colProc
        WorkerProzessLaeuft = True
        Exit For
    Next objProc

CleanExit:
    On Error Resume Next
    Set objProc = Nothing
    Set colProc = Nothing
    Set objWmi = Nothing
    On Error GoTo 0
    Exit Function

ErrHandler:
    WorkerProzessLaeuft = False
    Resume CleanExit
End Function

Private Sub WorkerCloseRecordset(ByRef rs As DAO.Recordset, Optional ByVal strKontext As String = "")
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
        If Trim$(strKontext) <> "" Then Debug.Print "[WORKER] Recordset freigegeben | Kontext=" & strKontext
    End If
    On Error GoTo 0
End Sub

Private Sub WorkerCloseDatabase(ByRef db As DAO.Database, Optional ByVal strKontext As String = "")
    On Error Resume Next
    If Not db Is Nothing Then
        Set db = Nothing
        If Trim$(strKontext) <> "" Then Debug.Print "[WORKER] Database-Referenz freigegeben | Kontext=" & strKontext
    End If
    On Error GoTo 0
End Sub

Private Sub WorkerReleaseObject(ByRef objRef As Object, _
                                Optional ByVal strObjektart As String = "Objekt", _
                                Optional ByVal strKontext As String = "")
    On Error Resume Next
    If Not objRef Is Nothing Then
        Set objRef = Nothing
        Debug.Print "[WORKER] " & strObjektart & " freigegeben" & IIf(Trim$(strKontext) <> "", " | Kontext=" & strKontext, "")
    End If
    On Error GoTo 0
End Sub

Private Sub WorkerCleanupRuntimeState(ByVal strWorkerId As String, _
                                      ByVal strKontext As String, _
                                      Optional ByVal lngJobID As Long = 0, _
                                      Optional ByVal blnDisconnectOutlook As Boolean = True, _
                                      Optional ByVal blnForceDaoCleanup As Boolean = True)
    On Error Resume Next

    Debug.Print "[WORKER] Runtime-Cleanup Start | WorkerId=" & strWorkerId & _
                IIf(lngJobID > 0, " | JobID=" & lngJobID, "") & _
                " | Kontext=" & strKontext & _
                " | Outlook=" & IIf(blnDisconnectOutlook, "1", "0") & _
                " | DAO=" & IIf(blnForceDaoCleanup, "1", "0")

    If blnDisconnectOutlook Then DisconnectAll
    If blnForceDaoCleanup Then ForceCleanup

    Debug.Print "[WORKER] Runtime-Cleanup Ende | WorkerId=" & strWorkerId & _
                IIf(lngJobID > 0, " | JobID=" & lngJobID, "") & _
                " | Kontext=" & strKontext
    On Error GoTo 0
End Sub


