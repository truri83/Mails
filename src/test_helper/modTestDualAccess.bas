Option Compare Database
Option Explicit

' ===========================================================================
' modTestDualAccess - Tests fuer Dual-Access Queue/Worker (v0.7)
' ===========================================================================
' Fokus:
'   - Setup/Migration (Tabellen + Worker-Config)
'   - Queue (Enqueue + Control-Row)
'   - Claim (queued -> running)
'   - Heartbeat + Worker-Lease
'   - Cancel-Handling
'   - Finalize
'
' AUFRUF (Direktbereich):
'   RunDualAccessTests
'   RunDualAccessWorkerLaunchSmokeTest
'   RunDualAccessEmailIntegrationTest "Postfach\\Posteingang\\Testordner"
'
' HINWEIS:
'   Die Tests arbeiten mit isolierten Testdaten (CreatedBy/WorkerId Prefix).
'   Es werden keine produktiven Sync-Laeufe gestartet.
'   Der Email-Integrationstest ist EXPLIZIT separat, da er echten Outlook-
'   und Sync-I/O ausfuehrt.
' ===========================================================================

Private Const MODUL_NAME As String = "modTestDualAccess"
Private Const TEST_PREFIX As String = "DA_TEST_"


' ---------------------------------------------------------------------------
' ENTRY POINT
' ---------------------------------------------------------------------------
Public Sub RunDualAccessTests()
    On Error GoTo ErrHandler

    TestRunStart "DUAL-ACCESS TESTS (Queue/Worker)"

    Debug.Print "  HINWEIS: RunDualAccessTests startet KEINEN echten Access-Worker."
    Debug.Print "           Fuer TaskID/Online-Nachweis: RunDualAccessWorkerLaunchSmokeTest"
    Debug.Print ""

    CleanupDualAccessTestData

    Test_DualAccessSetup
    Test_EnqueueCreatesQueueAndControl
    Test_ClaimNextJob
    Test_HeartbeatAndLease
    Test_CancelHandling
    Test_FinalizeJob

    CleanupDualAccessTestData

    TestRunEnd
    Exit Sub

ErrHandler:
    AssertFail "RunDualAccessTests Fehler: " & Err.Number & " - " & Err.Description
    CleanupDualAccessTestData
    TestRunEnd
End Sub


' ---------------------------------------------------------------------------
' OPTIONAL: ECHTER WORKER-LAUNCH-SMOKETEST
' ---------------------------------------------------------------------------
' Startet einen echten zweiten Access-Prozess und wartet auf Lease/Heartbeat.
' Dabei erscheinen TaskID und Online-Bestaetigung direkt im Direktfenster.
'
' Beispiel:
'   RunDualAccessWorkerLaunchSmokeTest
'   RunDualAccessWorkerLaunchSmokeTest 10
'   RunDualAccessWorkerLaunchSmokeTest 10, True
Public Sub RunDualAccessWorkerLaunchSmokeTest(Optional ByVal lngWaitSek As Long = 10, _
                                              Optional ByVal blnHideWorkerWindow As Boolean = True)
    On Error GoTo ErrHandler

    Dim strWorker As String
    Dim lngWorkerTaskID As Long
    Dim dblStep As Double
    Dim blnDiagPrinted As Boolean

    TestRunStart "DUAL-ACCESS WORKER-LAUNCH SMOKE"
    SuiteStart "Worker Start + Online"

    Dim strGrund As String

    If lngWaitSek < 1 Then lngWaitSek = 1
    If lngWaitSek > 60 Then lngWaitSek = 60

    dblStep = Timer
    If Not EnsureDualAccessReadyFast() Then
        AssertFail "SetupDualAccessNoAdmin(False, False) fehlgeschlagen"
        SuiteEnd
        GoTo CleanExit
    End If
    Debug.Print "[TIMING] SetupDualAccessReady: " & FormatSekundenDiff(dblStep)

    dblStep = Timer
    If Not WorkerStartVoraussetzungenOK(strGrund, CurrentDb.Name) Then
        WorkerStartVoraussetzungenReport CurrentDb.Name
        AssertSkip "Worker-Start Voraussetzungen nicht erfuellt: " & strGrund
        SuiteEnd
        GoTo CleanExit
    End If
    Debug.Print "[TIMING] WorkerStartVoraussetzungenOK: " & FormatSekundenDiff(dblStep)

    strWorker = BuildTestWorkerId("SMOKE")

    Debug.Print "[TEST] Worker-Windowmodus angefordert: " & IIf(blnHideWorkerWindow, "hide", "minimized")

    dblStep = Timer
    AssertIsTrue WorkerStarten(CurrentDb.Name, strWorker, lngWaitSek, lngWorkerTaskID, True, blnHideWorkerWindow), "WorkerStarten bestaetigt Online-Start"
    Debug.Print "[TIMING] WorkerStarten (+WaitForOnline): " & FormatSekundenDiff(dblStep)
    If lngWorkerTaskID > 0 And Not WorkerIstOnline(strWorker) Then
        PrintTestWorkerDiagnostics "Smoke: Start nicht rechtzeitig bestaetigt", strWorker, 0, 40
        blnDiagPrinted = True
    End If

    dblStep = Timer
    AssertIsTrue WaitForWorkerOnlineStatus(strWorker, 4), "WorkerIstOnline=True"
    Debug.Print "[TIMING] WorkerIstOnline: " & FormatSekundenDiff(dblStep)

    Debug.Print ""
    Debug.Print "[TEST-HINWEIS] Erwartete Direktfenster-Zeilen vor/nach diesem Test:"
    Debug.Print "[WORKER] Shell gestartet | WorkerId=" & strWorker & " | TaskID=..."
    Debug.Print "[WORKER] ONLINE bestaetigt | WorkerId=" & strWorker
    Debug.Print ""
    WorkerStatusReport strWorker

    dblStep = Timer
    Sleep 1500
    DoEvents
    Debug.Print "[TIMING] Sleep/Propagation: " & FormatSekundenDiff(dblStep)

    dblStep = Timer
    AssertIsTrue WaitForWorkerLease(strWorker, 8), "WorkerLease vorhanden"
    Debug.Print "[TIMING] Lease-Check (wait<=8s): " & FormatSekundenDiff(dblStep)

    dblStep = Timer
    AssertIsTrue WaitForWorkerHeartbeat(strWorker, 8), "WorkerHeartbeat vorhanden"
    Debug.Print "[TIMING] Heartbeat-Check: " & FormatSekundenDiff(dblStep)

    SuiteEnd

CleanExit:
    If Len(strWorker) > 0 And Not blnDiagPrinted Then
        PrintTestWorkerDiagnostics "Smoke: Abschlussdiagnose", strWorker, 0, 25
    End If
    SmokeTestCleanupWorkerProcess strWorker, lngWorkerTaskID, 20
    TestRunEnd
    Exit Sub

ErrHandler:
    AssertFail "Worker-Launch-Smoke Fehler: " & Err.Number & " - " & Err.Description
    PrintTestWorkerDiagnostics "Smoke: ErrHandler", strWorker, 0, 50
    SmokeTestCleanupWorkerProcess strWorker, lngWorkerTaskID, 20
    SuiteEnd
    TestRunEnd
End Sub


' ---------------------------------------------------------------------------
' OPTIONAL: ECHTER EMAIL-/WORKER-INTEGRATIONSTEST
' ---------------------------------------------------------------------------
' Startet einen echten Worker, queued einen echten Outlook-Ordner und wartet
' auf Claim + Abschluss. Nicht Teil von RunDualAccessTests, da echter I/O.
'
' Beispiel:
'   RunDualAccessEmailIntegrationTest "MeinPostfach\\Posteingang\\Test"
'   RunDualAccessEmailIntegrationTest "MeinPostfach\\Posteingang\\Test", 5, False, 120
Public Sub RunDualAccessEmailIntegrationTest(ByVal strOrdnerPfad As String, _
                                             Optional ByVal lngMaxMails As Long = 5, _
                                             Optional ByVal blnSubfolders As Boolean = False, _
                                             Optional ByVal lngTimeoutSek As Long = 120)
    On Error GoTo ErrHandler

    Dim strWorker As String
    Dim lngWorkerTaskID As Long
    Dim dblStep As Double
    Dim lngJobID As Long
    Dim blnDiagPrinted As Boolean
    Dim strPrevLogLevel As String

    TestRunStart "DUAL-ACCESS EMAIL-INTEGRATION"

    CleanupDualAccessTestData
    SuiteStart "Echter Worker + Email-Job"

    Dim strGrund As String

    If Trim$(strOrdnerPfad) = "" Then
        AssertSkip "Kein Outlook-Ordnerpfad uebergeben"
        SuiteEnd
        GoTo CleanExit
    End If

    strPrevLogLevel = CacheGetConfig(CFG_LOG_LEVEL, CStr(LOG_INFO))
    CacheSetConfig CFG_LOG_LEVEL, CStr(LOG_TRACE)
    g_intLogLevel = LOG_TRACE
    Debug.Print "[EMAIL-TEST] LogLevel fuer Worker-Debug auf TRACE gesetzt"

    dblStep = Timer
    If Not EnsureDualAccessReadyFast() Then
        AssertFail "SetupDualAccessNoAdmin(False, False) fehlgeschlagen"
        SuiteEnd
        GoTo CleanExit
    End If
    Debug.Print "[TIMING] SetupDualAccessReady: " & FormatSekundenDiff(dblStep)

    dblStep = Timer
    If Not WorkerStartVoraussetzungenOK(strGrund, CurrentDb.Name) Then
        WorkerStartVoraussetzungenReport CurrentDb.Name
        AssertSkip "Worker-Start Voraussetzungen nicht erfuellt: " & strGrund
        SuiteEnd
        GoTo CleanExit
    End If
    Debug.Print "[TIMING] WorkerStartVoraussetzungenOK: " & FormatSekundenDiff(dblStep)

    If Not ConnectRDO() Then
        AssertSkip "Outlook/RDO nicht verfuegbar"
        SuiteEnd
        GoTo CleanExit
    End If

    Dim objFolder As Object
    Set objFolder = OeffneOrdner(strOrdnerPfad)
    If objFolder Is Nothing Then
        AssertSkip "Outlook-Ordner nicht gefunden: " & strOrdnerPfad
        SuiteEnd
        GoTo CleanExit
    End If

    Dim lngItems As Long
    On Error Resume Next
    lngItems = CLng(objFolder.Items.Count)
    On Error GoTo ErrHandler
    If lngItems <= 0 Then
        AssertSkip "Outlook-Ordner ist leer: " & strOrdnerPfad
        SuiteEnd
        GoTo CleanExit
    End If

    strWorker = BuildTestWorkerId("MAIL")
    Debug.Print "[EMAIL-TEST] Ordner=" & strOrdnerPfad & " | Items=" & lngItems & " | MaxMails=" & lngMaxMails & " | Subfolders=" & IIf(blnSubfolders, "True", "False")

    dblStep = Timer
    AssertIsTrue WorkerStarten(CurrentDb.Name, strWorker, 10, lngWorkerTaskID), "WorkerStarten bestaetigt Online-Start"
    Debug.Print "[TIMING] WorkerStarten (+WaitForOnline): " & FormatSekundenDiff(dblStep)
    If lngWorkerTaskID > 0 And Not WorkerIstOnline(strWorker) Then
        PrintTestWorkerDiagnostics "Email: Start nicht rechtzeitig bestaetigt", strWorker, 0, 60
        blnDiagPrinted = True
    End If

    dblStep = Timer
    AssertIsTrue WaitForWorkerOnlineStatus(strWorker, 4), "WorkerIstOnline=True"
    Debug.Print "[TIMING] WorkerIstOnline: " & FormatSekundenDiff(dblStep)

    dblStep = Timer
    If Not WaitForBackendWriteReady(strWorker, 15) Then
        Debug.Print "[EMAIL-DIAG] Backend nach Worker-Start nicht schreibbar | WorkerId=" & strWorker
        PrintTestWorkerDiagnostics "Email: Backend nach Start blockiert", strWorker, 0, 80
        blnDiagPrinted = True
        AssertFail "Backend nach Worker-Start nicht schreibbar"
        SuiteEnd
        GoTo CleanExit
    End If
    Debug.Print "[TIMING] BackendWriteReady: " & FormatSekundenDiff(dblStep)

    Dim strCreatedBy As String
    strCreatedBy = TEST_PREFIX & "MAIL_" & Format$(Now, "yyyymmddhhnnss")

    dblStep = Timer
    lngJobID = EnqueueSyncJob(strOrdnerPfad, lngMaxMails, blnSubfolders, strCreatedBy, -32000)
    Debug.Print "[TIMING] EnqueueSyncJob: " & FormatSekundenDiff(dblStep)
    Debug.Print "[EMAIL-TEST] Enqueue Ergebnis | JobID=" & lngJobID & " | CreatedBy=" & strCreatedBy
    AssertGreaterThan lngJobID, 0, "Email-Integrationstest Job erstellt"
    If lngJobID <= 0 Then
        Debug.Print "[EMAIL-DIAG] Enqueue fehlgeschlagen | Pfad=" & strOrdnerPfad & " | WorkerId=" & strWorker
        WorkerStatusReport strWorker
        PrintTestWorkerDiagnostics "Email: Enqueue fehlgeschlagen", strWorker, 0, 80
        blnDiagPrinted = True
        SuiteEnd
        GoTo CleanExit
    End If

    Dim strStatus As String
    dblStep = Timer
    AssertIsTrue WaitForJobStarted(lngJobID, 45, strStatus), "Job wurde vom Worker uebernommen"
    Debug.Print "[TIMING] WaitForJobStarted (<=45s): " & FormatSekundenDiff(dblStep) & " | Status=" & Nz(strStatus, "")
    AssertContains JOB_STATUS_RUNNING & "," & JOB_STATUS_COMPLETED, strStatus, "Jobstatus ist running/completed"

    dblStep = Timer
    AssertIsTrue WaitForJobTerminalVerbose(lngJobID, lngTimeoutSek, strStatus, strWorker, 10), "Job erreicht terminalen Status"
    Debug.Print "[TIMING] WaitForJobTerminal: " & FormatSekundenDiff(dblStep) & " | Status=" & Nz(strStatus, "")
    AssertAreEqual JOB_STATUS_COMPLETED, strStatus, "Email-Job endet erfolgreich mit completed"

    AssertIsTrue (DCount("*", TBL_SYNC_HEARTBEAT, "WorkerId='" & SQLSafe(strWorker) & "'") > 0), "Heartbeat fuer gestarteten Worker vorhanden"
    AssertIsNotEmpty Nz(DLookup("Stage", TBL_SYNC_HEARTBEAT, "WorkerId='" & SQLSafe(strWorker) & "'"), ""), "Heartbeat Stage befuellt"

    Dim lngSyncLaufID As Long
    lngSyncLaufID = HoleSyncLaufIDZuJob(lngJobID)
    AssertGreaterThan lngSyncLaufID, 0, "SyncLauf zu Job gefunden"
    If lngSyncLaufID > 0 Then
        AssertAreEqual "Abgeschlossen", Nz(DLookup("Status", TBL_SYNC_LAUF, "SyncLaufID=" & lngSyncLaufID), ""), "SyncLauf Status = Abgeschlossen"
    End If

    WorkerStatusReport strWorker
    If LCase$(Nz(strStatus, "")) <> JOB_STATUS_COMPLETED Or lngSyncLaufID <= 0 Then
        PrintTestWorkerDiagnostics "Email: Job nicht sauber abgeschlossen", strWorker, lngJobID, 120
        blnDiagPrinted = True
    End If

    SuiteEnd

CleanExit:
    If strPrevLogLevel <> "" Then
        CacheSetConfig CFG_LOG_LEVEL, strPrevLogLevel
        g_intLogLevel = CInt(Val(strPrevLogLevel))
    End If
    If Len(strWorker) > 0 And Not blnDiagPrinted Then
        PrintTestWorkerDiagnostics "Email: Abschlussdiagnose", strWorker, lngJobID, 40
    End If
    SmokeTestCleanupWorkerProcess strWorker, lngWorkerTaskID, 20
    Set objFolder = Nothing
    TestRunEnd
    Exit Sub

ErrHandler:
    If strPrevLogLevel <> "" Then
        CacheSetConfig CFG_LOG_LEVEL, strPrevLogLevel
        g_intLogLevel = CInt(Val(strPrevLogLevel))
    End If
    AssertFail "Email-Integration Fehler: " & Err.Number & " - " & Err.Description
    PrintTestWorkerDiagnostics "Email: ErrHandler", strWorker, lngJobID, 120
    SmokeTestCleanupWorkerProcess strWorker, lngWorkerTaskID, 20
    Set objFolder = Nothing
    SuiteEnd
    TestRunEnd
End Sub


' ---------------------------------------------------------------------------
' SUITE 1: Setup / Schema / Config
' ---------------------------------------------------------------------------
Private Sub Test_DualAccessSetup()
    On Error GoTo ErrHandler
    SuiteStart "DualAccess Setup"

    Dim blnSetup As Boolean
    blnSetup = SetupDualAccessNoAdmin(False)

    AssertIsTrue blnSetup, "SetupDualAccessNoAdmin(False) erfolgreich"
    AssertIsTrue TabelleExistiert(TBL_SYNC_JOB), TBL_SYNC_JOB & " existiert"
    AssertIsTrue TabelleExistiert(TBL_SYNC_HEARTBEAT), TBL_SYNC_HEARTBEAT & " existiert"
    AssertIsTrue TabelleExistiert(TBL_SYNC_CONTROL), TBL_SYNC_CONTROL & " existiert"
    AssertIsTrue TabelleExistiert(TBL_WORKER_LEASE), TBL_WORKER_LEASE & " existiert"

    AssertIsNotEmpty CacheGetConfig(CFG_WORKER_POLL_MS, ""), CFG_WORKER_POLL_MS & " gesetzt"
    AssertIsNotEmpty CacheGetConfig(CFG_WORKER_HB_S, ""), CFG_WORKER_HB_S & " gesetzt"
    AssertIsNotEmpty CacheGetConfig(CFG_WORKER_STALE_S, ""), CFG_WORKER_STALE_S & " gesetzt"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Setup Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ---------------------------------------------------------------------------
' SUITE 2: Enqueue
' ---------------------------------------------------------------------------
Private Sub Test_EnqueueCreatesQueueAndControl()
    On Error GoTo ErrHandler
    SuiteStart "Queue Enqueue"

    Dim strCreatedBy As String
    strCreatedBy = TEST_PREFIX & "USER_" & Format$(Now, "yyyymmddhhnnss")

    Dim lngJobID As Long
    lngJobID = EnqueueSyncJob(TEST_PREFIX & "Mailbox\\Inbox", 42, False, strCreatedBy, -32000)

    Dim strExpectedPath As String
    strExpectedPath = BereinigeOutlookPfad(TEST_PREFIX & "Mailbox\\Inbox")

    AssertGreaterThan lngJobID, 0, "EnqueueSyncJob liefert JobID"
    AssertAreEqual JOB_STATUS_QUEUED, Nz(DLookup("Status", TBL_SYNC_JOB, "JobID=" & lngJobID), ""), "Status nach Enqueue = queued"
    AssertAreEqual strExpectedPath, Nz(DLookup("RequestedFolderPath", TBL_SYNC_JOB, "JobID=" & lngJobID), ""), "FolderPath gespeichert"
    AssertIsTrue (CLng(Nz(DLookup("Priority", TBL_SYNC_JOB, "JobID=" & lngJobID), 0)) <= -1), "Priority <= -1 fuer Testjob"

    AssertIsFalse CBool(Nz(DLookup("PauseRequested", TBL_SYNC_CONTROL, "JobID=" & lngJobID), False)), "ControlRow PauseRequested=False"
    AssertIsFalse CBool(Nz(DLookup("CancelRequested", TBL_SYNC_CONTROL, "JobID=" & lngJobID), False)), "ControlRow CancelRequested=False"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Enqueue Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ---------------------------------------------------------------------------
' SUITE 3: Claim
' ---------------------------------------------------------------------------
Private Sub Test_ClaimNextJob()
    On Error GoTo ErrHandler
    SuiteStart "ClaimNextJob"

    Dim strCreatedBy As String
    strCreatedBy = TEST_PREFIX & "CLAIM_" & Format$(Now, "hhnnss")

    Dim lngJobID As Long
    lngJobID = EnqueueSyncJob(TEST_PREFIX & "Claim\\Ordner", 10, False, strCreatedBy, -32000)
    AssertGreaterThan lngJobID, 0, "Claim-Testjob erstellt"

    Dim strWorker As String
    strWorker = BuildTestWorkerId("CLAIM")

    Dim lngClaimed As Long
    lngClaimed = ClaimNextJob(strWorker)
    AssertGreaterThan lngClaimed, 0, "ClaimNextJob liefert JobID"

    If lngClaimed = lngJobID Then
        AssertAreEqual JOB_STATUS_RUNNING, Nz(DLookup("Status", TBL_SYNC_JOB, "JobID=" & lngClaimed), ""), "Eigener Testjob steht auf running"
        AssertAreEqual strWorker, Nz(DLookup("WorkerId", TBL_SYNC_JOB, "JobID=" & lngClaimed), ""), "WorkerId am Testjob gesetzt"
    Else
        AssertSkip "Claim hat anderen Queue-Job gezogen (globale Queue nicht leer)"
        AssertAreEqual JOB_STATUS_RUNNING, Nz(DLookup("Status", TBL_SYNC_JOB, "JobID=" & lngClaimed), ""), "Geclaimter Job steht auf running"
    End If

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Claim Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ---------------------------------------------------------------------------
' SUITE 4: Heartbeat / Lease
' ---------------------------------------------------------------------------
Private Sub Test_HeartbeatAndLease()
    On Error GoTo ErrHandler
    SuiteStart "Heartbeat + Lease"

    Dim strWorker As String
    strWorker = BuildTestWorkerId("HB")

    HeartbeatUpdate strWorker, 0, "idle", 0, 0, "dual-access-test"

    AssertAreEqual "idle", Nz(DLookup("Stage", TBL_SYNC_HEARTBEAT, "WorkerId='" & SQLSafe(strWorker) & "'"), ""), "Heartbeat Stage gesetzt"
    AssertIsNotEmpty Nz(DLookup("UpdatedAt", TBL_SYNC_HEARTBEAT, "WorkerId='" & SQLSafe(strWorker) & "'"), ""), "Heartbeat UpdatedAt gesetzt"

    AssertIsNotEmpty Nz(DLookup("LeaseUntil", TBL_WORKER_LEASE, "WorkerId='" & SQLSafe(strWorker) & "'"), ""), "WorkerLease gesetzt"
    AssertIsNotEmpty Nz(DLookup("SessionUser", TBL_WORKER_LEASE, "WorkerId='" & SQLSafe(strWorker) & "'"), ""), "WorkerLease SessionUser gesetzt"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Heartbeat/Lease Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ---------------------------------------------------------------------------
' SUITE 5: Cancel Handling
' ---------------------------------------------------------------------------
Private Sub Test_CancelHandling()
    On Error GoTo ErrHandler
    SuiteStart "Pause/Cancel Handling"

    Dim strCreatedBy As String
    strCreatedBy = TEST_PREFIX & "CANCEL_" & Format$(Now, "hhnnss")

    Dim lngJobID As Long
    lngJobID = EnqueueSyncJob(TEST_PREFIX & "Cancel\\Ordner", 5, False, strCreatedBy, -32000)
    AssertGreaterThan lngJobID, 0, "Cancel-Testjob erstellt"

    Dim strWorker As String
    strWorker = BuildTestWorkerId("CANCEL")

    Dim lngClaimed As Long
    lngClaimed = ClaimNextJob(strWorker)
    AssertGreaterThan lngClaimed, 0, "Cancel-Testjob geclaimt oder Queue-Job geclaimt"

    If lngClaimed <> lngJobID Then
        AssertSkip "Cancel-Test nicht ausgefuehrt: Queue enthielt Fremdjob (Sicherheitsabbruch)"
        SuiteEnd
        Exit Sub
    End If

    CurrentDb.Execute "UPDATE [" & TBL_SYNC_CONTROL & "] SET CancelRequested=True, UpdatedAt=" & SQLJetzt() & " WHERE JobID=" & lngClaimed, dbFailOnError

    Dim strFlow As String
    strFlow = HandlePauseCancel(lngClaimed, strWorker)

    AssertAreEqual "canceled", strFlow, "HandlePauseCancel liefert canceled"
    AssertAreEqual JOB_STATUS_CANCEL_REQUESTED, Nz(DLookup("Status", TBL_SYNC_JOB, "JobID=" & lngClaimed), ""), "Status = cancel_requested"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Cancel-Handling Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ---------------------------------------------------------------------------
' SUITE 6: Finalize
' ---------------------------------------------------------------------------
Private Sub Test_FinalizeJob()
    On Error GoTo ErrHandler
    SuiteStart "FinalizeJob"

    Dim strCreatedBy As String
    strCreatedBy = TEST_PREFIX & "FINAL_" & Format$(Now, "hhnnss")

    Dim lngJobID As Long
    lngJobID = EnqueueSyncJob(TEST_PREFIX & "Finalize\\Ordner", 5, False, strCreatedBy, -32000)
    AssertGreaterThan lngJobID, 0, "Finalize-Testjob erstellt"

    Dim strWorker As String
    strWorker = BuildTestWorkerId("FINAL")

    Dim lngClaimed As Long
    lngClaimed = ClaimNextJob(strWorker)
    If lngClaimed <> lngJobID Then
        AssertSkip "Finalize-Test nicht ausgefuehrt: Queue enthielt Fremdjob (Sicherheitsabbruch)"
        SuiteEnd
        Exit Sub
    End If
    AssertGreaterThan lngClaimed, 0, "Finalize-Testjob geclaimt"

    FinalizeJob lngJobID, JOB_STATUS_COMPLETED, "", strWorker

    AssertAreEqual JOB_STATUS_COMPLETED, Nz(DLookup("Status", TBL_SYNC_JOB, "JobID=" & lngJobID), ""), "Status nach Finalize = completed"
    AssertIsNotEmpty Nz(DLookup("FinishedAt", TBL_SYNC_JOB, "JobID=" & lngJobID), ""), "FinishedAt gesetzt"

    AssertAreEqual "finalize", Nz(DLookup("Stage", TBL_SYNC_HEARTBEAT, "WorkerId='" & SQLSafe(strWorker) & "'"), ""), "Heartbeat Stage = finalize"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Finalize Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ---------------------------------------------------------------------------
' CLEANUP HELPERS
' ---------------------------------------------------------------------------
Private Sub CleanupDualAccessTestData()
    On Error Resume Next

    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    Set db = CurrentDb

    ' Test-Jobs iterieren und sauber kaskadiert entfernen
    Set rs = db.OpenRecordset( _
        "SELECT JobID FROM [" & TBL_SYNC_JOB & "] WHERE CreatedBy Like '" & SQLSafe(TEST_PREFIX) & "*'", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        DeleteJobCascade CLng(Nz(rs!JobID, 0))
        rs.MoveNext
    Loop

    rs.Close: Set rs = Nothing

    ' Verbleibende Test-Heartbeat/Lease entfernen
    db.Execute "DELETE FROM [" & TBL_SYNC_HEARTBEAT & "] WHERE WorkerId Like '" & SQLSafe(TEST_PREFIX) & "*'"
    db.Execute "DELETE FROM [" & TBL_WORKER_LEASE & "] WHERE WorkerId Like '" & SQLSafe(TEST_PREFIX) & "*'"
    If TabelleExistiert(TBL_WORKER_TRACE) Then
        db.Execute "DELETE FROM [" & TBL_WORKER_TRACE & "] WHERE WorkerId Like '" & SQLSafe(TEST_PREFIX) & "*'"
    End If

    Set db = Nothing
    On Error GoTo 0
End Sub

Private Sub DeleteJobCascade(ByVal lngJobID As Long)
    On Error Resume Next
    If lngJobID <= 0 Then Exit Sub

    CurrentDb.Execute "DELETE FROM [" & TBL_SYNC_CONTROL & "] WHERE JobID=" & lngJobID
    CurrentDb.Execute "DELETE FROM [" & TBL_SYNC_JOB & "] WHERE JobID=" & lngJobID

    On Error GoTo 0
End Sub

Private Function BuildTestWorkerId(ByVal strSuffix As String) As String
    BuildTestWorkerId = TEST_PREFIX & strSuffix & "_" & Environ("USERNAME") & "_" & Format$(Now, "hhnnss")
End Function

Private Function EnsureDualAccessReadyFast() As Boolean
    On Error GoTo Fallback

    If TabelleExistiert(TBL_SYNC_JOB) _
       And TabelleExistiert(TBL_SYNC_HEARTBEAT) _
       And TabelleExistiert(TBL_SYNC_CONTROL) _
       And TabelleExistiert(TBL_WORKER_LEASE) _
         And WorkerTraceFastReady() _
       And Len(Nz(CacheGetConfig(CFG_WORKER_POLL_MS, ""), "")) > 0 _
       And Len(Nz(CacheGetConfig(CFG_WORKER_HB_S, ""), "")) > 0 _
       And Len(Nz(CacheGetConfig(CFG_WORKER_STALE_S, ""), "")) > 0 Then
        Debug.Print "[SETUP] Skip Migration: Dual-Access bereits bereit"
        EnsureDualAccessReadyFast = True
        Exit Function
    End If

Fallback:
    EnsureDualAccessReadyFast = SetupDualAccessNoAdmin(False, False)
End Function

Private Function WorkerTraceFastReady() As Boolean
    On Error GoTo ErrHandler

    Dim strBackendPfad As String
    Dim db As DAO.Database
    Dim td As DAO.TableDef

    strBackendPfad = Trim$(CacheGetConfig(CFG_BACKEND_PFAD, ""))
    If strBackendPfad = "" Then
        WorkerTraceFastReady = TabelleExistiert(TBL_WORKER_TRACE)
        Exit Function
    End If

    If Not TabelleExistiert(TBL_WORKER_TRACE) Then Exit Function

    Set db = DBEngine.OpenDatabase(strBackendPfad)
    For Each td In db.TableDefs
        If td.Name = TBL_WORKER_TRACE Then
            WorkerTraceFastReady = True
            Exit For
        End If
    Next td

    db.Close
    Set db = Nothing
    Set td = Nothing
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    Set td = Nothing
    WorkerTraceFastReady = False
End Function

Private Function WaitForBackendWriteReady(ByVal strWorkerId As String, _
                                          Optional ByVal lngTimeoutSek As Long = 10) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    Dim lngTry As Long
    Dim strProbeWorkerId As String
    Dim strSqlInsert As String
    Dim strSqlDelete As String

    If lngTimeoutSek < 1 Then lngTimeoutSek = 1
    strProbeWorkerId = TEST_PREFIX & "PROBE_" & Format$(Now, "yyyymmddhhnnss")
    dblStart = Timer

    Do
        lngTry = lngTry + 1
        On Error GoTo WriteErr

        strSqlInsert = "INSERT INTO [" & TBL_WORKER_TRACE & "] (LoggedAt, LogLevel, LevelName, WorkerId, JobID, Modul, Prozedur, Nachricht, Details, HostName, SessionUser) VALUES (" & _
                       SQLJetzt() & ", " & LOG_DEBUG & ", 'DEBUG', '" & SQLSafe(strProbeWorkerId) & "', 0, '" & MODUL_NAME & "', 'WaitForBackendWriteReady', 'backend-probe', 'WorkerId=" & SQLSafe(strWorkerId) & "|Try=" & lngTry & "', '" & SQLSafe(Environ$("COMPUTERNAME")) & "', '" & SQLSafe(Environ$("USERNAME")) & "')"
        CurrentDb.Execute strSqlInsert, dbFailOnError

        strSqlDelete = "DELETE FROM [" & TBL_WORKER_TRACE & "] WHERE WorkerId='" & SQLSafe(strProbeWorkerId) & "'"
        CurrentDb.Execute strSqlDelete, dbFailOnError

        WaitForBackendWriteReady = True
        Exit Function

WriteErr:
        If IsLockError(Err.Number) And SekundenDiff(dblStart, Timer) < lngTimeoutSek Then
            Debug.Print "[BACKEND-READY] Retry | WorkerId=" & strWorkerId & " | Versuch=" & lngTry & " | Err=" & Err.Number & " - " & Err.Description
            Err.Clear
            Sleep 250 * lngTry
            DoEvents
        ElseIf Err.Number <> 0 Then
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Loop While SekundenDiff(dblStart, Timer) < lngTimeoutSek

    Debug.Print "[BACKEND-READY] Timeout | WorkerId=" & strWorkerId & " | Versuche=" & lngTry
    WaitForBackendWriteReady = False
    Exit Function

ErrHandler:
    Debug.Print "[BACKEND-READY] FEHLER | WorkerId=" & strWorkerId & " | Err=" & Err.Number & " - " & Err.Description
    WaitForBackendWriteReady = False
End Function

Private Sub SmokeTestCleanupWorkerProcess(ByVal strWorkerId As String, _
                                          ByVal lngTaskID As Long, _
                                          Optional ByVal lngGraceSek As Long = 12)
    On Error Resume Next

    If Trim$(strWorkerId) = "" Then Exit Sub

    Dim dblStart As Double
    dblStart = Timer

    If lngGraceSek < 1 Then lngGraceSek = 1

    If WorkerStoppen(strWorkerId, lngTaskID, lngGraceSek) Then
        Debug.Print "[WORKER-CLEANUP] Prozess beendet | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID & " | Dauer=" & FormatSekundenDiff(dblStart)
    Else
        Debug.Print "[WORKER-CLEANUP] WARN: Prozess evtl. noch aktiv | WorkerId=" & strWorkerId & " | TaskID=" & lngTaskID & " | Dauer=" & FormatSekundenDiff(dblStart)
    End If

    On Error GoTo 0
End Sub

Private Sub PrintTestWorkerDiagnostics(ByVal strKontext As String, _
                                       ByVal strWorkerId As String, _
                                       Optional ByVal lngJobID As Long = 0, _
                                       Optional ByVal lngTop As Long = 60)
    On Error Resume Next

    If Trim$(strWorkerId) = "" Then Exit Sub

    If lngTop < 1 Then lngTop = 1
    If lngTop > 200 Then lngTop = 200

    Debug.Print String(70, "-")
    Debug.Print "[TEST-DIAG] " & strKontext & " | WorkerId=" & strWorkerId & IIf(lngJobID > 0, " | JobID=" & lngJobID, "")
    If lngJobID > 0 Then PrintJobProgressSnapshot lngJobID, strWorkerId, strKontext
    WorkerStatusReport strWorkerId
    If lngJobID > 0 Then WorkerTraceReport strWorkerId, lngJobID, lngTop, LOG_TRACE
    WorkerTraceReport strWorkerId, 0, lngTop, LOG_TRACE
    Debug.Print String(70, "-")

    On Error GoTo 0
End Sub

Private Function FormatSekundenDiff(ByVal dblStart As Double) As String
    FormatSekundenDiff = Replace(Format$(SekundenDiff(dblStart, Timer), "0.000"), ".", ",") & "s"
End Function

Private Function WaitForJobStarted(ByVal lngJobID As Long, _
                                   ByVal lngTimeoutSek As Long, _
                                   ByRef strStatus As String) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    dblStart = Timer

    Do
        strStatus = SafeJobStatusLookup(lngJobID)
        Select Case LCase$(strStatus)
            Case JOB_STATUS_RUNNING, JOB_STATUS_COMPLETED, JOB_STATUS_FAILED, JOB_STATUS_CANCELED, JOB_STATUS_STALE
                WaitForJobStarted = True
                Exit Function
        End Select

        Sleep 250
        DoEvents
    Loop While SekundenDiff(dblStart, Timer) < lngTimeoutSek

    WaitForJobStarted = False
    Exit Function

ErrHandler:
    strStatus = ""
    WaitForJobStarted = False
End Function

Private Function WaitForWorkerLease(ByVal strWorkerId As String, Optional ByVal lngTimeoutSek As Long = 8) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    dblStart = Timer

    If lngTimeoutSek < 1 Then lngTimeoutSek = 1

    Do
        If SafeWorkerRowExists(TBL_WORKER_LEASE, strWorkerId) Then
            WaitForWorkerLease = True
            Exit Function
        End If

        Sleep 250
        DoEvents
    Loop While SekundenDiff(dblStart, Timer) < lngTimeoutSek

    Debug.Print "[LEASE-DIAG] Timeout ohne Lease | WorkerId=" & strWorkerId & _
                " | HeartbeatUpdatedAt=" & Nz(DLookup("UpdatedAt", TBL_SYNC_HEARTBEAT, "WorkerId='" & SQLSafe(strWorkerId) & "'"), "")

    WaitForWorkerLease = False
    Exit Function

ErrHandler:
    WaitForWorkerLease = False
End Function

Private Function WaitForWorkerHeartbeat(ByVal strWorkerId As String, Optional ByVal lngTimeoutSek As Long = 8) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    dblStart = Timer

    If lngTimeoutSek < 1 Then lngTimeoutSek = 1

    Do
        If SafeWorkerRowExists(TBL_SYNC_HEARTBEAT, strWorkerId) Then
            WaitForWorkerHeartbeat = True
            Exit Function
        End If

        Sleep 250
        DoEvents
    Loop While SekundenDiff(dblStart, Timer) < lngTimeoutSek

    Debug.Print "[HB-DIAG] Timeout ohne Heartbeat | WorkerId=" & strWorkerId

    WaitForWorkerHeartbeat = False
    Exit Function

ErrHandler:
    WaitForWorkerHeartbeat = False
End Function

Private Function WaitForWorkerOnlineStatus(ByVal strWorkerId As String, Optional ByVal lngTimeoutSek As Long = 4) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    dblStart = Timer

    If lngTimeoutSek < 1 Then lngTimeoutSek = 1

    Do
        If WorkerIstOnline(strWorkerId) Then
            WaitForWorkerOnlineStatus = True
            Exit Function
        End If

        Sleep 200
        DoEvents
    Loop While SekundenDiff(dblStart, Timer) < lngTimeoutSek

    WaitForWorkerOnlineStatus = False
    Exit Function

ErrHandler:
    WaitForWorkerOnlineStatus = False
End Function

Private Function SafeWorkerRowExists(ByVal strTable As String, ByVal strWorkerId As String) As Boolean
    On Error GoTo ErrHandler

    Dim lngTry As Long
    Dim lngCnt As Long

    For lngTry = 1 To 5
        On Error GoTo CountErr
        lngCnt = DCount("*", strTable, "WorkerId='" & SQLSafe(strWorkerId) & "'")
        SafeWorkerRowExists = (lngCnt > 0)
        Exit Function

CountErr:
        ' Transiente Lock/Busy Situationen kurz aussitzen.
        If IsLockError(Err.Number) And lngTry < 5 Then
            Err.Clear
            Sleep 120 * lngTry
            DoEvents
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Next lngTry

    SafeWorkerRowExists = False
    Exit Function

ErrHandler:
    SafeWorkerRowExists = False
End Function

Private Function WaitForJobTerminal(ByVal lngJobID As Long, _
                                    ByVal lngTimeoutSek As Long, _
                                    ByRef strStatus As String) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    dblStart = Timer

    Do
        strStatus = SafeJobStatusLookup(lngJobID)
        Select Case LCase$(strStatus)
            Case JOB_STATUS_COMPLETED, JOB_STATUS_FAILED, JOB_STATUS_CANCELED, JOB_STATUS_STALE
                WaitForJobTerminal = True
                Exit Function
        End Select

        Sleep 500
        DoEvents
    Loop While SekundenDiff(dblStart, Timer) < lngTimeoutSek

    WaitForJobTerminal = False
    Exit Function

ErrHandler:
    strStatus = ""
    WaitForJobTerminal = False
End Function

Private Function WaitForJobTerminalVerbose(ByVal lngJobID As Long, _
                                           ByVal lngTimeoutSek As Long, _
                                           ByRef strStatus As String, _
                                           Optional ByVal strWorkerId As String = "", _
                                           Optional ByVal lngReportIntervallSek As Long = 10) As Boolean
    On Error GoTo ErrHandler

    Dim dblStart As Double
    Dim dblLastReport As Double

    If lngReportIntervallSek < 2 Then lngReportIntervallSek = 2

    dblStart = Timer
    dblLastReport = dblStart

    Do
        strStatus = SafeJobStatusLookup(lngJobID)
        Select Case LCase$(strStatus)
            Case JOB_STATUS_COMPLETED, JOB_STATUS_FAILED, JOB_STATUS_CANCELED, JOB_STATUS_STALE
                WaitForJobTerminalVerbose = True
                Exit Function
        End Select

        If SekundenDiff(dblLastReport, Timer) >= lngReportIntervallSek Then
            PrintJobProgressSnapshot lngJobID, strWorkerId, "WaitForJobTerminal"
            dblLastReport = Timer
        End If

        Sleep 500
        DoEvents
    Loop While SekundenDiff(dblStart, Timer) < lngTimeoutSek

    PrintJobProgressSnapshot lngJobID, strWorkerId, "WaitForJobTerminal: Timeout"
    WaitForJobTerminalVerbose = False
    Exit Function

ErrHandler:
    strStatus = ""
    WaitForJobTerminalVerbose = False
End Function

Private Sub PrintJobProgressSnapshot(ByVal lngJobID As Long, _
                                     Optional ByVal strWorkerId As String = "", _
                                     Optional ByVal strKontext As String = "")
    On Error Resume Next

    Dim strJobCrit As String
    Dim strWorkerCrit As String
    Dim strJobStatus As String
    Dim strJobWorker As String
    Dim strHbWorker As String
    Dim strHbStage As String
    Dim strHbMessage As String
    Dim varStarted As Variant
    Dim varFinished As Variant
    Dim varHbUpdated As Variant
    Dim varLeaseUntil As Variant
    Dim lngCurrent As Long
    Dim lngTotal As Long
    Dim lngSyncLaufID As Long
    Dim strSyncStatus As String
    Dim lngSyncFehler As Long
    Dim strLastError As String

    If lngJobID <= 0 Then Exit Sub

    strJobCrit = "JobID=" & lngJobID
    strJobStatus = Nz(DLookup("Status", TBL_SYNC_JOB, strJobCrit), "")
    strJobWorker = Nz(DLookup("WorkerId", TBL_SYNC_JOB, strJobCrit), "")
    varStarted = DLookup("StartedAt", TBL_SYNC_JOB, strJobCrit)
    varFinished = DLookup("FinishedAt", TBL_SYNC_JOB, strJobCrit)
    strLastError = Nz(DLookup("LastError", TBL_SYNC_JOB, strJobCrit), "")

    If Trim$(strWorkerId) = "" Then strWorkerId = strJobWorker
    If Trim$(strWorkerId) <> "" Then
        strWorkerCrit = "WorkerId='" & SQLSafe(strWorkerId) & "'"
        strHbWorker = Nz(DLookup("WorkerId", TBL_SYNC_HEARTBEAT, strWorkerCrit), "")
        strHbStage = Nz(DLookup("Stage", TBL_SYNC_HEARTBEAT, strWorkerCrit), "")
        strHbMessage = Nz(DLookup("LastMessage", TBL_SYNC_HEARTBEAT, strWorkerCrit), "")
        varHbUpdated = DLookup("UpdatedAt", TBL_SYNC_HEARTBEAT, strWorkerCrit)
        varLeaseUntil = DLookup("LeaseUntil", TBL_WORKER_LEASE, strWorkerCrit)
        lngCurrent = CLng(Nz(DLookup("CurrentItem", TBL_SYNC_HEARTBEAT, strWorkerCrit), 0))
        lngTotal = CLng(Nz(DLookup("TotalItems", TBL_SYNC_HEARTBEAT, strWorkerCrit), 0))
    End If

    lngSyncLaufID = HoleSyncLaufIDZuJob(lngJobID)
    If lngSyncLaufID > 0 Then
        strSyncStatus = Nz(DLookup("Status", TBL_SYNC_LAUF, "SyncLaufID=" & lngSyncLaufID), "")
        lngSyncFehler = CLng(Nz(DLookup("AnzahlFehler", TBL_SYNC_LAUF, "SyncLaufID=" & lngSyncLaufID), 0))
    End If

    Debug.Print "[JOB-PROGRESS] " & IIf(Trim$(strKontext) = "", "Snapshot", strKontext) & _
                " | JobID=" & lngJobID & _
                " | Status=" & Nz(strJobStatus, "") & _
                " | WorkerId=" & Nz(strJobWorker, "") & _
                " | StartedAt=" & Nz(varStarted, "") & _
                " | FinishedAt=" & Nz(varFinished, "")

    If Trim$(strLastError) <> "" Then
        Debug.Print "[JOB-PROGRESS] LastError | " & Left$(strLastError, 255)
    End If

    If strHbWorker <> "" Then
        Debug.Print "[JOB-PROGRESS] Heartbeat | WorkerId=" & strHbWorker & _
                    " | Stage=" & strHbStage & _
                    " | Current=" & lngCurrent & "/" & lngTotal & _
                    " | UpdatedAt=" & Nz(varHbUpdated, "") & _
                    " | LeaseUntil=" & Nz(varLeaseUntil, "")
        If Trim$(strHbMessage) <> "" Then
            Debug.Print "[JOB-PROGRESS] HeartbeatMsg | " & Left$(strHbMessage, 255)
        End If
    ElseIf Trim$(strWorkerId) <> "" Then
        Debug.Print "[JOB-PROGRESS] Heartbeat | keine Zeile fuer WorkerId=" & strWorkerId
    End If

    Debug.Print "[JOB-PROGRESS] SyncLauf | SyncLaufID=" & lngSyncLaufID & _
                " | Status=" & Nz(strSyncStatus, "") & _
                " | Fehler=" & lngSyncFehler

    On Error GoTo 0
End Sub

Private Function SafeJobStatusLookup(ByVal lngJobID As Long) As String
    On Error GoTo ErrHandler

    Dim lngTry As Long

    For lngTry = 1 To 5
        On Error GoTo LookupErr
        SafeJobStatusLookup = Nz(DLookup("Status", TBL_SYNC_JOB, "JobID=" & lngJobID), "")
        Exit Function

LookupErr:
    If IsLockError(Err.Number) And lngTry < 5 Then
            Err.Clear
            Sleep 120 * lngTry
            DoEvents
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Next lngTry

    SafeJobStatusLookup = ""
    Exit Function

ErrHandler:
    SafeJobStatusLookup = ""
End Function

Private Function HoleSyncLaufIDZuJob(ByVal lngJobID As Long) As Long
    On Error Resume Next

    Dim strPhase As String
    strPhase = "Job" & Format$(lngJobID, "000000")
    HoleSyncLaufIDZuJob = CLng(Nz(DMax("SyncLaufID", TBL_SYNC_LAUF, _
        "Projekt='DualAccess' AND Phase='" & SQLSafe(strPhase) & "'"), 0))

    If Err.Number <> 0 Then HoleSyncLaufIDZuJob = 0: Err.Clear
    On Error GoTo 0
End Function


