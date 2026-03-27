Option Compare Database
Option Explicit

' ===========================================================================
' modGlobals - Globale Deklarationen, Konstanten und Initialisierung
' ===========================================================================
' Zentrale Stelle fuer alle projektweiten Konstanten, MAPI-Tags,
' globale Objekte und die Init/Cleanup-Logik.
'
' v0.5: TBL_*/CFG_*/PATH_*-Konstanten, keine hardcodierten Namen mehr
' ===========================================================================

' ---------------------------------------------------------------------------
' WINDOWS API
' ---------------------------------------------------------------------------
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ---------------------------------------------------------------------------
' LOG-LEVEL KONSTANTEN
' ---------------------------------------------------------------------------
Public Const LOG_NONE   As Integer = 0
Public Const LOG_ERROR  As Integer = 1
Public Const LOG_WARN   As Integer = 2
Public Const LOG_INFO   As Integer = 3
Public Const LOG_DEBUG  As Integer = 4
Public Const LOG_TRACE  As Integer = 5

' ---------------------------------------------------------------------------
' MAPI PROPERTY-TAGS (fuer RDOItem.Fields[] und PropertyAccessor)
' ---------------------------------------------------------------------------
Public Const PR_SUBJECT                     As Long = &H37001E
Public Const PR_TRANSPORT_MESSAGE_HEADERS   As Long = &H7D001E
Public Const PR_INTERNET_MESSAGE_ID         As Long = &H1035001E
Public Const PR_MESSAGE_SIZE                As Long = &HE080003
Public Const PR_SENDER_EMAIL_ADDRESS        As Long = &HC1F001E
Public Const PR_SENDER_NAME                 As Long = &HC1A001E
Public Const PR_DISPLAY_TO                  As Long = &HE04001E
Public Const PR_DISPLAY_CC                  As Long = &HE03001E
Public Const PR_SMTP_ADDRESS                As Long = &H39FE001E

' ---------------------------------------------------------------------------
' ATTACHMENT-TYPEN
' ---------------------------------------------------------------------------
Public Const ATT_BY_VALUE   As Integer = 1   ' Echte Datei
Public Const ATT_OLE        As Integer = 5   ' Eingebettetes OLE-Objekt
Public Const ATT_EMBEDDED   As Integer = 6   ' Weitergeleitete Mail als Anlage

' ---------------------------------------------------------------------------
' OOM FOLDER-IDS
' ---------------------------------------------------------------------------
Public Const olFolderInbox          As Integer = 6
Public Const olFolderSentMail       As Integer = 5
Public Const olFolderDeletedItems   As Integer = 4
Public Const olFolderDrafts         As Integer = 16
Public Const olFolderCalendar       As Integer = 9
Public Const olFolderContacts       As Integer = 10

' ---------------------------------------------------------------------------
' OUTLOOK KONSTANTEN
' ---------------------------------------------------------------------------
Public Const olMSG      As Integer = 3
Public Const olMail     As Integer = 43

' ---------------------------------------------------------------------------
' REDEMPTION DLL-PFADE
' ---------------------------------------------------------------------------
Public Const RDO_DLL_64 As String = "Redemption64.dll"
Public Const RDO_DLL_32 As String = "Redemption.dll"

' ---------------------------------------------------------------------------
' DEFAULTS
' ---------------------------------------------------------------------------
Public Const DEFAULT_NAME   As String = "Unbekannt"
Public Const DEFAULT_EMAIL  As String = "unbekannt@example.com"
Public Const MAX_RETRIES    As Integer = 3

' ---------------------------------------------------------------------------
' TABELLEN-KONSTANTEN (TBL_*)
' ---------------------------------------------------------------------------
Public Const TBL_CONFIG             As String = "tblConfig"
Public Const TBL_SYNC_LAUF         As String = "tblSyncLauf"
Public Const TBL_KONTAKTE          As String = "tblKontakte"
Public Const TBL_OUTLOOK_ORDNER    As String = "tblOutlookOrdner"
Public Const TBL_EMAIL_THREADS     As String = "tblEmailThreads"
Public Const TBL_EMAILS            As String = "tblEmails"
Public Const TBL_EMAIL_CONTENT     As String = "tblEmailContent"
Public Const TBL_EMAIL_EMPFAENGER  As String = "tblEmailEmpfaenger"
Public Const TBL_EMAIL_ANHAENGE    As String = "tblEmailAnhaenge"
Public Const TBL_EMAIL_STATUS      As String = "tblEmailStatus"
Public Const TBL_SYNC_PROFIL       As String = "tblSyncProfil"
Public Const TBL_SYNC_PROFIL_ORDNER As String = "tblSyncProfilOrdner"
Public Const TBL_LOG               As String = "tblLog"
Public Const TBL_PROJEKTE          As String = "tblProjekte"
Public Const TBL_EMAIL_PROJEKT     As String = "tblEmailProjekt"
Public Const TBL_SYNC_JOB          As String = "tblSyncJob"
Public Const TBL_SYNC_HEARTBEAT    As String = "tblSyncHeartbeat"
Public Const TBL_SYNC_CONTROL      As String = "tblSyncControl"
Public Const TBL_WORKER_LEASE      As String = "tblWorkerLease"
Public Const TBL_WORKER_TRACE      As String = "tblWorkerTrace"

' ---------------------------------------------------------------------------
' CONFIG-KEY-KONSTANTEN (CFG_*)
' ---------------------------------------------------------------------------
Public Const CFG_LOG_LEVEL          As String = "LogLevel"
Public Const CFG_EXPORT_PFAD       As String = "ExportBasisPfad"
Public Const CFG_MAX_MAILS         As String = "MaxMailsProSync"
Public Const CFG_ANHAENGE          As String = "AnhaengeExtrahieren"
Public Const CFG_MSG_EXPORT        As String = "MSGExportieren"
Public Const CFG_SIGNATUR_FILTER   As String = "SignaturBilderFiltern"
Public Const CFG_SCHEMA_VERSION    As String = "SchemaVersion"
Public Const CFG_BACKEND_PFAD      As String = "BackendPfad"
Public Const CFG_TEMP_PFAD         As String = "TempPfad"
Public Const CFG_BUFFER_GROESSE    As String = "BufferGroesse"
Public Const CFG_NETZWERK_RETRIES  As String = "NetzwerkRetries"
Public Const CFG_NETZWERK_PAUSE    As String = "NetzwerkRetryPause"
Public Const CFG_NETZWERK_TIMEOUT  As String = "NetzwerkTimeoutSek"
Public Const CFG_RDO_PFAD          As String = "RedemptionPfad"
Public Const CFG_DEV_MODUS         As String = "DevModus"
Public Const CFG_EINGANG_ORDNER    As String = "EingangOrdnerName"
Public Const CFG_WORKER_POLL_MS    As String = "WorkerPollMs"
Public Const CFG_WORKER_HB_S       As String = "WorkerHeartbeatSek"
Public Const CFG_WORKER_STALE_S    As String = "WorkerStaleSek"

' ---------------------------------------------------------------------------
' SYNC-JOB STATUS-KONSTANTEN
' ---------------------------------------------------------------------------
Public Const JOB_STATUS_QUEUED          As String = "queued"
Public Const JOB_STATUS_RUNNING         As String = "running"
Public Const JOB_STATUS_PAUSE_REQUESTED As String = "pause_requested"
Public Const JOB_STATUS_PAUSED          As String = "paused"
Public Const JOB_STATUS_CANCEL_REQUESTED As String = "cancel_requested"
Public Const JOB_STATUS_COMPLETED       As String = "completed"
Public Const JOB_STATUS_FAILED          As String = "failed"
Public Const JOB_STATUS_STALE           As String = "stale"
Public Const JOB_STATUS_CANCELED        As String = "canceled"

' ---------------------------------------------------------------------------
' PROJEKT-STATUS-KONSTANTEN
' ---------------------------------------------------------------------------
Public Const PROJ_STATUS_AKTIV      As String = "Aktiv"
Public Const PROJ_STATUS_ARCHIVIERT As String = "Archiviert"
Public Const PROJ_STATUS_GESPERRT   As String = "Gesperrt"

' ---------------------------------------------------------------------------
' EMAIL-PROJEKT ZUORDNUNGS-QUELLEN
' ---------------------------------------------------------------------------
Public Const EP_QUELLE_AUTO         As String = "AutoSync"
Public Const EP_QUELLE_MANUELL      As String = "Manuell"
Public Const EP_QUELLE_MIGRATION    As String = "Migration"

' ---------------------------------------------------------------------------
' EMAIL-STATUS-KONSTANTEN
' ---------------------------------------------------------------------------
Public Const EMAIL_STATUS_NEU          As String = "Neu"
Public Const EMAIL_STATUS_VERARBEITET  As String = "Verarbeitet"
Public Const EMAIL_STATUS_ARCHIVIERT   As String = "Archiviert"
Public Const EMAIL_STATUS_IRRELEVANT   As String = "Irrelevant"

' ---------------------------------------------------------------------------
' PFAD-DEFAULTS (PATH_*)
' ---------------------------------------------------------------------------
Public Const PATH_DEFAULT_FALLBACK As String = "\OutlookSync\"
Public Const BE_DEV_DATEINAME      As String = "OutlookSync_BE_DEV.accdb"

' ---------------------------------------------------------------------------
' GLOBALE OBJEKTE
' ---------------------------------------------------------------------------
Public g_objOutlook     As Object   ' Outlook.Application
Public g_objRDO         As Object   ' Redemption.RDOSession
Public g_intLogLevel    As Integer  ' Aktives Log-Level (Standard: LOG_INFO=3)
Public g_blnAbbrechen   As Boolean  ' Abbruch-Flag fuer Sync
Public g_dtSyncStart    As Double   ' Timer-Wert bei Sync-Start (fuer Dauer/Countdown)


' ===========================================================================
' INITIALISIERUNG
' ===========================================================================
Public Sub InitGlobals()
    On Error GoTo ErrHandler

    ' Cache-System initialisieren (vor Config-Zugriff)
    CacheInit

    ' Log-Level aus Config lesen (ueber Cache fuer Performance)
    g_intLogLevel = CInt(CacheGetConfig(CFG_LOG_LEVEL, CStr(LOG_INFO)))
    g_blnAbbrechen = False
    g_blnBackendOffline = False
    g_dtSyncStart = 0

    ' Jet/ACE Timeouts optimieren (Netzwerk-Dialog-Praevention)
    BackendOptimierTimeouts

    ' Backend-Watchdog starten: Form-Timer bevorzugt, Fallback SetTimer API
    StarteFormWatchdog

    ' Crash-Recovery: Pruefen ob letzter Sync unsauber beendet wurde
    ZeigeCrashRecovery

    LogInfo "InitGlobals: LogLevel=" & g_intLogLevel
    Exit Sub

ErrHandler:
    ' Fallback wenn tblConfig noch nicht existiert
    g_intLogLevel = LOG_INFO
    g_blnAbbrechen = False
    g_dtSyncStart = 0
    Debug.Print "[WARN] InitGlobals: " & Err.Number & " - " & Err.Description & " (Defaults verwendet)"
End Sub


' ===========================================================================
' AUFRAEUMEN
' ===========================================================================
Public Sub CleanupGlobals()
    On Error Resume Next

    ' Watchdog stoppen (Form + OnTime)
    StoppeFormWatchdog
    StoppBackendWatchdog

    ' Cache leeren (vor COM-Cleanup)
    CacheReset

    If Not g_objRDO Is Nothing Then
        g_objRDO.Logoff
        Set g_objRDO = Nothing
    End If

    Set g_objOutlook = Nothing
    g_blnAbbrechen = False
    g_dtSyncStart = 0

    On Error GoTo 0
    LogInfo "CleanupGlobals: Ressourcen freigegeben"
End Sub

