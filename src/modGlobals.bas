Attribute VB_Name = "modGlobals"
Option Compare Database
Option Explicit

' ===========================================================================
' modGlobals - Globale Deklarationen, Konstanten und Initialisierung
' ===========================================================================
' Zentrale Stelle fuer alle projektweiten Konstanten, MAPI-Tags,
' globale Objekte und die Init/Cleanup-Logik.
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
Public Const RDO_DLL_64 As String = "D:\Redemption64.dll"
Public Const RDO_DLL_32 As String = "D:\Redemption.dll"

' ---------------------------------------------------------------------------
' DEFAULTS
' ---------------------------------------------------------------------------
Public Const DEFAULT_NAME   As String = "Unbekannt"
Public Const DEFAULT_EMAIL  As String = "unbekannt@example.com"
Public Const MAX_RETRIES    As Integer = 3

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

    ' Log-Level aus Config lesen (Fallback: INFO)
    g_intLogLevel = CInt(LeseConfig("LogLevel", CStr(LOG_INFO)))
    g_blnAbbrechen = False
    g_dtSyncStart = 0

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
