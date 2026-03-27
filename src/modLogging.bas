Attribute VB_Name = "modLogging"
Option Compare Database
Option Explicit

' ===========================================================================
' modLogging - Zentrale Logging-Infrastruktur
' ===========================================================================
' Alle Log-Ausgaben gehen an:
'   1. Debug.Print (Direktbereich)
'   2. Logfile (tagesbasiert: Logfile_yyyymmdd.txt im DB-Verzeichnis)
'
' Kurzfunktionen: LogInfo, LogWarn, LogError, LogDebug, LogTrace
' Fehler-Logging: LogVBAError (mit Err.Number + Err.Description)
'
' Log-Level wird ueber g_intLogLevel gesteuert (aus tblConfig.LogLevel).
' ===========================================================================

Private Const LOG_FILE_PREFIX As String = "Logfile_"


' ---------------------------------------------------------------------------
' ZENTRALE LOG-FUNKTION
' ---------------------------------------------------------------------------
Public Sub SchreibeLog(ByVal strNachricht As String, _
                       Optional ByVal iLevel As Integer = 3, _
                       Optional ByVal strKategorie As String = "")

    ' Log-Level pruefen
    If iLevel > g_intLogLevel Then Exit Sub

    Dim strPrefix   As String
    Dim strEintrag  As String

    Select Case iLevel
        Case LOG_ERROR: strPrefix = "[ERROR]"
        Case LOG_WARN:  strPrefix = "[WARN ]"
        Case LOG_INFO:  strPrefix = "[INFO ]"
        Case LOG_DEBUG: strPrefix = "[DEBUG]"
        Case LOG_TRACE: strPrefix = "[TRACE]"
        Case Else:      strPrefix = "[     ]"
    End Select

    If strKategorie <> "" Then
        strEintrag = Format(Now, "hh:nn:ss") & " " & strPrefix & " [" & strKategorie & "] " & strNachricht
    Else
        strEintrag = Format(Now, "hh:nn:ss") & " " & strPrefix & " " & strNachricht
    End If

    ' 1. Debug.Print (Direktbereich)
    Debug.Print strEintrag

    ' 2. Logfile (tagesbasiert)
    Call SchreibeLogDatei(strEintrag)
End Sub


' ---------------------------------------------------------------------------
' KURZFUNKTIONEN
' ---------------------------------------------------------------------------
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


' ---------------------------------------------------------------------------
' VBA-FEHLER LOGGEN (mit Err-Objekt-Daten)
' ---------------------------------------------------------------------------
Public Sub LogVBAError(ByVal strKontext As String, _
                       Optional ByVal lngErrNum As Long = 0, _
                       Optional ByVal strErrDesc As String = "")
    If lngErrNum = 0 Then lngErrNum = Err.Number
    If strErrDesc = "" Then strErrDesc = Err.Description

    LogError strKontext & " - Err " & lngErrNum & ": " & strErrDesc
End Sub


' ---------------------------------------------------------------------------
' IN DATEI SCHREIBEN (tagesbasiert)
' ---------------------------------------------------------------------------
Private Sub SchreibeLogDatei(ByVal strText As String)
    On Error Resume Next
    Dim strPath As String
    Dim f       As Integer

    strPath = CurrentProject.Path & "\" & LOG_FILE_PREFIX & Format(Date, "yyyymmdd") & ".txt"
    f = FreeFile
    Open strPath For Append As #f
    Print #f, strText
    Close #f
End Sub
