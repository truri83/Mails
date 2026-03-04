Attribute VB_Name = "modTransactionManager"
Option Compare Database
Option Explicit

' ===========================================================================
' modTransactionManager - Zentrale Transaktions-Steuerung
' ===========================================================================
' VERSION: 1.0 (adaptiert aus mdl_Transaction_Manager v2.1)
'
' Zentrale Steuerung von DB-Zugriffen und Transaktionen.
' Verhindert "Error 3034" und DB-Locks durch Clean-Up-Routinen.
'
' FEATURES:
'   - Logically Flattened Transactions (Tiefenzaehler)
'   - Doom-Flag: Inner Rollback → aeusserer Rollback
'   - ForceCleanup: Loop-basierter Rollback bis Error 3034
'   - Auto-Recovery nach kritischen Fehlern
'   - Safe SQL Execution mit Retry-Logik
'   - TLookup/TCount/TSum (schneller als DLookup/DCount/DSum)
'   - Lock-Error-Erkennung fuer Retry-Entscheidung
'
' AUFRUF:
'   modTransactionManager.BeginTransaction
'   ... Datenoperationen ...
'   modTransactionManager.CommitTransaction    ' oder RollbackTransaction
'
'   ' Sichere DB-Referenz:
'   Dim db As DAO.Database
'   Set db = modTransactionManager.GetSafeDB()
'
'   ' Schnelle Lookups:
'   Dim lngAnzahl As Long
'   lngAnzahl = modTransactionManager.TCount("tblEmails", "SyncLaufID=5")
'
' WICHTIG:
'   Im Error-Handler von Batch-Prozessen IMMER aufrufen:
'     modTransactionManager.ForceCleanup
' ===========================================================================

Private Const MODUL_NAME As String = "modTransactionManager"


' ---------------------------------------------------------------------------
' API DEKLARATIONEN
' ---------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


' ---------------------------------------------------------------------------
' MODUL-VARIABLEN
' ---------------------------------------------------------------------------

' Transaktions-Status
Private m_iTransDepth       As Integer      ' Verschachtelungstiefe (logisch, nicht physisch)
Private m_bTransactionDoomed As Boolean     ' Doom-Flag: Inner Rollback → Global Rollback
Private m_bRecoveryMode     As Boolean      ' Recovery laeuft gerade
Private m_dbCurrent         As DAO.Database ' Singleton DB-Referenz

' Recovery-Flag: Wird bei kritischem Fehler gesetzt
Private m_bNeedsRecovery    As Boolean


' ===========================================================================
' 1. DATABASE ACCESS (SINGLETON)
' ===========================================================================

' Liefert eine sichere DB-Referenz (erstellt sie bei Bedarf neu).
' Prueft automatisch auf Recovery-Bedarf.
Public Function GetSafeDB() As DAO.Database
    On Error GoTo ErrHandler

    ' Auto-Recovery pruefen
    If m_bNeedsRecovery Then
        If Not CheckAndRecover() Then
            Err.Raise vbObjectError + 30003, "GetSafeDB", _
                "Auto-Recovery fehlgeschlagen"
        End If
    End If

    ' Objekt-Validierung
    If m_dbCurrent Is Nothing Then
        Set m_dbCurrent = CurrentDb
    Else
        ' Testzugriff: Ist die Referenz noch gueltig?
        On Error Resume Next
        Dim sName As String
        sName = m_dbCurrent.Name
        If Err.Number <> 0 Then
            Set m_dbCurrent = CurrentDb
            Err.Clear
        End If
        On Error GoTo ErrHandler
    End If

    Set GetSafeDB = m_dbCurrent
    Exit Function

ErrHandler:
    ' Bei kritischem Fehler: Recovery + EINMALIGER Retry
    If Not m_bNeedsRecovery Then
        Debug.Print "[TransMgr] GetSafeDB FEHLER - markiere Recovery: " & Err.Description
        MarkForRecovery

        On Error Resume Next
        If CheckAndRecover() Then
            Set m_dbCurrent = CurrentDb
            Set GetSafeDB = m_dbCurrent
            If Err.Number = 0 Then
                Debug.Print "[TransMgr] GetSafeDB Retry erfolgreich"
                Exit Function
            End If
        End If
        On Error GoTo 0
    End If

    Err.Raise Err.Number, "GetSafeDB", Err.Description
End Function


' ===========================================================================
' 2. TRANSACTION MANAGEMENT (LOGICALLY FLATTENED)
' ===========================================================================

' Startet eine logische Transaktion.
' Nur die ERSTE (aeusserste) Ebene startet eine physische DAO-Transaktion.
' Innere Aufrufe erhoehen nur den Tiefenzaehler.
Public Sub BeginTransaction()
    On Error GoTo ErrHandler

    ' Auto-Recovery bei neuer aeusserer Transaktion
    If m_bNeedsRecovery And m_iTransDepth = 0 Then
        CheckAndRecover
    End If

    ' Physische Transaktion nur auf Ebene 0
    If m_iTransDepth = 0 Then
        DBEngine.Workspaces(0).BeginTrans
        m_bTransactionDoomed = False
    End If

    m_iTransDepth = m_iTransDepth + 1
    Exit Sub

ErrHandler:
    Debug.Print "[TransMgr] BeginTransaction FEHLER: " & Err.Description
    MarkForRecovery
End Sub


' Gibt True zurueck wenn eine Transaktion offen ist (Tiefe > 0)
Public Function IsTransactionOpen() As Boolean
    IsTransactionOpen = (m_iTransDepth > 0)
End Function


' Committed die logische Transaktion.
' Nur die LETZTE (aeusserste) Ebene fuehrt physisches Commit/Rollback aus.
' Bei Doom-Flag: Rollback statt Commit + Err.Raise
Public Sub CommitTransaction()
    On Error GoTo ErrHandler

    If m_iTransDepth > 0 Then
        m_iTransDepth = m_iTransDepth - 1

        If m_iTransDepth = 0 Then
            ' Oberste Ebene: Commit oder Doom-Rollback
            If m_bTransactionDoomed Then
                m_bTransactionDoomed = False
                DBEngine.Workspaces(0).Rollback
                Debug.Print "[TransMgr] DOOMED: Global Rollback statt Commit"
                Err.Raise vbObjectError + 22001, MODUL_NAME & ".CommitTransaction", _
                    "Transaction DOOMED by inner rollback -> Global Rollback"
            Else
                DBEngine.Workspaces(0).CommitTrans
            End If
        End If
    End If
    Exit Sub

ErrHandler:
    ' Commit fehlgeschlagen: Tiefe zuruecksetzen + physisch aufraeuemen
    ' Ohne Dekrement wuerde der naechste BeginTransaction keine neue physische
    ' Transaktion oeffnen und CommitTransaction wuerde versuchen, eine
    ' bereits abgebrochene Transaktion zu committen.
    If m_iTransDepth > 0 Then m_iTransDepth = m_iTransDepth - 1
    If m_iTransDepth = 0 Then
        On Error Resume Next
        DBEngine.Workspaces(0).Rollback
        On Error GoTo 0
    End If
    Debug.Print "[TransMgr] CommitTransaction FEHLER: " & Err.Description
    MarkForRecovery
End Sub


' Markiert die aktuelle Transaktion fuer Rollback.
' Innere Ebenen setzen den Doom-Flag; physischer Rollback erst auf Ebene 0.
Public Sub RollbackTransaction()
    On Error GoTo ErrHandler

    If m_iTransDepth > 0 Then
        m_bTransactionDoomed = True
        m_iTransDepth = m_iTransDepth - 1

        If m_iTransDepth = 0 Then
            ' Aeusserste Ebene: Jetzt wirklich Rollback
            DBEngine.Workspaces(0).Rollback
        End If
    End If
    Exit Sub

ErrHandler:
    Dim sDesc As String: sDesc = Err.Description
    Dim lNum As Long: lNum = Err.Number
    ' Tiefe hart zuruecksetzen + physischen Rollback erzwingen
    m_iTransDepth = 0
    m_bTransactionDoomed = False
    On Error Resume Next
    DBEngine.Workspaces(0).Rollback
    On Error GoTo 0
    Debug.Print "[TransMgr] RollbackTransaction FEHLER: [" & lNum & "] " & sDesc
End Sub


' ===========================================================================
' 3. SAFETY & CLEANUP
' ===========================================================================

' Setzt ALLES zurueck: Schliesst alle offenen Transaktionen (Rollback)
' und gibt die DB-Variable frei.
' WICHTIG: Im Error-Handler von Batch-Prozessen IMMER aufrufen!
Public Sub ForceCleanup()
    On Error Resume Next

    Debug.Print "[TransMgr] ForceCleanup: TransDepth=" & m_iTransDepth & _
                " Doomed=" & m_bTransactionDoomed

    ' 1. Alle Transaktionsebenen abbauen via Loop-Rollback
    '    (bis Error 3034 = Keine Transaktion aktiv)
    Dim i As Integer
    For i = 1 To 10  ' Sicherheitsbremse
        DBEngine.Workspaces(0).Rollback
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next i

    ' 2. Counter zuruecksetzen
    m_iTransDepth = 0
    m_bTransactionDoomed = False
    m_bRecoveryMode = False
    m_bNeedsRecovery = False

    ' 3. DB-Objekt freigeben (wird bei naechstem GetSafeDB neu erstellt)
    Set m_dbCurrent = Nothing

    Debug.Print "[TransMgr] ForceCleanup abgeschlossen"
End Sub


' Markiert fuer Recovery (nach Ueberlauf, haengenden Transaktionen etc.)
Public Sub MarkForRecovery()
    m_bNeedsRecovery = True
    Debug.Print "[TransMgr] Recovery-Flag gesetzt (TransDepth=" & m_iTransDepth & ")"
End Sub


' Prueft ob Recovery noetig ist und fuehrt es durch.
' Gibt True zurueck wenn danach alles OK ist.
Public Function CheckAndRecover() As Boolean
    On Error Resume Next

    ' Nichts zu tun?
    If Not m_bNeedsRecovery And m_iTransDepth = 0 Then
        CheckAndRecover = True
        Exit Function
    End If

    ' SOFORT Flag zuruecksetzen um Rekursion zu vermeiden
    Dim bNeeded As Boolean
    bNeeded = m_bNeedsRecovery
    m_bNeedsRecovery = False

    If Not bNeeded And m_iTransDepth = 0 Then
        CheckAndRecover = True
        Exit Function
    End If

    Debug.Print "[TransMgr] Auto-Recovery: NeedsRecovery=" & bNeeded & _
                " TransDepth=" & m_iTransDepth

    ' Recovery durchfuehren
    ForceCleanup

    ' Kurze Pause fuer DB-Stabilisierung
    Sleep 500

    ' Test ob DB wieder verfuegbar (Flag ist jetzt False → keine Rekursion)
    Dim db As DAO.Database
    Set db = GetSafeDB()
    If db Is Nothing Then
        Debug.Print "[TransMgr] Recovery FEHLGESCHLAGEN - DB nicht verfuegbar"
        CheckAndRecover = False
        Exit Function
    End If

    Debug.Print "[TransMgr] Recovery ERFOLGREICH"
    CheckAndRecover = True
End Function


' Gibt zurueck ob Recovery noetig ist
Public Function NeedsRecovery() As Boolean
    NeedsRecovery = m_bNeedsRecovery Or _
                    (m_iTransDepth > 0 And m_bTransactionDoomed)
End Function


' ===========================================================================
' 4. SQL EXECUTION (SAFE WRAPPERS MIT RETRY)
' ===========================================================================

' Fuehrt eine Aktionsabfrage (INSERT/UPDATE/DELETE) mit Retry-Logik aus.
' Bei Lock-Fehlern: Automatische Wiederholung mit optionalem User-Feedback.
' Bei Ueberlauf/kritischen Fehlern: Auto-Recovery (via GetSafeDB)
Public Sub ExecuteActionQuery(ByVal sSQL As String, _
                               Optional ByVal withUserFeedback As Boolean = True, _
                               Optional ByVal maxRetries As Integer = 3)
    On Error GoTo ErrHandler

    If withUserFeedback Then
        ExecuteWithRetryAndFeedback sSQL, maxRetries
    Else
        ExecuteWithRetrySimple sSQL, maxRetries
    End If
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case 6  ' Ueberlauf
            MarkForRecovery
            Debug.Print "[TransMgr] ExecuteActionQuery OVERFLOW: " & Left(sSQL, 100)
        Case 5  ' Ungueltiger Prozeduraufruf
            MarkForRecovery
            Debug.Print "[TransMgr] ExecuteActionQuery ERR=5: " & Err.Description
        Case 3061  ' Parameter-Fehler
            Debug.Print "[TransMgr] SQL Parameter Mismatch: " & Left(sSQL, 200)
        Case Else
            Debug.Print "[TransMgr] ExecuteActionQuery ERR=" & Err.Number & _
                        ": " & Err.Description & " | SQL=" & Left(sSQL, 200)
    End Select
    Err.Raise Err.Number, "ExecuteActionQuery", _
        Err.Description & " | SQL: " & Left(sSQL, 200)
End Sub


' Interne Retry-Logik MIT Benutzer-Feedback (fuer interaktive Operationen)
Private Sub ExecuteWithRetryAndFeedback(ByVal sSQL As String, _
                                         ByVal maxRetries As Integer)
    Dim db As DAO.Database
    Dim attempt As Integer
    Dim lastError As String
    Dim waitMs As Long
    Dim maxAttempts As Integer

    Set db = GetSafeDB()

    maxAttempts = maxRetries
    attempt = 0

    ' Do-While statt For-To: Obere Grenze kann durch User-Retry erweitert werden
    Do While attempt < maxAttempts
        attempt = attempt + 1
        On Error Resume Next
        Err.Clear
        db.Execute sSQL, dbFailOnError

        If Err.Number = 0 Then Exit Sub  ' Erfolg!

        lastError = Err.Description

        ' Nur bei Lock-Fehlern Retry sinnvoll
        If Not IsLockError(Err.Number) Then
            On Error GoTo 0
            Err.Raise Err.Number, "ExecuteActionQuery", _
                Err.Description & " | SQL: " & Left(sSQL, 100)
        End If

        ' Bei letztem Versuch: User fragen
        If attempt >= maxAttempts Then
            Dim response As VbMsgBoxResult
            response = MsgBox( _
                "Datenbank-Operation fehlgeschlagen:" & vbCrLf & vbCrLf & _
                lastError & vbCrLf & vbCrLf & _
                "Die Datenbank ist moeglicherweise durch einen anderen " & _
                "Benutzer gesperrt." & vbCrLf & vbCrLf & _
                "Vorgang erneut versuchen?", _
                vbRetryCancel + vbExclamation, _
                "Datenbank-Sperre")

            If response = vbRetry Then
                maxAttempts = maxAttempts + 3
                Debug.Print "[TransMgr] User waehlt Retry (neue Versuche: " & _
                            maxAttempts & ")"
            Else
                On Error GoTo 0
                Err.Raise Err.Number, "ExecuteActionQuery", _
                    "Benutzer abgebrochen: " & lastError
            End If
        End If

        ' Exponential Backoff (max 2 Sekunden)
        waitMs = 200 * (2 ^ (attempt - 1))
        If waitMs > 2000 Then waitMs = 2000

        Debug.Print "[TransMgr] Retry " & attempt & "/" & maxAttempts & _
                    " in " & waitMs & "ms"
        Sleep waitMs
    Loop

    ' Alle Versuche fehlgeschlagen
    On Error GoTo 0
    Err.Raise Err.Number, "ExecuteActionQuery", lastError
End Sub


' Interne Retry-Logik OHNE Benutzer-Feedback (fuer Hintergrund-Operationen)
Private Sub ExecuteWithRetrySimple(ByVal sSQL As String, _
                                    ByVal maxRetries As Integer)
    Dim db As DAO.Database
    Dim attempt As Integer

    Set db = GetSafeDB()

    For attempt = 1 To maxRetries
        On Error Resume Next
        Err.Clear
        db.Execute sSQL, dbFailOnError

        If Err.Number = 0 Then Exit Sub

        If Not IsLockError(Err.Number) Then
            On Error GoTo 0
            Err.Raise Err.Number, "ExecuteActionQuery", _
                Err.Description & " | SQL: " & Left(sSQL, 100)
        End If

        If attempt < maxRetries Then
            Dim waitMs As Long
            waitMs = 200 * (2 ^ (attempt - 1))
            If waitMs > 2000 Then waitMs = 2000

            Debug.Print "[TransMgr] Retry " & attempt & "/" & maxRetries & _
                        " (kein User-Prompt)"
            Sleep waitMs
        End If
    Next attempt

    ' Alle Versuche fehlgeschlagen
    On Error GoTo 0
    Err.Raise Err.Number, "ExecuteActionQuery", _
        Err.Description & " | SQL: " & Left(sSQL, 100)
End Sub


' ===========================================================================
' 5. LOCK-ERROR-ERKENNUNG
' ===========================================================================

' Prueft ob ein Fehler ein Lock-/Sperr-Problem ist (retry-faehig)
Public Function IsLockError(ByVal errNum As Long) As Boolean
    Select Case errNum
        Case 3188  ' Aktualisierung nicht moeglich (gesperrt)
        Case 3197  ' Daten wurden von anderem Benutzer geaendert
        Case 3211  ' Datenbank kann nicht gesperrt werden
        Case 3218  ' Konnte Sperre nicht abrufen
        Case 3260  ' Tabelle aktuell gesperrt
        Case 3300  ' Beziehung verhindert Update
        Case 3622  ' Konflikt bei gleichzeitigem Zugriff
        Case 3704  ' Datenbank exklusiv gesperrt
        Case 3734  ' Datenbank muss exklusiv geoeffnet sein
        Case 3008  ' Tabelle exklusiv geoeffnet
        Case Else
            IsLockError = False
            Exit Function
    End Select
    IsLockError = True
End Function


' ===========================================================================
' 6. PERFORMANCE LOOKUPS (schneller als DLookup/DCount/DSum)
' ===========================================================================

' Schneller Ersatz fuer DLookup
'   Beispiel: TLookup("Wert", "tblConfig", "Schluessel='BackendPfad'")
Public Function TLookup(ByVal sField As String, _
                         ByVal sDomain As String, _
                         Optional ByVal sCriteria As String = "") As Variant
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sSQL As String

    sSQL = "SELECT TOP 1 " & sField & " FROM " & sDomain
    If Len(sCriteria) > 0 Then sSQL = sSQL & " WHERE " & sCriteria

    Set rs = GetSafeDB().OpenRecordset(sSQL, dbOpenSnapshot)
    If Not rs.EOF Then
        TLookup = rs(0)
    Else
        TLookup = Null
    End If
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next: If Not rs Is Nothing Then rs.Close: On Error GoTo 0
    TLookup = Null
End Function


' Schneller Ersatz fuer DCount
'   Beispiel: TCount("tblEmails", "SyncLaufID=5")
Public Function TCount(ByVal sDomain As String, _
                        Optional ByVal sCriteria As String = "") As Long
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sSQL As String

    sSQL = "SELECT Count(*) FROM " & sDomain
    If Len(sCriteria) > 0 Then sSQL = sSQL & " WHERE " & sCriteria

    Set rs = GetSafeDB().OpenRecordset(sSQL, dbOpenSnapshot)
    If Not rs.EOF Then
        TCount = Nz(rs(0), 0)
    Else
        TCount = 0
    End If
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next: If Not rs Is Nothing Then rs.Close: On Error GoTo 0
    TCount = 0
End Function


' Schneller Ersatz fuer DSum
'   Beispiel: TSum("Groesse", "tblEmails", "SyncLaufID=5")
Public Function TSum(ByVal sField As String, _
                      ByVal sDomain As String, _
                      Optional ByVal sCriteria As String = "") As Currency
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sSQL As String

    sSQL = "SELECT Sum(" & sField & ") FROM " & sDomain
    If Len(sCriteria) > 0 Then sSQL = sSQL & " WHERE " & sCriteria

    Set rs = GetSafeDB().OpenRecordset(sSQL, dbOpenSnapshot)
    If Not rs.EOF Then
        TSum = Nz(rs(0), 0)
    Else
        TSum = 0
    End If
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next: If Not rs Is Nothing Then rs.Close: On Error GoTo 0
    TSum = 0
End Function


' Schneller Ersatz fuer DMax
Public Function TMax(ByVal sField As String, _
                      ByVal sDomain As String, _
                      Optional ByVal sCriteria As String = "") As Variant
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sSQL As String

    sSQL = "SELECT Max(" & sField & ") FROM " & sDomain
    If Len(sCriteria) > 0 Then sSQL = sSQL & " WHERE " & sCriteria

    Set rs = GetSafeDB().OpenRecordset(sSQL, dbOpenSnapshot)
    If Not rs.EOF Then
        TMax = rs(0)
    Else
        TMax = Null
    End If
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next: If Not rs Is Nothing Then rs.Close: On Error GoTo 0
    TMax = Null
End Function


' Schneller Abruf eines einzelnen Wertes per beliebigem SQL
Public Function ExecuteScalar(ByVal sSQL As String) As Variant
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset

    Set rs = GetSafeDB().OpenRecordset(sSQL, dbOpenSnapshot)
    If Not rs.EOF Then
        ExecuteScalar = rs(0)
    Else
        ExecuteScalar = Null
    End If
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next: If Not rs Is Nothing Then rs.Close: On Error GoTo 0
    ExecuteScalar = Null
End Function


' Scalar als Long (0 bei leer/NULL/Fehler)
Public Function GetScalarLong(ByVal db As DAO.Database, _
                               ByVal sSQL As String) As Long
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
    If Not rs.EOF Then GetScalarLong = Nz(rs.Fields(0).Value, 0)
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next: If Not rs Is Nothing Then rs.Close: On Error GoTo 0
    GetScalarLong = 0
End Function


' Scalar als String ("" bei leer/NULL/Fehler)
Public Function GetScalarString(ByVal db As DAO.Database, _
                                 ByVal sSQL As String) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
    If Not rs.EOF Then GetScalarString = Nz(rs.Fields(0).Value, "")
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next: If Not rs Is Nothing Then rs.Close: On Error GoTo 0
    GetScalarString = ""
End Function
