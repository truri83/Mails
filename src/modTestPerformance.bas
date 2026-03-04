Attribute VB_Name = "modTestPerformance"
Option Compare Database
Option Explicit

' ===========================================================================
' modTestPerformance - Performance-Tests fuer Architektur-Entscheidungen
' ===========================================================================
' Testet systematisch alle Performance-relevanten Aspekte:
'
'   BEREICH 1: DB-WRITE Performance (Netzwerk-Backend)
'     - Linked Tables: Batch-Groessen 1/10/25/50/100 (mit Transaktion)
'     - Direct DAO.OpenDatabase: Gleiche Batch-Groessen
'     - Schreibmethoden: Recordset.AddNew vs Execute INSERT vs QueryDef
'     - Memo-Feld-Performance (1KB / 50KB / 500KB HTML Body)
'
'   BEREICH 2: DB-READ Performance (Hash-Lookup, DLookup)
'     - Batch-Hash-Pruefung per SELECT ... WHERE Hash IN (...)
'     - DLookup vs Recordset-Lookup vs DCount
'
'   BEREICH 3: OUTLOOK-EXTRAKTION (OOM vs Redemption)
'     - OOM Basic (Subject, Date, Size - kein Security Issue)
'     - OOM + Body/HTML (grosse Properties)
'     - Redemption SafeMailItem (alle Felder, kein Security Prompt)
'     - Redemption RDO (direkter MAPI-Zugriff)
'     - Extract-Release Timing (COM-Haltezeit messen)
'
'   BEREICH 4: SHA256-HASHING
'     - CryptoAPI Durchsatz (verschiedene Eingabegroessen)
'
' AUFRUF IM DIREKTFENSTER:
'   TestAllePerformance "\\Server\Share\Pfad\"
'   TestAllePerformance "\\Server\Share\Pfad\", 100    ' 100 Test-Records
'
' EINZELTESTS:
'   TestDBPerformance "\\Server\Share\Pfad\"            ' Nur DB
'   TestOutlookExtraktion                               ' Nur Outlook
'   TestOutlookExtraktion 100                           ' 100 Mails testen
'   TestHashingSpeed                                    ' Nur SHA256
'
' Abhaengigkeiten: Keine (komplett eigenstaendig)
' ===========================================================================


' ---------------------------------------------------------------------------
' WINDOWS API DEKLARATIONEN
' ---------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    ' High-Resolution Timer (Sub-Millisekunde)
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" _
        (lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" _
        (lpFrequency As Currency) As Long

    ' CryptoAPI fuer SHA256
    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" _
        Alias "CryptAcquireContextA" _
        (ByRef phProv As LongPtr, ByVal pszContainer As String, _
         ByVal pszProvider As String, ByVal dwProvType As Long, _
         ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" _
        (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, _
         ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long
    Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" _
        (ByVal hHash As LongPtr, ByRef pbData As Byte, ByVal dwDataLen As Long, _
         ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" _
        (ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Byte, _
         ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" _
        (ByVal hHash As LongPtr) As Long
    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" _
        (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Private Declare Function QueryPerformanceCounter Lib "kernel32" _
        (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" _
        (lpFrequency As Currency) As Long

    Private Declare Function CryptAcquireContext Lib "advapi32.dll" _
        Alias "CryptAcquireContextA" _
        (ByRef phProv As Long, ByVal pszContainer As String, _
         ByVal pszProvider As String, ByVal dwProvType As Long, _
         ByVal dwFlags As Long) As Long
    Private Declare Function CryptCreateHash Lib "advapi32.dll" _
        (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, _
         ByVal dwFlags As Long, ByRef phHash As Long) As Long
    Private Declare Function CryptHashData Lib "advapi32.dll" _
        (ByVal hHash As Long, ByRef pbData As Byte, ByVal dwDataLen As Long, _
         ByVal dwFlags As Long) As Long
    Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
        (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Byte, _
         ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
    Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
        (ByVal hHash As Long) As Long
    Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
        (ByVal hProv As Long, ByVal dwFlags As Long) As Long
#End If

' CryptoAPI Konstanten
Private Const PROV_RSA_AES          As Long = 24
Private Const CRYPT_VERIFYCONTEXT   As Long = &HF0000000
Private Const CALG_SHA_256          As Long = &H800C
Private Const HP_HASHVAL            As Long = 2


' ---------------------------------------------------------------------------
' MODUL-VARIABLEN
' ---------------------------------------------------------------------------
Private m_strNetzwerkPfad   As String   ' Netzwerk-Basispfad (User-Eingabe)
Private m_strTestDBPfad     As String   ' Voller Pfad zur Test-Backend-DB
Private m_lngAnzahlRecords  As Long     ' Anzahl Test-Records pro Test
Private m_curFrequency      As Currency ' QPC Frequenz fuer Zeitmessung
Private m_lngBaselineLinked As Long     ' Baseline-Zeit Linked Tables (ohne Trans)
Private m_lngBaselineDirect As Long     ' Baseline-Zeit Direct DAO (ohne Trans)
Private m_lngLetzteHashZeit As Long     ' Letzte Hash-Benchmark-Zeit (fuer Bewertung)


' ===========================================================================
' HAUPT-EINSTIEGSPUNKT
' ===========================================================================

' Alle Performance-Tests durchfuehren.
'   strNetzwerkPfad: Netzlaufwerk-Pfad fuer DB-Tests (Pflicht fuer DB-Tests)
'   lngAnzahlRecords: Anzahl Test-Records (Standard: 100)
Public Sub TestAllePerformance(Optional ByVal strNetzwerkPfad As String = "", _
                                Optional ByVal lngAnzahlRecords As Long = 100)
    On Error GoTo ErrHandler

    ' QPC Frequenz initialisieren
    QueryPerformanceFrequency m_curFrequency

    Debug.Print String(80, "=")
    Debug.Print "  PERFORMANCE-TESTS FUER ARCHITEKTUR-ENTSCHEIDUNGEN"
    Debug.Print String(80, "=")
    Debug.Print "  Zeitpunkt   : " & Now()
#If VBA7 Then
    Debug.Print "  VBA7/64-Bit : Ja"
#Else
    Debug.Print "  VBA7/64-Bit : Nein"
#End If
    Debug.Print "  QPC Freq    : " & Format(m_curFrequency * 10000, "#,##0") & " Hz"
    Debug.Print String(80, "=")
    Debug.Print ""

    ' --- BEREICH 1+2: DB-Performance ---
    If strNetzwerkPfad <> "" Then
        TestDBPerformance strNetzwerkPfad, lngAnzahlRecords
    Else
        Debug.Print "*** DB-Tests uebersprungen (kein Netzwerkpfad angegeben)"
        Debug.Print "    Aufruf mit: TestAllePerformance ""\\Server\Share\Pfad\"""
        Debug.Print ""
    End If

    ' --- BEREICH 3: Outlook-Extraktion ---
    Debug.Print ""
    TestOutlookExtraktion 50

    ' --- BEREICH 4: SHA256-Hashing ---
    Debug.Print ""
    TestHashingSpeed

    ' --- ZUSAMMENFASSUNG ---
    Debug.Print ""
    Debug.Print String(80, "=")
    Debug.Print "  ALLE TESTS ABGESCHLOSSEN - " & Now()
    Debug.Print String(80, "=")
    Exit Sub

ErrHandler:
    Debug.Print "*** FEHLER in TestAllePerformance: " & Err.Description
End Sub


' ===========================================================================
' BEREICH 1+2: DATENBANK-PERFORMANCE
' ===========================================================================

' Hauptroutine fuer alle DB-Performance-Tests
Public Sub TestDBPerformance(Optional ByVal strNetzwerkPfad As String = "", _
                              Optional ByVal lngAnzahlRecords As Long = 100)
    On Error GoTo ErrHandler

    ' QPC Frequenz initialisieren (falls Einzelaufruf)
    QueryPerformanceFrequency m_curFrequency

    ' Pfad validieren
    If strNetzwerkPfad = "" Then
        Debug.Print "*** FEHLER: Netzwerkpfad erforderlich!"
        Debug.Print "    Aufruf: TestDBPerformance ""\\Server\Share\Pfad\"""
        Exit Sub
    End If

    m_strNetzwerkPfad = strNetzwerkPfad
    If Right(m_strNetzwerkPfad, 1) <> "\" Then m_strNetzwerkPfad = m_strNetzwerkPfad & "\"

    m_lngAnzahlRecords = lngAnzahlRecords
    If m_lngAnzahlRecords < 10 Then m_lngAnzahlRecords = 10
    If m_lngAnzahlRecords > 1000 Then m_lngAnzahlRecords = 1000

    Debug.Print String(80, "=")
    Debug.Print "  DATENBANK-PERFORMANCE TESTS"
    Debug.Print String(80, "=")
    Debug.Print "  Netzwerkpfad : " & m_strNetzwerkPfad
    Debug.Print "  Test-Records : " & m_lngAnzahlRecords
    Debug.Print ""

    ' --- Schritt 1: Test-Backend erstellen ---
    Debug.Print "Erstelle Test-Backend auf Netzwerk..."
    If Not ErstelleTestBackend() Then
        Debug.Print "*** FEHLER: Test-Backend konnte nicht erstellt werden! Abbruch."
        Exit Sub
    End If
    Debug.Print "  -> " & m_strTestDBPfad
    Debug.Print ""

    ' --- Schritt 2: Linked Table Tests ---
    Debug.Print "Verknuepfe Test-Tabellen..."
    If VerknuepeTestTabellen() Then
        Debug.Print ""
        TestBatchSchreiben_Linked
        Debug.Print ""
        TestSchreibMethoden_Linked
    Else
        Debug.Print "*** Linked-Table-Tests uebersprungen (Verknuepfung fehlgeschlagen)"
    End If

    ' Verknuepfungen wieder entfernen
    EntferneTestVerknuepfungen
    Debug.Print ""

    ' --- Schritt 3: Direct DAO Tests ---
    TestBatchSchreiben_Direct
    Debug.Print ""

    ' --- Schritt 4: Memo-Feld Tests ---
    TestMemoSchreiben
    Debug.Print ""

    ' --- Schritt 5: Hash-Lookup Tests ---
    TestHashLookup
    Debug.Print ""

    ' --- Schritt 6: Ergebnis-Tabelle ---
    ' (wird innerhalb jedes Tests bereits ausgegeben)

    ' --- Aufraeumen ---
    Debug.Print "Loesche Test-Backend..."
    LoescheTestBackend
    Debug.Print "DB-Tests abgeschlossen."
    Debug.Print String(80, "=")
    Exit Sub

ErrHandler:
    Debug.Print "*** FEHLER in TestDBPerformance: " & Err.Description
    On Error Resume Next
    EntferneTestVerknuepfungen
    LoescheTestBackend
    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' TEST-BACKEND ERSTELLEN (temporaere .accdb auf Netzwerk)
' ---------------------------------------------------------------------------
Private Function ErstelleTestBackend() As Boolean
    On Error GoTo ErrHandler

    m_strTestDBPfad = m_strNetzwerkPfad & "PerfTest_" & _
                      Format(Now, "yyyymmdd_hhnnss") & ".accdb"

    ' Datenbank erstellen
    Dim ws As DAO.Workspace
    Dim dbBE As DAO.Database

    Set ws = DBEngine.Workspaces(0)
    Set dbBE = ws.CreateDatabase(m_strTestDBPfad, dbLangGeneral)

    ' --- Tabelle: tblTestEmails (simuliert tblEmails) ---
    dbBE.Execute "CREATE TABLE tblTestEmails (" & _
        "ID AUTOINCREMENT PRIMARY KEY, " & _
        "Hash TEXT(64), " & _
        "Betreff TEXT(255), " & _
        "Absender TEXT(255), " & _
        "AbsenderEmail TEXT(255), " & _
        "DatumGesendet DATETIME, " & _
        "DatumEmpfangen DATETIME, " & _
        "Groesse LONG, " & _
        "AnhangAnzahl INTEGER, " & _
        "OrdnerID LONG, " & _
        "SyncLaufID LONG, " & _
        "ErstelltAm DATETIME)"

    ' Index auf Hash (fuer Duplikatpruefung)
    dbBE.Execute "CREATE INDEX idxHash ON tblTestEmails (Hash)"

    ' --- Tabelle: tblTestContent (simuliert tblEmailContent mit Memo) ---
    dbBE.Execute "CREATE TABLE tblTestContent (" & _
        "ID AUTOINCREMENT PRIMARY KEY, " & _
        "EmailID LONG, " & _
        "BodyHTML MEMO, " & _
        "BodyPlain MEMO, " & _
        "ErstelltAm DATETIME)"

    ' --- Tabelle: tblTestEmpfaenger (simuliert tblEmailEmpfaenger) ---
    dbBE.Execute "CREATE TABLE tblTestEmpfaenger (" & _
        "ID AUTOINCREMENT PRIMARY KEY, " & _
        "EmailID LONG, " & _
        "Typ TEXT(10), " & _
        "EmpfName TEXT(255), " & _
        "EmpfEmail TEXT(255))"

    dbBE.Close
    Set dbBE = Nothing
    Set ws = Nothing

    ErstelleTestBackend = True
    Exit Function

ErrHandler:
    Debug.Print "  ErstelleTestBackend FEHLER: " & Err.Description
    ErstelleTestBackend = False
End Function


' ---------------------------------------------------------------------------
' TEST-TABELLEN VERKNUEPFEN (fuer Linked-Table-Tests)
' ---------------------------------------------------------------------------
Private Function VerknuepeTestTabellen() As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef

    Set db = CurrentDb

    ' tblPerfTest_Emails -> tblTestEmails
    Set td = db.CreateTableDef("tblPerfTest_Emails")
    td.Connect = ";DATABASE=" & m_strTestDBPfad
    td.SourceTableName = "tblTestEmails"
    db.TableDefs.Append td

    ' tblPerfTest_Content -> tblTestContent
    Set td = db.CreateTableDef("tblPerfTest_Content")
    td.Connect = ";DATABASE=" & m_strTestDBPfad
    td.SourceTableName = "tblTestContent"
    db.TableDefs.Append td

    ' tblPerfTest_Empfaenger -> tblTestEmpfaenger
    Set td = db.CreateTableDef("tblPerfTest_Empfaenger")
    td.Connect = ";DATABASE=" & m_strTestDBPfad
    td.SourceTableName = "tblTestEmpfaenger"
    db.TableDefs.Append td

    db.TableDefs.Refresh
    Set db = Nothing

    Debug.Print "  -> 3 Tabellen verknuepft (tblPerfTest_*)"
    VerknuepeTestTabellen = True
    Exit Function

ErrHandler:
    Debug.Print "  VerknuepeTestTabellen FEHLER: " & Err.Description
    VerknuepeTestTabellen = False
End Function


' ---------------------------------------------------------------------------
' TEST-VERKNUEPFUNGEN ENTFERNEN
' ---------------------------------------------------------------------------
Private Sub EntferneTestVerknuepfungen()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb

    db.TableDefs.Delete "tblPerfTest_Emails"
    db.TableDefs.Delete "tblPerfTest_Content"
    db.TableDefs.Delete "tblPerfTest_Empfaenger"
    db.TableDefs.Refresh

    Set db = Nothing
    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' TEST: Batch-Schreiben ueber LINKED TABLES (verschiedene Batch-Groessen)
' ---------------------------------------------------------------------------
Private Sub TestBatchSchreiben_Linked()
    On Error GoTo ErrHandler

    Debug.Print "--- LINKED TABLE WRITES (Recordset.AddNew + Transaction) ---"
    Debug.Print ""
    Debug.Print PadR("Batch-Groesse", 16) & PadR("Records", 10) & _
                PadR("Zeit", 12) & PadR("Pro Record", 12) & "Faktor"
    Debug.Print String(62, "-")

    ' Baseline: Ohne Transaktion (Batch=1)
    Dim lngBaseline As Long
    lngBaseline = RunBatchWriteTest_Linked(1, False)

    ' Mit Transaktion: verschiedene Batch-Groessen
    Dim lngZeit As Long

    lngZeit = RunBatchWriteTest_Linked(10, True)
    lngZeit = RunBatchWriteTest_Linked(25, True)
    lngZeit = RunBatchWriteTest_Linked(50, True)
    lngZeit = RunBatchWriteTest_Linked(100, True)

    Debug.Print String(62, "-")
    Debug.Print ""
    Exit Sub

ErrHandler:
    Debug.Print "  FEHLER: " & Err.Description
End Sub


' Einzelner Batch-Write-Test ueber Linked Table
' Gibt die Gesamtzeit in ms zurueck
Private Function RunBatchWriteTest_Linked(ByVal lngBatchGroesse As Long, _
                                           ByVal blnTransaction As Boolean) As Long
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rs As DAO.Recordset
    Dim t1 As Long, t2 As Long
    Dim n As Long, lngBatchCount As Long

    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)

    ' Tabelle leeren
    db.Execute "DELETE FROM tblPerfTest_Emails", dbFailOnError

    ' Schreiben
    Set rs = db.OpenRecordset("tblPerfTest_Emails", dbOpenDynaset)

    t1 = GetTickCount()
    lngBatchCount = 0

    If blnTransaction Then ws.BeginTrans

    For n = 1 To m_lngAnzahlRecords
        rs.AddNew
        rs!Hash = GeneriereTestHash(n)
        rs!Betreff = "Testmail Nr. " & n & " - Performance-Messung Batch=" & lngBatchGroesse
        rs!Absender = "Testuser " & (n Mod 50)
        rs!AbsenderEmail = "test" & (n Mod 50) & "@example.com"
        rs!DatumGesendet = DateAdd("h", -n, Now)
        rs!DatumEmpfangen = DateAdd("h", -n + 0.1, Now)
        rs!Groesse = 10000 + (n * 137)
        rs!AnhangAnzahl = n Mod 5
        rs!OrdnerID = 1
        rs!SyncLaufID = 1
        rs!ErstelltAm = Now
        rs.Update

        lngBatchCount = lngBatchCount + 1

        ' Batch-Grenze: Commit + neuer Trans
        If blnTransaction And lngBatchCount >= lngBatchGroesse Then
            ws.CommitTrans
            DoEvents
            ws.BeginTrans
            lngBatchCount = 0
        End If
    Next n

    If blnTransaction And lngBatchCount > 0 Then ws.CommitTrans

    t2 = GetTickCount()
    rs.Close: Set rs = Nothing
    Set db = Nothing: Set ws = Nothing

    ' Ergebnis ausgeben
    Dim lngZeit As Long: lngZeit = t2 - t1
    Dim strBatch As String
    If blnTransaction Then
        strBatch = CStr(lngBatchGroesse) & " (Trans)"
    Else
        strBatch = "1 (kein Trans)"
    End If

    Debug.Print PadR(strBatch, 16) & _
                PadR(CStr(m_lngAnzahlRecords), 10) & _
                PadR(FormatMS(lngZeit), 12) & _
                PadR(FormatProRec(lngZeit, m_lngAnzahlRecords), 12) & _
                IIf(m_lngBaselineLinked > 0 And lngZeit > 0, _
                    Format(CDbl(m_lngBaselineLinked) / CDbl(lngZeit), "0.0") & "x", _
                    "Baseline")

    ' Baseline merken
    If Not blnTransaction Then m_lngBaselineLinked = lngZeit

    RunBatchWriteTest_Linked = lngZeit
    Exit Function

ErrHandler:
    On Error Resume Next
    If blnTransaction Then ws.Rollback
    Debug.Print "  FEHLER bei Batch=" & lngBatchGroesse & ": " & Err.Description
    RunBatchWriteTest_Linked = -1
End Function


' ---------------------------------------------------------------------------
' TEST: Schreibmethoden ueber LINKED TABLE (Recordset vs Execute vs QueryDef)
' ---------------------------------------------------------------------------
Private Sub TestSchreibMethoden_Linked()
    On Error GoTo ErrHandler

    Debug.Print "--- LINKED TABLE: SCHREIBMETHODEN-VERGLEICH (Batch=25, Trans) ---"
    Debug.Print ""
    Debug.Print PadR("Methode", 24) & PadR("Records", 10) & _
                PadR("Zeit", 12) & PadR("Pro Record", 12) & "Info"
    Debug.Print String(70, "-")

    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)

    Dim t1 As Long, t2 As Long
    Dim n As Long, lngBatch As Long
    Dim lngBatchGroesse As Long: lngBatchGroesse = 25

    ' --- Methode A: Recordset.AddNew ---
    db.Execute "DELETE FROM tblPerfTest_Emails", dbFailOnError
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblPerfTest_Emails", dbOpenDynaset)

    t1 = GetTickCount()
    lngBatch = 0
    ws.BeginTrans
    For n = 1 To m_lngAnzahlRecords
        rs.AddNew
        rs!Hash = GeneriereTestHash(n)
        rs!Betreff = "Recordset Test " & n
        rs!Absender = "User " & n
        rs!AbsenderEmail = "user" & n & "@test.de"
        rs!DatumGesendet = Now
        rs!Groesse = 50000
        rs!ErstelltAm = Now
        rs.Update
        lngBatch = lngBatch + 1
        If lngBatch >= lngBatchGroesse Then
            ws.CommitTrans: DoEvents: ws.BeginTrans
            lngBatch = 0
        End If
    Next n
    If lngBatch > 0 Then ws.CommitTrans
    t2 = GetTickCount()
    rs.Close: Set rs = Nothing

    Debug.Print PadR("Recordset.AddNew", 24) & _
                PadR(CStr(m_lngAnzahlRecords), 10) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, m_lngAnzahlRecords), 12) & _
                "Aktueller Ansatz in modDAO"

    ' --- Methode B: db.Execute INSERT ---
    db.Execute "DELETE FROM tblPerfTest_Emails", dbFailOnError

    t1 = GetTickCount()
    lngBatch = 0
    ws.BeginTrans
    For n = 1 To m_lngAnzahlRecords
        db.Execute "INSERT INTO tblPerfTest_Emails " & _
            "(Hash, Betreff, Absender, AbsenderEmail, DatumGesendet, Groesse, ErstelltAm) " & _
            "VALUES ('" & GeneriereTestHash(n) & "', " & _
            "'Execute Test " & n & "', " & _
            "'User " & n & "', " & _
            "'user" & n & "@test.de', " & _
            "#" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "#, " & _
            "50000, " & _
            "#" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "#)", dbFailOnError
        lngBatch = lngBatch + 1
        If lngBatch >= lngBatchGroesse Then
            ws.CommitTrans: DoEvents: ws.BeginTrans
            lngBatch = 0
        End If
    Next n
    If lngBatch > 0 Then ws.CommitTrans
    t2 = GetTickCount()

    Debug.Print PadR("Execute INSERT SQL", 24) & _
                PadR(CStr(m_lngAnzahlRecords), 10) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, m_lngAnzahlRecords), 12) & _
                "SQL-String pro Record"

    ' --- Methode C: QueryDef mit Parametern ---
    db.Execute "DELETE FROM tblPerfTest_Emails", dbFailOnError

    Dim qd As DAO.QueryDef
    Set qd = db.CreateQueryDef("", _
        "PARAMETERS [pHash] Text(64), [pBetreff] Text(255), " & _
        "[pAbsender] Text(255), [pEmail] Text(255), " & _
        "[pDatum] DateTime, [pGroesse] Long, [pErstellt] DateTime; " & _
        "INSERT INTO tblPerfTest_Emails " & _
        "(Hash, Betreff, Absender, AbsenderEmail, DatumGesendet, Groesse, ErstelltAm) " & _
        "VALUES ([pHash], [pBetreff], [pAbsender], [pEmail], [pDatum], [pGroesse], [pErstellt])")

    t1 = GetTickCount()
    lngBatch = 0
    ws.BeginTrans
    For n = 1 To m_lngAnzahlRecords
        qd.Parameters("pHash") = GeneriereTestHash(n)
        qd.Parameters("pBetreff") = "QueryDef Test " & n
        qd.Parameters("pAbsender") = "User " & n
        qd.Parameters("pEmail") = "user" & n & "@test.de"
        qd.Parameters("pDatum") = Now
        qd.Parameters("pGroesse") = 50000
        qd.Parameters("pErstellt") = Now
        qd.Execute dbFailOnError
        lngBatch = lngBatch + 1
        If lngBatch >= lngBatchGroesse Then
            ws.CommitTrans: DoEvents: ws.BeginTrans
            lngBatch = 0
        End If
    Next n
    If lngBatch > 0 Then ws.CommitTrans
    t2 = GetTickCount()
    qd.Close: Set qd = Nothing

    Debug.Print PadR("QueryDef Parameter", 24) & _
                PadR(CStr(m_lngAnzahlRecords), 10) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, m_lngAnzahlRecords), 12) & _
                "Parametrisiert, SQL-Injection-sicher"

    Debug.Print String(70, "-")
    Debug.Print ""

    Set db = Nothing: Set ws = Nothing
    Exit Sub

ErrHandler:
    On Error Resume Next
    ws.Rollback
    On Error GoTo 0
    Debug.Print "  FEHLER: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST: Batch-Schreiben ueber DIRECT DAO (verschiedene Batch-Groessen)
' ---------------------------------------------------------------------------
Private Sub TestBatchSchreiben_Direct()
    On Error GoTo ErrHandler

    Debug.Print "--- DIRECT DAO WRITES (Recordset.AddNew + Transaction) ---"
    Debug.Print "    (DAO.OpenDatabase direkt auf Netzwerk-DBl kein Linked Table)"
    Debug.Print ""
    Debug.Print PadR("Batch-Groesse", 16) & PadR("Records", 10) & _
                PadR("Zeit", 12) & PadR("Pro Record", 12) & "Faktor"
    Debug.Print String(62, "-")

    m_lngBaselineDirect = 0

    ' Baseline: Ohne Transaktion
    RunBatchWriteTest_Direct 1, False

    ' Mit Transaktion: verschiedene Batch-Groessen
    RunBatchWriteTest_Direct 10, True
    RunBatchWriteTest_Direct 25, True
    RunBatchWriteTest_Direct 50, True
    RunBatchWriteTest_Direct 100, True

    Debug.Print String(62, "-")
    Debug.Print ""

    ' Vergleich Linked vs Direct
    If m_lngBaselineLinked > 0 And m_lngBaselineDirect > 0 Then
        Debug.Print "VERGLEICH (ohne Transaction, Baseline):"
        Debug.Print "  Linked : " & FormatMS(m_lngBaselineLinked) & " (" & _
                    FormatProRec(m_lngBaselineLinked, m_lngAnzahlRecords) & " pro Record)"
        Debug.Print "  Direct : " & FormatMS(m_lngBaselineDirect) & " (" & _
                    FormatProRec(m_lngBaselineDirect, m_lngAnzahlRecords) & " pro Record)"
        If m_lngBaselineDirect > 0 Then
            Debug.Print "  -> Direct ist " & _
                Format(CDbl(m_lngBaselineLinked) / CDbl(m_lngBaselineDirect), "0.0") & _
                "x " & IIf(m_lngBaselineDirect < m_lngBaselineLinked, "schneller", "langsamer")
        End If
        Debug.Print ""
    End If

    Exit Sub

ErrHandler:
    Debug.Print "  FEHLER: " & Err.Description
End Sub


' Einzelner Batch-Write-Test ueber Direct DAO
Private Function RunBatchWriteTest_Direct(ByVal lngBatchGroesse As Long, _
                                           ByVal blnTransaction As Boolean) As Long
    On Error GoTo ErrHandler

    Dim dbDirect As DAO.Database
    Dim ws As DAO.Workspace
    Dim rs As DAO.Recordset
    Dim t1 As Long, t2 As Long
    Dim n As Long, lngBatchCount As Long

    Set ws = DBEngine.Workspaces(0)
    Set dbDirect = ws.OpenDatabase(m_strTestDBPfad)

    ' Tabelle leeren
    dbDirect.Execute "DELETE FROM tblTestEmails", dbFailOnError

    ' Schreiben
    Set rs = dbDirect.OpenRecordset("tblTestEmails", dbOpenDynaset)

    t1 = GetTickCount()
    lngBatchCount = 0

    If blnTransaction Then ws.BeginTrans

    For n = 1 To m_lngAnzahlRecords
        rs.AddNew
        rs!Hash = GeneriereTestHash(n)
        rs!Betreff = "Direct Test " & n & " Batch=" & lngBatchGroesse
        rs!Absender = "Testuser " & (n Mod 50)
        rs!AbsenderEmail = "test" & (n Mod 50) & "@example.com"
        rs!DatumGesendet = DateAdd("h", -n, Now)
        rs!DatumEmpfangen = DateAdd("h", -n + 0.1, Now)
        rs!Groesse = 10000 + (n * 137)
        rs!AnhangAnzahl = n Mod 5
        rs!OrdnerID = 1
        rs!SyncLaufID = 1
        rs!ErstelltAm = Now
        rs.Update

        lngBatchCount = lngBatchCount + 1

        If blnTransaction And lngBatchCount >= lngBatchGroesse Then
            ws.CommitTrans
            DoEvents
            ws.BeginTrans
            lngBatchCount = 0
        End If
    Next n

    If blnTransaction And lngBatchCount > 0 Then ws.CommitTrans

    t2 = GetTickCount()
    rs.Close: Set rs = Nothing
    dbDirect.Close: Set dbDirect = Nothing: Set ws = Nothing

    ' Ergebnis
    Dim lngZeit As Long: lngZeit = t2 - t1
    Dim strBatch As String
    If blnTransaction Then
        strBatch = CStr(lngBatchGroesse) & " (Trans)"
    Else
        strBatch = "1 (kein Trans)"
    End If

    Debug.Print PadR(strBatch, 16) & _
                PadR(CStr(m_lngAnzahlRecords), 10) & _
                PadR(FormatMS(lngZeit), 12) & _
                PadR(FormatProRec(lngZeit, m_lngAnzahlRecords), 12) & _
                IIf(m_lngBaselineDirect > 0 And lngZeit > 0, _
                    Format(CDbl(m_lngBaselineDirect) / CDbl(lngZeit), "0.0") & "x", _
                    "Baseline")

    If Not blnTransaction Then m_lngBaselineDirect = lngZeit

    RunBatchWriteTest_Direct = lngZeit
    Exit Function

ErrHandler:
    On Error Resume Next
    If blnTransaction Then ws.Rollback
    If Not dbDirect Is Nothing Then dbDirect.Close
    On Error GoTo 0
    Debug.Print "  FEHLER bei Batch=" & lngBatchGroesse & ": " & Err.Description
    RunBatchWriteTest_Direct = -1
End Function


' ---------------------------------------------------------------------------
' TEST: Memo-Feld Performance (HTML Body verschiedene Groessen)
' ---------------------------------------------------------------------------
Private Sub TestMemoSchreiben()
    On Error GoTo ErrHandler

    Debug.Print "--- MEMO-FELD PERFORMANCE (tblTestContent, Direct DAO, Trans=25) ---"
    Debug.Print ""
    Debug.Print PadR("HTML-Groesse", 16) & PadR("Records", 10) & _
                PadR("Zeit", 12) & PadR("Pro Record", 12) & "Daten-Vol."
    Debug.Print String(62, "-")

    ' 3 Groessen testen: 1KB, 50KB, 500KB
    RunMemoWriteTest 1
    RunMemoWriteTest 50
    RunMemoWriteTest 500

    Debug.Print String(62, "-")
    Debug.Print ""
    Exit Sub

ErrHandler:
    Debug.Print "  FEHLER: " & Err.Description
End Sub


Private Sub RunMemoWriteTest(ByVal lngKB As Long)
    On Error GoTo ErrHandler

    Dim dbDirect As DAO.Database
    Dim ws As DAO.Workspace
    Dim rs As DAO.Recordset
    Dim t1 As Long, t2 As Long
    Dim n As Long, lngBatch As Long
    Dim strHTML As String

    Set ws = DBEngine.Workspaces(0)
    Set dbDirect = ws.OpenDatabase(m_strTestDBPfad)

    ' Tabelle leeren
    dbDirect.Execute "DELETE FROM tblTestContent", dbFailOnError

    ' HTML-Testinhalt generieren
    strHTML = GeneriereTestHTML(lngKB)

    ' Anzahl Records (weniger bei grossen Memos)
    Dim lngCount As Long
    If lngKB >= 500 Then
        lngCount = 10
    ElseIf lngKB >= 50 Then
        lngCount = 25
    Else
        lngCount = m_lngAnzahlRecords
    End If

    Set rs = dbDirect.OpenRecordset("tblTestContent", dbOpenDynaset)

    t1 = GetTickCount()
    lngBatch = 0
    ws.BeginTrans

    For n = 1 To lngCount
        rs.AddNew
        rs!EmailID = n
        rs!BodyHTML = strHTML
        rs!BodyPlain = Left(strHTML, lngKB * 100)  ' ~10% als Plain
        rs!ErstelltAm = Now
        rs.Update
        lngBatch = lngBatch + 1
        If lngBatch >= 25 Then
            ws.CommitTrans: DoEvents: ws.BeginTrans
            lngBatch = 0
        End If
    Next n

    If lngBatch > 0 Then ws.CommitTrans
    t2 = GetTickCount()

    rs.Close: Set rs = Nothing
    dbDirect.Close: Set dbDirect = Nothing: Set ws = Nothing

    Dim lngZeit As Long: lngZeit = t2 - t1
    Dim dblVolMB As Double
    dblVolMB = (CDbl(lngKB) * CDbl(lngCount)) / 1024

    Debug.Print PadR(lngKB & " KB", 16) & _
                PadR(CStr(lngCount), 10) & _
                PadR(FormatMS(lngZeit), 12) & _
                PadR(FormatProRec(lngZeit, lngCount), 12) & _
                Format(dblVolMB, "0.0") & " MB total"
    Exit Sub

ErrHandler:
    On Error Resume Next
    ws.Rollback
    If Not dbDirect Is Nothing Then dbDirect.Close
    On Error GoTo 0
    Debug.Print "  FEHLER bei " & lngKB & "KB: " & Err.Description
End Sub


' ---------------------------------------------------------------------------
' TEST: Hash-Lookup Performance (Duplikatpruefung)
' ---------------------------------------------------------------------------
Private Sub TestHashLookup()
    On Error GoTo ErrHandler

    Debug.Print "--- HASH-LOOKUP PERFORMANCE (Duplikatpruefung) ---"
    Debug.Print ""

    Dim dbDirect As DAO.Database
    Dim ws As DAO.Workspace
    Dim rs As DAO.Recordset
    Dim t1 As Long, t2 As Long
    Dim n As Long

    Set ws = DBEngine.Workspaces(0)
    Set dbDirect = ws.OpenDatabase(m_strTestDBPfad)

    ' --- Vorbereitung: 1000 Records einfuegen ---
    Debug.Print "Fuege 1000 Test-Records fuer Lookup ein..."
    dbDirect.Execute "DELETE FROM tblTestEmails", dbFailOnError

    Set rs = dbDirect.OpenRecordset("tblTestEmails", dbOpenDynaset)
    ws.BeginTrans
    For n = 1 To 1000
        rs.AddNew
        rs!Hash = GeneriereTestHash(n)
        rs!Betreff = "Lookup Test " & n
        rs!Absender = "user"
        rs!AbsenderEmail = "user@test.de"
        rs!DatumGesendet = Now
        rs!Groesse = 10000
        rs!ErstelltAm = Now
        rs.Update
        If n Mod 100 = 0 Then
            ws.CommitTrans: ws.BeginTrans
        End If
    Next n
    ws.CommitTrans
    rs.Close: Set rs = Nothing
    Debug.Print "  -> 1000 Records eingefuegt."
    Debug.Print ""

    Debug.Print PadR("Methode", 30) & PadR("Batch", 8) & _
                PadR("Zeit", 12) & PadR("Pro Lookup", 12) & "Treffer"
    Debug.Print String(74, "-")

    ' --- Test A: DCount einzeln (25 Lookups) ---
    Dim lngFound As Long
    t1 = GetTickCount()
    lngFound = 0
    For n = 1 To 25
        ' 13 existierende + 12 nicht-existierende Hashes
        Dim strH As String
        If n Mod 2 = 0 Then
            strH = GeneriereTestHash(n * 3)        ' existiert (n*3 <= 1000)
        Else
            strH = GeneriereTestHash(n + 5000)     ' existiert NICHT
        End If
        If DCount("*", "tblTestEmails", "Hash='" & strH & "'") > 0 Then
            lngFound = lngFound + 1
        End If
    Next n
    t2 = GetTickCount()

    Debug.Print PadR("DCount (einzeln)", 30) & PadR("1", 8) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, 25), 12) & _
                lngFound & "/25"

    ' HINWEIS: DCount/DLookup geht ueber Linked Table!
    ' Fuer Direct DAO testen wir SQL-Queries

    ' --- Test B: SELECT WHERE Hash = '...' einzeln (Direct DAO) ---
    t1 = GetTickCount()
    lngFound = 0
    For n = 1 To 25
        If n Mod 2 = 0 Then
            strH = GeneriereTestHash(n * 3)
        Else
            strH = GeneriereTestHash(n + 5000)
        End If
        Set rs = dbDirect.OpenRecordset( _
            "SELECT Hash FROM tblTestEmails WHERE Hash='" & strH & "'", dbOpenSnapshot)
        If Not rs.EOF Then lngFound = lngFound + 1
        rs.Close: Set rs = Nothing
    Next n
    t2 = GetTickCount()

    Debug.Print PadR("SELECT einzeln (Direct)", 30) & PadR("1", 8) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, 25), 12) & _
                lngFound & "/25"

    ' --- Test C: SELECT WHERE Hash IN (...) Batch=25 (Direct DAO) ---
    ' Genau wie modAsyncBuffer.PruefeHashes() es macht
    Dim strIN As String
    strIN = ""
    For n = 1 To 25
        If n Mod 2 = 0 Then
            strH = GeneriereTestHash(n * 3)
        Else
            strH = GeneriereTestHash(n + 5000)
        End If
        If strIN <> "" Then strIN = strIN & ","
        strIN = strIN & "'" & strH & "'"
    Next n

    t1 = GetTickCount()
    lngFound = 0
    Set rs = dbDirect.OpenRecordset( _
        "SELECT Hash FROM tblTestEmails WHERE Hash IN (" & strIN & ")", dbOpenSnapshot)
    Do While Not rs.EOF
        lngFound = lngFound + 1
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    t2 = GetTickCount()

    Debug.Print PadR("SELECT IN() Batch (Direct)", 30) & PadR("25", 8) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, 25), 12) & _
                lngFound & "/25"

    ' --- Test D: SELECT IN() Batch=50 ---
    strIN = ""
    For n = 1 To 50
        If n Mod 2 = 0 Then
            strH = GeneriereTestHash(n * 3)
        Else
            strH = GeneriereTestHash(n + 5000)
        End If
        If strIN <> "" Then strIN = strIN & ","
        strIN = strIN & "'" & strH & "'"
    Next n

    t1 = GetTickCount()
    lngFound = 0
    Set rs = dbDirect.OpenRecordset( _
        "SELECT Hash FROM tblTestEmails WHERE Hash IN (" & strIN & ")", dbOpenSnapshot)
    Do While Not rs.EOF
        lngFound = lngFound + 1
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    t2 = GetTickCount()

    Debug.Print PadR("SELECT IN() Batch (Direct)", 30) & PadR("50", 8) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, 50), 12) & _
                lngFound & "/50"

    ' --- Test E: SELECT IN() Batch=100 ---
    strIN = ""
    For n = 1 To 100
        If n Mod 2 = 0 Then
            strH = GeneriereTestHash(n * 3)
        Else
            strH = GeneriereTestHash(n + 5000)
        End If
        If strIN <> "" Then strIN = strIN & ","
        strIN = strIN & "'" & strH & "'"
    Next n

    t1 = GetTickCount()
    lngFound = 0
    Set rs = dbDirect.OpenRecordset( _
        "SELECT Hash FROM tblTestEmails WHERE Hash IN (" & strIN & ")", dbOpenSnapshot)
    Do While Not rs.EOF
        lngFound = lngFound + 1
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    t2 = GetTickCount()

    Debug.Print PadR("SELECT IN() Batch (Direct)", 30) & PadR("100", 8) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, 100), 12) & _
                lngFound & "/100"

    Debug.Print String(74, "-")
    Debug.Print ""
    Debug.Print "HINWEIS: DCount geht ueber Linked Table (ACE-Overhead)."
    Debug.Print "         SELECT-Tests gehen ueber Direct DAO (kein Linked-Table-Overhead)."
    Debug.Print "         IN()-Batch reduziert Netzwerk-Roundtrips drastisch."
    Debug.Print ""

    dbDirect.Close: Set dbDirect = Nothing: Set ws = Nothing
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not dbDirect Is Nothing Then dbDirect.Close
    On Error GoTo 0
    Debug.Print "  FEHLER: " & Err.Description
End Sub


' ===========================================================================
' BEREICH 3: OUTLOOK-EXTRAKTION
' ===========================================================================

' Testet verschiedene Methoden zur Mail-Daten-Extraktion.
' Vergleicht: OOM Basic, OOM + Body, Redemption SafeMailItem, Redemption RDO,
'             und das vollstaendige Extract-Release-Pattern.
Public Sub TestOutlookExtraktion(Optional ByVal lngAnzahl As Long = 50)
    On Error GoTo ErrHandler

    ' QPC Frequenz (falls Einzelaufruf)
    QueryPerformanceFrequency m_curFrequency

    Debug.Print String(80, "=")
    Debug.Print "  OUTLOOK-EXTRAKTION PERFORMANCE"
    Debug.Print String(80, "=")
    Debug.Print ""

    ' --- Outlook verbinden ---
    Dim objApp As Object
    On Error Resume Next
    Set objApp = GetObject(, "Outlook.Application")
    If objApp Is Nothing Then Set objApp = CreateObject("Outlook.Application")
    On Error GoTo ErrHandler

    If objApp Is Nothing Then
        Debug.Print "*** FEHLER: Outlook ist nicht verfuegbar!"
        Debug.Print "    Outlook muss laufen fuer diesen Test."
        Exit Sub
    End If

    Dim objNS As Object
    Set objNS = objApp.GetNamespace("MAPI")

    ' Inbox holen
    Dim objFolder As Object
    Set objFolder = objNS.GetDefaultFolder(6) ' olFolderInbox = 6

    Debug.Print "  Outlook     : " & objApp.Version
    Debug.Print "  Ordner      : " & objFolder.FolderPath
    Debug.Print "  Elemente    : " & objFolder.Items.Count
    Debug.Print "  Test-Anzahl : " & lngAnzahl & " Mails"
    Debug.Print ""

    If objFolder.Items.Count = 0 Then
        Debug.Print "*** Ordner ist leer! Bitte Ordner mit Mails waehlen."
        Exit Sub
    End If

    If lngAnzahl > objFolder.Items.Count Then lngAnzahl = objFolder.Items.Count

    ' --- Nur MailItems zaehlen ---
    Debug.Print "Suche MailItems..."
    Dim colItems As Object
    Set colItems = objFolder.Items

    ' Sammle Indizes von MailItems
    Dim aMailIdx() As Long
    ReDim aMailIdx(0 To lngAnzahl - 1)
    Dim lngGefunden As Long: lngGefunden = 0
    Dim i As Long

    For i = 1 To colItems.Count
        If lngGefunden >= lngAnzahl Then Exit For
        On Error Resume Next
        Dim strTypName As String
        strTypName = TypeName(colItems(i))
        On Error GoTo ErrHandler
        If strTypName = "MailItem" Then
            aMailIdx(lngGefunden) = i
            lngGefunden = lngGefunden + 1
        End If
    Next i

    If lngGefunden = 0 Then
        Debug.Print "*** Keine MailItems im Ordner gefunden!"
        Exit Sub
    End If
    If lngGefunden < lngAnzahl Then
        lngAnzahl = lngGefunden
        ReDim Preserve aMailIdx(0 To lngAnzahl - 1)
    End If
    Debug.Print "  -> " & lngAnzahl & " MailItems gefunden."
    Debug.Print ""

    Debug.Print PadR("Methode", 32) & PadR("Mails", 8) & _
                PadR("Gesamt", 12) & PadR("Pro Mail", 12) & "Details"
    Debug.Print String(76, "-")

    ' =====================================================================
    ' TEST 1: OOM Basic (Subject, Date, Size - KEIN Security Issue)
    ' =====================================================================
    Dim t1 As Long, t2 As Long
    Dim objItem As Object
    Dim strDummy As String
    Dim dtDummy As Date
    Dim lngDummy As Long
    Dim intDummy As Integer
    Dim blnDummy As Boolean

    t1 = GetTickCount()
    For i = 0 To lngAnzahl - 1
        Set objItem = colItems(aMailIdx(i))
        On Error Resume Next
        strDummy = objItem.Subject
        strDummy = objItem.SenderName
        dtDummy = objItem.SentOn
        dtDummy = objItem.ReceivedTime
        lngDummy = objItem.Size
        intDummy = objItem.Importance
        blnDummy = objItem.UnRead
        strDummy = objItem.MessageClass
        strDummy = objItem.ConversationTopic
        strDummy = objItem.EntryID
        intDummy = objItem.Attachments.Count
        On Error GoTo ErrHandler
        Set objItem = Nothing
    Next i
    t2 = GetTickCount()

    Debug.Print PadR("OOM Basic (kein Security)", 32) & _
                PadR(CStr(lngAnzahl), 8) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, lngAnzahl), 12) & _
                "Subject,Date,Size,EntryID..."

    ' =====================================================================
    ' TEST 2: OOM + Body/HTMLBody (grosse Properties)
    ' =====================================================================
    t1 = GetTickCount()
    For i = 0 To lngAnzahl - 1
        Set objItem = colItems(aMailIdx(i))
        On Error Resume Next
        strDummy = objItem.Subject
        strDummy = objItem.SenderName
        dtDummy = objItem.SentOn
        dtDummy = objItem.ReceivedTime
        lngDummy = objItem.Size
        strDummy = objItem.Body
        strDummy = objItem.HTMLBody
        strDummy = objItem.ConversationTopic
        On Error GoTo ErrHandler
        Set objItem = Nothing
    Next i
    t2 = GetTickCount()

    Debug.Print PadR("OOM + Body/HTMLBody", 32) & _
                PadR(CStr(lngAnzahl), 8) & _
                PadR(FormatMS(t2 - t1), 12) & _
                PadR(FormatProRec(t2 - t1, lngAnzahl), 12) & _
                "Inkl. HTML+PlainText Laden"

    ' =====================================================================
    ' TEST 3: OOM + SenderEmailAddress (ACHTUNG: Security Guard!)
    ' =====================================================================
    Debug.Print PadR("OOM + Email (SECURITY!)", 32) & _
                PadR("-", 8) & PadR("-", 12) & PadR("-", 12) & _
                "UEBERSPRUNGEN (OOM Security Guard)"

    ' =====================================================================
    ' TEST 4: Redemption SafeMailItem (alle Felder)
    ' =====================================================================
    Dim blnRedemptionOK As Boolean: blnRedemptionOK = False
    Dim objSafe As Object

    On Error Resume Next
    Set objSafe = CreateObject("Redemption.SafeMailItem")
    blnRedemptionOK = (Err.Number = 0 And Not objSafe Is Nothing)
    Err.Clear
    On Error GoTo ErrHandler

    If blnRedemptionOK Then
        t1 = GetTickCount()
        For i = 0 To lngAnzahl - 1
            Set objItem = colItems(aMailIdx(i))
            objSafe.Item = objItem

            On Error Resume Next
            strDummy = objSafe.Subject
            strDummy = objSafe.SenderName
            strDummy = objSafe.SenderEmailAddress  ' KEIN Security Prompt!
            dtDummy = objSafe.SentOn
            dtDummy = objSafe.ReceivedTime
            lngDummy = objSafe.Size
            intDummy = objSafe.Importance
            blnDummy = objSafe.UnRead
            strDummy = objSafe.Body
            strDummy = objSafe.HTMLBody
            strDummy = objSafe.ConversationTopic
            strDummy = objSafe.EntryID
            intDummy = objSafe.Attachments.Count
            On Error GoTo ErrHandler
            Set objItem = Nothing
        Next i
        t2 = GetTickCount()

        Debug.Print PadR("Redemption SafeMailItem", 32) & _
                    PadR(CStr(lngAnzahl), 8) & _
                    PadR(FormatMS(t2 - t1), 12) & _
                    PadR(FormatProRec(t2 - t1, lngAnzahl), 12) & _
                    "Alle Felder, KEIN Security Prompt"
    Else
        Debug.Print PadR("Redemption SafeMailItem", 32) & _
                    PadR("-", 8) & PadR("-", 12) & PadR("-", 12) & _
                    "NICHT VERFUEGBAR (Redemption nicht installiert)"
    End If

    ' =====================================================================
    ' TEST 5: Redemption RDO (direkter MAPI-Zugriff)
    ' =====================================================================
    Dim blnRDOOK As Boolean: blnRDOOK = False
    Dim objRDOSession As Object
    Dim objRDOFolder As Object

    On Error Resume Next
    Set objRDOSession = CreateObject("Redemption.RDOSession")
    If Err.Number = 0 And Not objRDOSession Is Nothing Then
        ' MAPI-Session von Outlook uebernehmen
        objRDOSession.MAPIOBJECT = objNS.MAPIOBJECT
        If Err.Number = 0 Then
            ' Ordner per EntryID/StoreID holen
            Set objRDOFolder = objRDOSession.GetFolderFromID( _
                objFolder.EntryID, objFolder.StoreID)
            blnRDOOK = (Err.Number = 0 And Not objRDOFolder Is Nothing)
        End If
    End If
    Err.Clear
    On Error GoTo ErrHandler

    If blnRDOOK Then
        Dim objRDOMail As Object

        t1 = GetTickCount()
        For i = 0 To lngAnzahl - 1
            ' Hole RDO-Mail ueber OOM-EntryID
            Set objItem = colItems(aMailIdx(i))
            On Error Resume Next
            Set objRDOMail = objRDOSession.GetMessageFromID( _
                objItem.EntryID, objFolder.StoreID)
            On Error GoTo ErrHandler

            If Not objRDOMail Is Nothing Then
                On Error Resume Next
                strDummy = objRDOMail.Subject
                strDummy = objRDOMail.SenderName
                strDummy = objRDOMail.SenderEmailAddress
                dtDummy = objRDOMail.SentOn
                dtDummy = objRDOMail.ReceivedTime
                lngDummy = objRDOMail.Size
                intDummy = objRDOMail.Importance
                blnDummy = objRDOMail.UnRead
                strDummy = objRDOMail.Body
                strDummy = objRDOMail.HTMLBody
                strDummy = objRDOMail.EntryID
                intDummy = objRDOMail.Attachments.Count
                ' MAPI Properties (wie im echten ExtrahiereMailDaten)
                strDummy = objRDOMail.Fields(&H1035001E)  ' PR_INTERNET_MESSAGE_ID
                strDummy = objRDOMail.Fields(&HE04001E)   ' PR_DISPLAY_TO
                strDummy = objRDOMail.Fields(&H7D001E)    ' PR_TRANSPORT_MESSAGE_HEADERS
                On Error GoTo ErrHandler
                Set objRDOMail = Nothing
            End If
            Set objItem = Nothing
        Next i
        t2 = GetTickCount()

        Debug.Print PadR("Redemption RDO (MAPI)", 32) & _
                    PadR(CStr(lngAnzahl), 8) & _
                    PadR(FormatMS(t2 - t1), 12) & _
                    PadR(FormatProRec(t2 - t1, lngAnzahl), 12) & _
                    "Alle Felder + MAPI Properties"
    Else
        Debug.Print PadR("Redemption RDO (MAPI)", 32) & _
                    PadR("-", 8) & PadR("-", 12) & PadR("-", 12) & _
                    "NICHT VERFUEGBAR"
    End If

    ' =====================================================================
    ' TEST 6: Recipients (Empfaenger-Auflistung) via Redemption
    ' =====================================================================
    If blnRedemptionOK Then
        t1 = GetTickCount()
        Dim lngTotalRecp As Long: lngTotalRecp = 0
        For i = 0 To lngAnzahl - 1
            Set objItem = colItems(aMailIdx(i))
            objSafe.Item = objItem
            On Error Resume Next
            Dim objRecipients As Object
            Set objRecipients = objSafe.Recipients
            If Not objRecipients Is Nothing Then
                Dim r As Long
                For r = 1 To objRecipients.Count
                    strDummy = objRecipients(r).Name
                    strDummy = objRecipients(r).Address
                    intDummy = objRecipients(r).Type
                    lngTotalRecp = lngTotalRecp + 1
                Next r
            End If
            On Error GoTo ErrHandler
            Set objItem = Nothing
        Next i
        t2 = GetTickCount()

        Debug.Print PadR("Redemption Recipients", 32) & _
                    PadR(CStr(lngAnzahl), 8) & _
                    PadR(FormatMS(t2 - t1), 12) & _
                    PadR(FormatProRec(t2 - t1, lngAnzahl), 12) & _
                    lngTotalRecp & " Empfaenger total"
    End If

    ' =====================================================================
    ' TEST 7: Extract-Release Pattern (COM-Haltezeit messen)
    ' =====================================================================
    If blnRDOOK Then
        Dim lngCOMZeitTotal As Long: lngCOMZeitTotal = 0
        Dim lngHashZeitTotal As Long: lngHashZeitTotal = 0
        Dim lngReleaseCount As Long: lngReleaseCount = 0

        t1 = GetTickCount()
        For i = 0 To lngAnzahl - 1
            Set objItem = colItems(aMailIdx(i))
            On Error Resume Next
            Set objRDOMail = objRDOSession.GetMessageFromID( _
                objItem.EntryID, objFolder.StoreID)
            On Error GoTo ErrHandler

            If Not objRDOMail Is Nothing Then
                ' --- PHASE A: COM aktiv - Daten extrahieren ---
                Dim tA1 As Long: tA1 = GetTickCount()

                Dim strBetreff As String, strAbsender As String
                Dim strAbsEmail As String, strBody As String
                Dim strHTML As String, strMsgID As String
                Dim dtEmpfangen As Date, dtGesendet As Date
                Dim lngGroesse As Long, intAnz As Integer

                On Error Resume Next
                strBetreff = objRDOMail.Subject
                strAbsender = objRDOMail.SenderName
                strAbsEmail = objRDOMail.SenderEmailAddress
                dtGesendet = objRDOMail.SentOn
                dtEmpfangen = objRDOMail.ReceivedTime
                lngGroesse = objRDOMail.Size
                strBody = objRDOMail.Body
                strHTML = objRDOMail.HTMLBody
                strMsgID = objRDOMail.Fields(&H1035001E)
                intAnz = objRDOMail.Attachments.Count
                On Error GoTo ErrHandler

                ' --- COM freigeben ---
                Set objRDOMail = Nothing
                Set objItem = Nothing

                Dim tA2 As Long: tA2 = GetTickCount()
                lngCOMZeitTotal = lngCOMZeitTotal + (tA2 - tA1)

                ' --- PHASE B: Ohne COM - Hash berechnen ---
                Dim tB1 As Long: tB1 = GetTickCount()
                Dim strHash As String
                strHash = TestSHA256(strBetreff & "|" & strAbsEmail & "|" & _
                                     Format(dtEmpfangen, "yyyymmddhhnnss"))
                Dim tB2 As Long: tB2 = GetTickCount()
                lngHashZeitTotal = lngHashZeitTotal + (tB2 - tB1)

                lngReleaseCount = lngReleaseCount + 1
            Else
                Set objItem = Nothing
            End If
        Next i
        t2 = GetTickCount()

        Debug.Print String(76, "-")
        Debug.Print ""
        Debug.Print "EXTRACT-RELEASE PATTERN (COM-Haltezeit):"
        Debug.Print "  Getestete Mails    : " & lngReleaseCount
        Debug.Print "  Gesamt-Zeit        : " & FormatMS(t2 - t1)
        Debug.Print "  COM-Haltezeit ges. : " & FormatMS(lngCOMZeitTotal) & _
                    " (" & FormatProRec(lngCOMZeitTotal, lngReleaseCount) & " pro Mail)"
        Debug.Print "  Hash-Zeit ges.     : " & FormatMS(lngHashZeitTotal) & _
                    " (" & FormatProRec(lngHashZeitTotal, lngReleaseCount) & " pro Mail)"
        If (t2 - t1) > 0 Then
            Debug.Print "  COM-Anteil         : " & _
                Format(CDbl(lngCOMZeitTotal) / CDbl(t2 - t1) * 100, "0.0") & "%"
        End If
        Debug.Print ""
        Debug.Print "  BEWERTUNG:"
        If lngReleaseCount > 0 Then
            Dim dblProMail As Double
            dblProMail = CDbl(lngCOMZeitTotal) / CDbl(lngReleaseCount)
            If dblProMail < 10 Then
                Debug.Print "  -> HERVORRAGEND: COM unter 10ms pro Mail."
                Debug.Print "     Extract-Release entlastet Outlook effektiv."
            ElseIf dblProMail < 50 Then
                Debug.Print "  -> GUT: COM unter 50ms pro Mail."
                Debug.Print "     Extract-Release sinnvoll."
            ElseIf dblProMail < 200 Then
                Debug.Print "  -> MAESSIG: COM 50-200ms pro Mail."
                Debug.Print "     Body/HTML Laden dominiert. Evtl. lazy loading erwaegen."
            Else
                Debug.Print "  -> LANGSAM: COM ueber 200ms pro Mail."
                Debug.Print "     Exchange-Latenz? Netzwerk-Probleme?"
            End If
        End If
        Debug.Print "  HINWEIS: Ohne Anhang-/MSG-Speicherung gemessen!"
        Debug.Print "           Anhang-SaveAsFile wuerde COM-Zeit erhoehen."
    End If

    Debug.Print ""
    Debug.Print String(76, "-")
    Debug.Print ""

    ' --- Aufraeumen ---
    On Error Resume Next
    Set objSafe = Nothing
    Set objRDOFolder = Nothing
    If Not objRDOSession Is Nothing Then
        objRDOSession.Logoff
        Set objRDOSession = Nothing
    End If
    Set objFolder = Nothing
    Set objNS = Nothing
    ' objApp NICHT schliessen (Outlook laeuft weiter)
    Set objApp = Nothing
    On Error GoTo 0

    Debug.Print "Outlook-Tests abgeschlossen."
    Debug.Print String(80, "=")
    Exit Sub

ErrHandler:
    Debug.Print "*** FEHLER in TestOutlookExtraktion: " & Err.Description
    On Error Resume Next
    Set objSafe = Nothing
    Set objRDOFolder = Nothing
    If Not objRDOSession Is Nothing Then objRDOSession.Logoff
    Set objRDOSession = Nothing
    Set objRDOMail = Nothing
    Set objItem = Nothing
    Set objFolder = Nothing
    Set objNS = Nothing
    Set objApp = Nothing
    On Error GoTo 0
End Sub


' ===========================================================================
' BEREICH 4: SHA256-HASHING PERFORMANCE
' ===========================================================================

' Testet SHA256-Durchsatz mit verschiedenen Eingabegroessen
Public Sub TestHashingSpeed(Optional ByVal lngIterationen As Long = 1000)
    On Error GoTo ErrHandler

    ' QPC Frequenz (falls Einzelaufruf)
    QueryPerformanceFrequency m_curFrequency

    Debug.Print String(80, "=")
    Debug.Print "  SHA256-HASHING PERFORMANCE (CryptoAPI)"
    Debug.Print String(80, "=")
    Debug.Print ""
    Debug.Print PadR("Eingabe", 14) & PadR("Anzahl", 10) & _
                PadR("Gesamt", 12) & PadR("Pro Hash", 12) & "Durchsatz"
    Debug.Print String(60, "-")

    ' Test A: 100 Bytes (typischer Mail-Hash-Input: Betreff|Email|Datum)
    RunHashTest 100, lngIterationen

    ' Test B: 1 KB (laengerer Input)
    RunHashTest 1024, lngIterationen

    ' Test C: 50 KB (Body-Hash Szenario)
    RunHashTest 50 * 1024, CLng(lngIterationen / 10)

    ' Test D: 500 KB (grosser HTML-Body)
    RunHashTest 500 * 1024, CLng(lngIterationen / 100)

    Debug.Print String(60, "-")
    Debug.Print ""

    ' Praxisbewertung
    Debug.Print "PRAXIS-BEWERTUNG:"
    Debug.Print "  Mail-Hash (100 Bytes): Der typische UseCase fuer Duplikatpruefung."
    Debug.Print "  Bei 10.000 Mails mit 100-Byte-Hash circa " & _
                Format(CDbl(lngIterationen) / CDbl(IIf(m_lngLetzteHashZeit > 0, _
                m_lngLetzteHashZeit, 1)) * 10, "#,##0") & " Hashes pro Sekunde."
    Debug.Print "  -> SHA256 ist kein Flaschenhals fuer die Sync-Performance."
    Debug.Print ""
    Debug.Print String(80, "=")
    Exit Sub

ErrHandler:
    Debug.Print "*** FEHLER: " & Err.Description
End Sub

Private Sub RunHashTest(ByVal lngBytes As Long, ByVal lngAnzahl As Long)
    On Error GoTo ErrHandler

    If lngAnzahl < 1 Then lngAnzahl = 1

    ' Teststring generieren
    Dim strInput As String
    strInput = String(lngBytes, "A")
    ' Etwas Variation einbauen
    If lngBytes > 10 Then Mid(strInput, 5, 5) = "12345"

    Dim t1 As Long, t2 As Long
    Dim n As Long
    Dim strResult As String

    t1 = GetTickCount()
    For n = 1 To lngAnzahl
        strResult = TestSHA256(strInput)
        If n Mod 100 = 0 Then DoEvents
    Next n
    t2 = GetTickCount()

    Dim lngZeit As Long: lngZeit = t2 - t1
    If lngBytes <= 100 Then m_lngLetzteHashZeit = lngZeit

    ' Durchsatz berechnen
    Dim strDurchsatz As String
    If lngZeit > 0 Then
        Dim dblPS As Double
        dblPS = CDbl(lngAnzahl) / (CDbl(lngZeit) / 1000)
        strDurchsatz = Format(dblPS, "#,##0") & " /Sek"
    Else
        strDurchsatz = ">>1000 /Sek"
    End If

    ' Eingabegroesse formatieren
    Dim strGroesse As String
    If lngBytes >= 1024 Then
        strGroesse = Format(lngBytes / 1024, "0") & " KB"
    Else
        strGroesse = lngBytes & " Bytes"
    End If

    Debug.Print PadR(strGroesse, 14) & _
                PadR(CStr(lngAnzahl), 10) & _
                PadR(FormatMS(lngZeit), 12) & _
                PadR(FormatProRec(lngZeit, lngAnzahl), 12) & _
                strDurchsatz
    Exit Sub

ErrHandler:
    Debug.Print "  FEHLER bei " & lngBytes & " Bytes: " & Err.Description
End Sub


' ===========================================================================
' SHA256 - Eigenstaendige Implementierung (fuer Benchmark)
' ===========================================================================

' SHA256-Hash aus String (ANSI) berechnen.
' Gibt 64 Zeichen Hex-String zurueck, oder "" bei Fehler.
Private Function TestSHA256(ByVal strInput As String) As String
    On Error GoTo ErrHandler

    #If VBA7 Then
        Dim hProv As LongPtr, hHash As LongPtr
    #Else
        Dim hProv As Long, hHash As Long
    #End If

    Dim abData()    As Byte
    Dim abHash(0 To 31) As Byte
    Dim lngHashLen  As Long
    Dim j           As Long
    Dim strResult   As String

    If Len(strInput) = 0 Then TestSHA256 = "": Exit Function

    ' String -> Byte-Array (ANSI)
    abData = StrConv(strInput, vbFromUnicode)

    ' Provider oeffnen
    If CryptAcquireContext(hProv, vbNullString, vbNullString, _
                           PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then
        TestSHA256 = "": Exit Function
    End If

    ' Hash-Objekt erstellen
    If CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash) = 0 Then
        CryptReleaseContext hProv, 0
        TestSHA256 = "": Exit Function
    End If

    ' Daten hashen
    If CryptHashData(hHash, abData(0), UBound(abData) + 1, 0) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        TestSHA256 = "": Exit Function
    End If

    ' Hash-Wert lesen
    lngHashLen = 32
    CryptGetHashParam hHash, HP_HASHVAL, abHash(0), lngHashLen, 0

    ' In Hex-String
    strResult = ""
    For j = 0 To 31
        strResult = strResult & Right("0" & Hex(abHash(j)), 2)
    Next j

    ' Aufraeumen
    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0

    TestSHA256 = LCase(strResult)
    Exit Function

ErrHandler:
    On Error Resume Next
    If hHash <> 0 Then CryptDestroyHash hHash
    If hProv <> 0 Then CryptReleaseContext hProv, 0
    TestSHA256 = ""
End Function


' ===========================================================================
' TEST-BACKEND LOESCHEN
' ===========================================================================
Private Sub LoescheTestBackend()
    On Error Resume Next

    If m_strTestDBPfad <> "" Then
        If Dir(m_strTestDBPfad) <> "" Then
            Kill m_strTestDBPfad
        End If
        ' Auch .laccdb Lockfile loeschen
        Dim strLock As String
        strLock = Replace(m_strTestDBPfad, ".accdb", ".laccdb")
        If Dir(strLock) <> "" Then
            Sleep 500  ' Kurz warten bis Lock freigegeben
            Kill strLock
        End If
    End If

    On Error GoTo 0
    Debug.Print "  -> Test-DB geloescht: " & m_strTestDBPfad
End Sub


' ===========================================================================
' HILFSFUNKTIONEN
' ===========================================================================

' String rechts mit Leerzeichen auffuellen
Private Function PadR(ByVal strText As String, ByVal lngLen As Long) As String
    If Len(strText) >= lngLen Then
        PadR = Left(strText, lngLen)
    Else
        PadR = strText & Space(lngLen - Len(strText))
    End If
End Function

' Millisekunden formatieren (z.B. "3.5s" oder "78ms")
Private Function FormatMS(ByVal lngMS As Long) As String
    If lngMS < 0 Then
        FormatMS = "FEHLER"
    ElseIf lngMS >= 10000 Then
        FormatMS = Format(CDbl(lngMS) / 1000, "0.0") & "s"
    ElseIf lngMS >= 1000 Then
        FormatMS = Format(CDbl(lngMS) / 1000, "0.00") & "s"
    Else
        FormatMS = lngMS & "ms"
    End If
End Function

' Pro-Record-Zeit formatieren
Private Function FormatProRec(ByVal lngTotalMS As Long, ByVal lngRecords As Long) As String
    If lngRecords <= 0 Or lngTotalMS < 0 Then
        FormatProRec = "-"
        Exit Function
    End If
    Dim dblMS As Double
    dblMS = CDbl(lngTotalMS) / CDbl(lngRecords)
    If dblMS >= 100 Then
        FormatProRec = Format(dblMS, "0") & " ms"
    ElseIf dblMS >= 1 Then
        FormatProRec = Format(dblMS, "0.0") & " ms"
    ElseIf dblMS > 0 Then
        FormatProRec = Format(dblMS, "0.00") & " ms"
    Else
        FormatProRec = "<0.01 ms"
    End If
End Function

' Test-Hash generieren (deterministisch, 64 Zeichen Hex)
Private Function GeneriereTestHash(ByVal lngSeed As Long) As String
    ' Schneller Pseudo-Hash fuer Testdaten (kein echtes SHA256 noetig)
    Dim strBase As String
    strBase = Right("0000000000000000" & Hex(lngSeed * 2654435761#), 16)
    GeneriereTestHash = strBase & strBase & strBase & strBase
End Function

' Test-HTML generieren (konfigurierbare Groesse in KB)
Private Function GeneriereTestHTML(ByVal lngKB As Long) As String
    Dim strHTML As String
    Dim lngZielBytes As Long
    lngZielBytes = lngKB * 1024

    strHTML = "<html><head><title>Performance-Test</title></head><body>" & vbCrLf

    ' Realistischen HTML-Content erzeugen
    Dim strAbsatz As String
    strAbsatz = "<p>Dies ist ein Testabsatz fuer die Performance-Messung. " & _
                "Er simuliert typischen E-Mail-HTML-Content mit verschiedenen " & _
                "Formatierungen. <b>Fettdruck</b>, <i>Kursiv</i>, " & _
                "<a href=""https://example.com"">Links</a> und " & _
                "<span style=""color: red;"">farbiger Text</span>. " & _
                "Lorem ipsum dolor sit amet, " & _
                "consectetur adipiscing elit, sed do eiusmod tempor incididunt " & _
                "ut labore et dolore magna aliqua.</p>" & vbCrLf

    ' Tabelle (realistisch fuer Business-Mails)
    Dim strTabelle As String
    strTabelle = "<table border=""1"" cellpadding=""5"">" & vbCrLf & _
                 "<tr><th>Datum</th><th>Beschreibung</th><th>Betrag</th></tr>" & vbCrLf
    Dim t As Long
    For t = 1 To 10
        strTabelle = strTabelle & "<tr><td>" & Format(DateAdd("d", -t, Now), "dd.mm.yyyy") & _
                     "</td><td>Position " & t & " Leistungsbeschreibung</td>" & _
                     "<td>" & Format(t * 1234.56, "#,##0.00") & " EUR</td></tr>" & vbCrLf
    Next t
    strTabelle = strTabelle & "</table>" & vbCrLf

    strHTML = strHTML & strTabelle

    ' Absaetze wiederholen bis Zielgroesse erreicht
    Do While Len(strHTML) < lngZielBytes
        strHTML = strHTML & strAbsatz
        ' Alle 10 Absaetze eine Tabelle einfuegen
        If Len(strHTML) Mod (Len(strAbsatz) * 10) < Len(strAbsatz) Then
            strHTML = strHTML & strTabelle
        End If
    Loop

    strHTML = Left(strHTML, lngZielBytes - 20) & vbCrLf & "</body></html>"
    GeneriereTestHTML = strHTML
End Function

' DCount-Hilfsfunktion (Linked Table via CurrentDb, nicht Direct DAO)
' Wird fuer den DCount-Hash-Lookup-Test verwendet
Private Function DCountLinked(ByVal strExpr As String, _
                               ByVal strDomain As String, _
                               Optional ByVal strCriteria As String = "") As Long
    On Error Resume Next
    If strCriteria <> "" Then
        DCountLinked = DCount(strExpr, strDomain, strCriteria)
    Else
        DCountLinked = DCount(strExpr, strDomain)
    End If
    If Err.Number <> 0 Then DCountLinked = 0
    Err.Clear
    On Error GoTo 0
End Function

' Liest Config-Wert aus tblConfig (Standalone-Version)
Private Function LeseConfigStandalone(ByVal strKey As String, _
                                       ByVal strDefault As String) As String
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Wert FROM tblConfig WHERE Schluessel='" & strKey & "'", _
        dbOpenSnapshot)
    If Not rs.EOF Then
        LeseConfigStandalone = Nz(rs!Wert, strDefault)
    Else
        LeseConfigStandalone = strDefault
    End If
    rs.Close: Set rs = Nothing

    If Err.Number <> 0 Then LeseConfigStandalone = strDefault
    Err.Clear
    On Error GoTo 0
End Function
