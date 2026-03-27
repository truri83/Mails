Option Compare Database
Option Explicit

' ===========================================================================
' modSchema - Tabellenschema-Verwaltung fuer OutlookSync
' ===========================================================================
' Erstellt alle 14 Tabellen + Indizes + Standardkonfiguration.
' Bestehende Tabellen werden NICHT ueberschrieben.
' Nutzt TBL_*/CFG_* Konstanten aus modGlobals (v0.5).
'
' WICHTIG - ACCESS DDL RESTRIKTIONEN:
'   Access SQL (Jet/ACE DDL) unterstuetzt KEIN "DEFAULT" in CREATE TABLE!
'   Error 3290 = "Syntax error in CREATE TABLE statement" bei DEFAULT.
'   Defaults muessen nachtraeglich per DAO Field.DefaultValue gesetzt werden.
'   -> SetzeDefaultWert() / SetzeAlleDefaults()
'
' TABELLEN-ZUORDNUNG (Frontend/Backend):
'   FRONTEND (lokal, tblConfig im FE-accdb):
'     tblConfig, tblSyncProfil, tblSyncProfilOrdner
'   BACKEND (Netzlaufwerk, verknuepft via modBackend):
'     tblSyncLauf, tblKontakte, tblOutlookOrdner, tblEmailThreads,
'     tblEmails, tblEmailContent, tblEmailEmpfaenger, tblEmailAnhaenge,
'     tblEmailStatus, tblProjekte, tblEmailProjekt
'
'   ErstelleAlleTabellen      -> Erstellt ALLE lokal (Ersteinrichtung)
'   ErstelleBackendTabellenInDB(db) -> Erstellt BE-Tabellen in externer DB
'                                      (wird von modBackend.ErstelleBackendDB
'                                       aufgerufen bei Backend-Verknuepfung)
'
' Aufruf:
'   ErstelleAlleTabellen       ' Einmalig im Direktbereich (Strg+G)
'   LoescheAlleTabellen        ' VORSICHT: Loescht alle Daten!
'   InitStandardConfig         ' Config-Defaults neu setzen
'
' Zum kompletten Neuaufbau:
'   LoescheAlleTabellen
'   ErstelleAlleTabellen
' ===========================================================================

Private Const SCHEMA_VERSION As String = "0.7.0"

' ---------------------------------------------------------------------------
' HAUPTROUTINE: Alle Tabellen erstellen
' ---------------------------------------------------------------------------
Public Sub ErstelleAlleTabellen()
    Debug.Print String(70, "=")
    Debug.Print "=== Schema-Erstellung v" & SCHEMA_VERSION & " - " & Now() & " ==="
    Debug.Print String(70, "=")

    Call Erstelle_tblConfig
    Call Erstelle_tblSyncLauf
    Call Erstelle_tblKontakte
    Call Erstelle_tblOutlookOrdner
    Call Erstelle_tblEmailThreads
    Call Erstelle_tblEmails
    Call Erstelle_tblEmailContent
    Call Erstelle_tblEmailEmpfaenger
    Call Erstelle_tblEmailAnhaenge
    Call Erstelle_tblEmailStatus
    Call Erstelle_tblSyncProfil
    Call Erstelle_tblSyncProfilOrdner
    Call Erstelle_tblProjekte
    Call Erstelle_tblEmailProjekt
    Call Erstelle_tblSyncJob
    Call Erstelle_tblSyncHeartbeat
    Call Erstelle_tblSyncControl
    Call Erstelle_tblWorkerLease
    Call Erstelle_tblWorkerTrace
    Call Erstelle_tblLog

    Call ErstelleIndizes
    Call InitStandardConfig

    Debug.Print String(70, "=")
    Debug.Print "=== Schema-Erstellung abgeschlossen ==="
    Debug.Print String(70, "=")
End Sub

' ---------------------------------------------------------------------------
' Alle Sync-Tabellen loeschen (VORSICHT - alle Daten gehen verloren!)
' ---------------------------------------------------------------------------
Public Sub LoescheAlleTabellen()
    Dim db As DAO.Database
    Dim tblNames As Variant
    Dim i As Long

    Set db = CurrentDb

    ' Reihenfolge beachten (abhaengige Tabellen zuerst)
    tblNames = Array(TBL_WORKER_TRACE, TBL_WORKER_LEASE, TBL_SYNC_CONTROL, TBL_SYNC_HEARTBEAT, TBL_SYNC_JOB, _
                     TBL_EMAIL_PROJEKT, TBL_SYNC_PROFIL_ORDNER, TBL_SYNC_PROFIL, TBL_EMAIL_STATUS, _
                     TBL_EMAIL_ANHAENGE, TBL_EMAIL_EMPFAENGER, TBL_EMAIL_CONTENT, TBL_EMAILS, _
                     TBL_EMAIL_THREADS, TBL_OUTLOOK_ORDNER, TBL_KONTAKTE, TBL_SYNC_LAUF, _
                     TBL_PROJEKTE, TBL_CONFIG)

    Debug.Print String(70, "=")
    Debug.Print "=== LOESCHE alle Tabellen ==="

    For i = LBound(tblNames) To UBound(tblNames)
        If TabelleExistiert(CStr(tblNames(i))) Then
            On Error Resume Next
            db.Execute "DROP TABLE [" & tblNames(i) & "]"
            If Err.Number = 0 Then
                Debug.Print "  [DROP] " & tblNames(i)
            Else
                Debug.Print "  [FAIL] " & tblNames(i) & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        Else
            Debug.Print "  [SKIP] " & tblNames(i) & " (existiert nicht)"
        End If
    Next i

    Debug.Print "=== Loeschen abgeschlossen ==="
    Debug.Print String(70, "=")
    Set db = Nothing
End Sub


' ===========================================================================
' EINZELNE TABELLEN
' ===========================================================================

Private Sub Erstelle_tblConfig()
    If TabelleExistiert(TBL_CONFIG) Then
        Debug.Print "  [SKIP] " & TBL_CONFIG & " (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_CONFIG & "] ("
    sql = sql & "ConfigID AUTOINCREMENT CONSTRAINT PK_Config PRIMARY KEY, "
    sql = sql & "Schluessel TEXT(100) NOT NULL, "
    sql = sql & "Wert TEXT(255), "
    sql = sql & "Beschreibung TEXT(255))"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Config_Key ON [" & TBL_CONFIG & "] (Schluessel)"
    Debug.Print "  [OK  ] " & TBL_CONFIG
End Sub

Private Sub Erstelle_tblSyncLauf()
    If TabelleExistiert(TBL_SYNC_LAUF) Then
        Debug.Print "  [SKIP] " & TBL_SYNC_LAUF & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_SYNC_LAUF & "] ("
    sql = sql & "SyncLaufID AUTOINCREMENT CONSTRAINT PK_SyncLauf PRIMARY KEY, "
    sql = sql & "StartZeit DATETIME, "
    sql = sql & "EndeZeit DATETIME, "
    sql = sql & "Status TEXT(20), "
    sql = sql & "AnzahlGelesen LONG, "
    sql = sql & "AnzahlNeu LONG, "
    sql = sql & "AnzahlDuplikate LONG, "
    sql = sql & "AnzahlFehler LONG, "
    sql = sql & "OrdnerPfad TEXT(255), "
    sql = sql & "Projekt TEXT(100), "
    sql = sql & "Phase TEXT(100))"
    db.Execute sql

    ' Defaults per DAO setzen (Access DDL kennt kein DEFAULT)
    SetzeDefaultWert db, TBL_SYNC_LAUF, "Status", "'Gestartet'"
    SetzeDefaultWert db, TBL_SYNC_LAUF, "AnzahlGelesen", "0"
    SetzeDefaultWert db, TBL_SYNC_LAUF, "AnzahlNeu", "0"
    SetzeDefaultWert db, TBL_SYNC_LAUF, "AnzahlDuplikate", "0"
    SetzeDefaultWert db, TBL_SYNC_LAUF, "AnzahlFehler", "0"

    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_SYNC_LAUF
End Sub

Private Sub Erstelle_tblKontakte()
    If TabelleExistiert(TBL_KONTAKTE) Then
        Debug.Print "  [SKIP] " & TBL_KONTAKTE & " (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_KONTAKTE & "] ("
    sql = sql & "KontaktID AUTOINCREMENT CONSTRAINT PK_Kontakte PRIMARY KEY, "
    sql = sql & "Anzeigename TEXT(255), "
    sql = sql & "Email TEXT(255), "
    sql = sql & "EmailTyp TEXT(10), "
    sql = sql & "Vorname TEXT(100), "
    sql = sql & "Nachname TEXT(100), "
    sql = sql & "Titel TEXT(50), "
    sql = sql & "Namenszusatz TEXT(100), "
    sql = sql & "Institution TEXT(255), "
    sql = sql & "Sortiername TEXT(255), "
    sql = sql & "KontaktTyp TEXT(20), "
    sql = sql & "ErstelltAm DATETIME, "
    sql = sql & "AktualisiertAm DATETIME)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE INDEX idx_Kontakte_Email ON [" & TBL_KONTAKTE & "] (Email)"
    Debug.Print "  [OK  ] " & TBL_KONTAKTE
End Sub

Private Sub Erstelle_tblOutlookOrdner()
    If TabelleExistiert(TBL_OUTLOOK_ORDNER) Then
        Debug.Print "  [SKIP] " & TBL_OUTLOOK_ORDNER & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_OUTLOOK_ORDNER & "] ("
    sql = sql & "OrdnerID AUTOINCREMENT CONSTRAINT PK_Ordner PRIMARY KEY, "
    sql = sql & "OrdnerName TEXT(255), "
    sql = sql & "OrdnerPfad TEXT(255), "
    sql = sql & "ParentID LONG, "
    sql = sql & "PostfachName TEXT(255), "
    sql = sql & "StoreID TEXT(255), "
    sql = sql & "ElementAnzahl LONG, "
    sql = sql & "LetzterSync DATETIME)"
    db.Execute sql

    SetzeDefaultWert db, TBL_OUTLOOK_ORDNER, "ParentID", "0"
    SetzeDefaultWert db, TBL_OUTLOOK_ORDNER, "ElementAnzahl", "0"

    db.Execute "CREATE UNIQUE INDEX idx_Ordner_Pfad ON [" & TBL_OUTLOOK_ORDNER & "] (OrdnerPfad)"
    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_OUTLOOK_ORDNER
End Sub

Private Sub Erstelle_tblEmailThreads()
    If TabelleExistiert(TBL_EMAIL_THREADS) Then
        Debug.Print "  [SKIP] " & TBL_EMAIL_THREADS & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_EMAIL_THREADS & "] ("
    sql = sql & "ThreadID AUTOINCREMENT CONSTRAINT PK_Threads PRIMARY KEY, "
    sql = sql & "ThreadBetreff TEXT(255), "
    sql = sql & "ThreadIdentifier TEXT(255), "
    sql = sql & "Antwortanzahl LONG, "
    sql = sql & "ErsterAbsender TEXT(255), "
    sql = sql & "ErstesMailDatum DATETIME, "
    sql = sql & "LetztesMailDatum DATETIME, "
    sql = sql & "ErstelltAm DATETIME)"
    db.Execute sql

    SetzeDefaultWert db, TBL_EMAIL_THREADS, "Antwortanzahl", "1"

    db.Execute "CREATE UNIQUE INDEX idx_Thread_Ident ON [" & TBL_EMAIL_THREADS & "] (ThreadIdentifier)"
    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_EMAIL_THREADS
End Sub

Private Sub Erstelle_tblEmails()
    If TabelleExistiert(TBL_EMAILS) Then
        Debug.Print "  [SKIP] " & TBL_EMAILS & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_EMAILS & "] ("
    sql = sql & "EmailID AUTOINCREMENT CONSTRAINT PK_Emails PRIMARY KEY, "
    sql = sql & "OutlookEntryID TEXT(255), "
    sql = sql & "UniqueHash TEXT(64), "
    sql = sql & "ThreadID LONG, "
    sql = sql & "OrdnerID LONG, "
    sql = sql & "KontaktID_Absender LONG, "
    sql = sql & "SyncLaufID LONG, "
    sql = sql & "Betreff TEXT(255), "
    sql = sql & "BetreffBereinigt TEXT(255), "
    sql = sql & "AbsenderName TEXT(255), "
    sql = sql & "AbsenderEmail TEXT(255), "
    sql = sql & "EmpfangenAm DATETIME, "
    sql = sql & "GesendetAm DATETIME, "
    sql = sql & "Groesse LONG, "
    sql = sql & "Wichtigkeit SHORT, "
    sql = sql & "Gelesen YESNO, "
    sql = sql & "HatAnhaenge YESNO, "
    sql = sql & "AnhangAnzahl SHORT, "
    sql = sql & "MessageClass TEXT(50), "
    sql = sql & "InternetMessageID TEXT(255), "
    sql = sql & "MSGDateiPfad TEXT(255), "
    sql = sql & "Status TEXT(20), "
    sql = sql & "ErstelltAm DATETIME)"
    db.Execute sql

    SetzeDefaultWert db, TBL_EMAILS, "ThreadID", "0"
    SetzeDefaultWert db, TBL_EMAILS, "OrdnerID", "0"
    SetzeDefaultWert db, TBL_EMAILS, "KontaktID_Absender", "0"
    SetzeDefaultWert db, TBL_EMAILS, "SyncLaufID", "0"
    SetzeDefaultWert db, TBL_EMAILS, "Groesse", "0"
    SetzeDefaultWert db, TBL_EMAILS, "Wichtigkeit", "1"
    SetzeDefaultWert db, TBL_EMAILS, "AnhangAnzahl", "0"
    SetzeDefaultWert db, TBL_EMAILS, "Status", "'Neu'"

    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_EMAILS
End Sub

Private Sub Erstelle_tblEmailContent()
    If TabelleExistiert(TBL_EMAIL_CONTENT) Then
        Debug.Print "  [SKIP] " & TBL_EMAIL_CONTENT & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_EMAIL_CONTENT & "] ("
    sql = sql & "ContentID AUTOINCREMENT CONSTRAINT PK_Content PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "HTMLBody MEMO, "
    sql = sql & "PlainTextBody MEMO, "
    sql = sql & "HatHTML YESNO, "
    sql = sql & "GroesseHTML LONG, "
    sql = sql & "GroesseText LONG)"
    db.Execute sql

    SetzeDefaultWert db, TBL_EMAIL_CONTENT, "GroesseHTML", "0"
    SetzeDefaultWert db, TBL_EMAIL_CONTENT, "GroesseText", "0"

    db.Execute "CREATE UNIQUE INDEX idx_Content_EmailID ON [" & TBL_EMAIL_CONTENT & "] (EmailID)"
    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_EMAIL_CONTENT
End Sub

Private Sub Erstelle_tblEmailEmpfaenger()
    If TabelleExistiert(TBL_EMAIL_EMPFAENGER) Then
        Debug.Print "  [SKIP] " & TBL_EMAIL_EMPFAENGER & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_EMAIL_EMPFAENGER & "] ("
    sql = sql & "EmpfaengerID AUTOINCREMENT CONSTRAINT PK_Empfaenger PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "KontaktID LONG, "
    sql = sql & "Typ TEXT(5), "
    sql = sql & "Anzeigename TEXT(255), "
    sql = sql & "Email TEXT(255))"
    db.Execute sql

    SetzeDefaultWert db, TBL_EMAIL_EMPFAENGER, "KontaktID", "0"

    db.Execute "CREATE INDEX idx_Empf_EmailID ON [" & TBL_EMAIL_EMPFAENGER & "] (EmailID)"
    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_EMAIL_EMPFAENGER
End Sub

Private Sub Erstelle_tblEmailAnhaenge()
    If TabelleExistiert(TBL_EMAIL_ANHAENGE) Then
        Debug.Print "  [SKIP] " & TBL_EMAIL_ANHAENGE & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_EMAIL_ANHAENGE & "] ("
    sql = sql & "AnhangID AUTOINCREMENT CONSTRAINT PK_Anhaenge PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "Dateiname TEXT(255), "
    sql = sql & "DateinameBereinigt TEXT(255), "
    sql = sql & "Erweiterung TEXT(20), "
    sql = sql & "Groesse LONG, "
    sql = sql & "MimeType TEXT(100), "
    sql = sql & "AnhangTyp SHORT, "
    sql = sql & "IstVersteckt YESNO, "
    sql = sql & "IstGespeichert YESNO, "
    sql = sql & "DateiPfad TEXT(255), "
    sql = sql & "ErstelltAm DATETIME)"
    db.Execute sql

    SetzeDefaultWert db, TBL_EMAIL_ANHAENGE, "Groesse", "0"
    SetzeDefaultWert db, TBL_EMAIL_ANHAENGE, "AnhangTyp", "1"

    db.Execute "CREATE INDEX idx_Anh_EmailID ON [" & TBL_EMAIL_ANHAENGE & "] (EmailID)"
    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_EMAIL_ANHAENGE
End Sub

Private Sub Erstelle_tblEmailStatus()
    If TabelleExistiert(TBL_EMAIL_STATUS) Then
        Debug.Print "  [SKIP] " & TBL_EMAIL_STATUS & " (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_EMAIL_STATUS & "] ("
    sql = sql & "StatusID AUTOINCREMENT CONSTRAINT PK_EmailStatus PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "Status TEXT(50), "
    sql = sql & "GeaendertVon TEXT(100), "
    sql = sql & "Bemerkung TEXT(255), "
    sql = sql & "GeaendertAm DATETIME)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE INDEX idx_Status_EmailID ON [" & TBL_EMAIL_STATUS & "] (EmailID)"
    Debug.Print "  [OK  ] " & TBL_EMAIL_STATUS
End Sub


' ---------------------------------------------------------------------------
' tblSyncProfil - Gespeicherte Sync-Profile (selektiver Sync)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblSyncProfil()
    If TabelleExistiert(TBL_SYNC_PROFIL) Then
        Debug.Print "  [SKIP] " & TBL_SYNC_PROFIL & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_SYNC_PROFIL & "] ("
    sql = sql & "ProfilID AUTOINCREMENT CONSTRAINT PK_SyncProfil PRIMARY KEY, "
    sql = sql & "ProfilName TEXT(100) NOT NULL, "
    sql = sql & "Beschreibung TEXT(255), "
    sql = sql & "IstAktiv YESNO, "
    sql = sql & "Projekt TEXT(100), "
    sql = sql & "Phase TEXT(100), "
    sql = sql & "MaxMailsProOrdner LONG, "
    sql = sql & "MaxTiefe SHORT, "
    sql = sql & "ExportPfad TEXT(255), "
    sql = sql & "ErstelltAm DATETIME)"
    db.Execute sql

    SetzeDefaultWert db, TBL_SYNC_PROFIL, "MaxMailsProOrdner", "500"
    SetzeDefaultWert db, TBL_SYNC_PROFIL, "MaxTiefe", "5"

    db.Execute "CREATE UNIQUE INDEX idx_Profil_Name ON [" & TBL_SYNC_PROFIL & "] (ProfilName)"
    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_SYNC_PROFIL
End Sub


' ---------------------------------------------------------------------------
' tblSyncProfilOrdner - Ordnerauswahl je Profil
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblSyncProfilOrdner()
    If TabelleExistiert(TBL_SYNC_PROFIL_ORDNER) Then
        Debug.Print "  [SKIP] " & TBL_SYNC_PROFIL_ORDNER & " (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_SYNC_PROFIL_ORDNER & "] ("
    sql = sql & "ID AUTOINCREMENT CONSTRAINT PK_SyncProfilOrdner PRIMARY KEY, "
    sql = sql & "ProfilID LONG NOT NULL, "
    sql = sql & "OrdnerPfad TEXT(255) NOT NULL, "
    sql = sql & "PostfachName TEXT(255), "
    sql = sql & "IstAktiv YESNO)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE INDEX idx_ProfilOrdner_Profil ON [" & TBL_SYNC_PROFIL_ORDNER & "] (ProfilID)"
    Debug.Print "  [OK  ] " & TBL_SYNC_PROFIL_ORDNER
End Sub


' ---------------------------------------------------------------------------
' tblProjekte - Projekt-Stammdaten (v0.6)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblProjekte()
    If TabelleExistiert(TBL_PROJEKTE) Then
        Debug.Print "  [SKIP] " & TBL_PROJEKTE & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_PROJEKTE & "] ("
    sql = sql & "ProjektID AUTOINCREMENT CONSTRAINT PK_Projekte PRIMARY KEY, "
    sql = sql & "Name TEXT(100) NOT NULL, "
    sql = sql & "Kuerzel TEXT(20), "
    sql = sql & "Beschreibung TEXT(255), "
    sql = sql & "Phase TEXT(100), "
    sql = sql & "Status TEXT(20), "
    sql = sql & "Farbe TEXT(7), "
    sql = sql & "SortierNr LONG, "
    sql = sql & "ErstelltVon TEXT(100), "
    sql = sql & "ErstelltAm DATETIME, "
    sql = sql & "AktualisiertAm DATETIME)"
    db.Execute sql

    SetzeDefaultWert db, TBL_PROJEKTE, "Status", "'Aktiv'"
    SetzeDefaultWert db, TBL_PROJEKTE, "SortierNr", "0"

    db.Execute "CREATE UNIQUE INDEX idx_Projekt_Name ON [" & TBL_PROJEKTE & "] (Name)"
    db.Execute "CREATE INDEX idx_Projekt_Status ON [" & TBL_PROJEKTE & "] (Status)"
    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_PROJEKTE
End Sub


' ---------------------------------------------------------------------------
' tblEmailProjekt - n:m Zuordnung Email <-> Projekt (v0.6)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblEmailProjekt()
    If TabelleExistiert(TBL_EMAIL_PROJEKT) Then
        Debug.Print "  [SKIP] " & TBL_EMAIL_PROJEKT & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_EMAIL_PROJEKT & "] ("
    sql = sql & "EmailProjektID AUTOINCREMENT CONSTRAINT PK_EmailProjekt PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "ProjektID LONG NOT NULL, "
    sql = sql & "Quelle TEXT(20), "
    sql = sql & "ZugeordnetVon TEXT(100), "
    sql = sql & "ZugeordnetAm DATETIME)"
    db.Execute sql

    SetzeDefaultWert db, TBL_EMAIL_PROJEKT, "Quelle", "'Manuell'"

    db.Execute "CREATE INDEX idx_EP_EmailID ON [" & TBL_EMAIL_PROJEKT & "] (EmailID)"
    db.Execute "CREATE INDEX idx_EP_ProjektID ON [" & TBL_EMAIL_PROJEKT & "] (ProjektID)"
    db.Execute "CREATE UNIQUE INDEX idx_EP_Unique ON [" & TBL_EMAIL_PROJEKT & "] (EmailID, ProjektID)"
    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_EMAIL_PROJEKT
End Sub


' ---------------------------------------------------------------------------
' tblSyncJob - Job-Queue fuer No-Admin Worker (v0.7)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblSyncJob()
    If TabelleExistiert(TBL_SYNC_JOB) Then
        Debug.Print "  [SKIP] " & TBL_SYNC_JOB & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_SYNC_JOB & "] ("
    sql = sql & "JobID AUTOINCREMENT CONSTRAINT PK_SyncJob PRIMARY KEY, "
    sql = sql & "CreatedAt DATETIME, "
    sql = sql & "CreatedBy TEXT(100), "
    sql = sql & "RequestedFolderPath TEXT(255), "
    sql = sql & "RequestedMaxMails LONG, "
    sql = sql & "RequestedSubfolders YESNO, "
    sql = sql & "Status TEXT(30), "
    sql = sql & "WorkerId TEXT(100), "
    sql = sql & "StartedAt DATETIME, "
    sql = sql & "FinishedAt DATETIME, "
    sql = sql & "LastError MEMO, "
    sql = sql & "Priority INTEGER)"
    db.Execute sql

    SetzeDefaultWert db, TBL_SYNC_JOB, "Status", "'" & JOB_STATUS_QUEUED & "'"
    SetzeDefaultWert db, TBL_SYNC_JOB, "Priority", "100"
    SetzeDefaultWert db, TBL_SYNC_JOB, "RequestedMaxMails", "500"
    SetzeDefaultWert db, TBL_SYNC_JOB, "RequestedSubfolders", "False"

    db.Execute "CREATE INDEX idx_SyncJob_Status ON [" & TBL_SYNC_JOB & "] (Status)"
    db.Execute "CREATE INDEX idx_SyncJob_CreatedAt ON [" & TBL_SYNC_JOB & "] (CreatedAt)"
    db.Execute "CREATE INDEX idx_SyncJob_PrioCreated ON [" & TBL_SYNC_JOB & "] (Priority, CreatedAt)"

    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_SYNC_JOB
End Sub


' ---------------------------------------------------------------------------
' tblSyncHeartbeat - Worker Heartbeat/Progress (v0.7)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblSyncHeartbeat()
    If TabelleExistiert(TBL_SYNC_HEARTBEAT) Then
        Debug.Print "  [SKIP] " & TBL_SYNC_HEARTBEAT & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_SYNC_HEARTBEAT & "] ("
    sql = sql & "WorkerId TEXT(100) CONSTRAINT PK_SyncHeartbeat PRIMARY KEY, "
    sql = sql & "JobID LONG, "
    sql = sql & "Stage TEXT(50), "
    sql = sql & "CurrentItem LONG, "
    sql = sql & "TotalItems LONG, "
    sql = sql & "COMRetries LONG, "
    sql = sql & "COMReconnects LONG, "
    sql = sql & "UpdatedAt DATETIME, "
    sql = sql & "LastMessage MEMO)"
    db.Execute sql

    SetzeDefaultWert db, TBL_SYNC_HEARTBEAT, "CurrentItem", "0"
    SetzeDefaultWert db, TBL_SYNC_HEARTBEAT, "TotalItems", "0"
    SetzeDefaultWert db, TBL_SYNC_HEARTBEAT, "COMRetries", "0"
    SetzeDefaultWert db, TBL_SYNC_HEARTBEAT, "COMReconnects", "0"

    db.Execute "CREATE INDEX idx_SyncHB_UpdatedAt ON [" & TBL_SYNC_HEARTBEAT & "] (UpdatedAt)"
    db.Execute "CREATE INDEX idx_SyncHB_JobID ON [" & TBL_SYNC_HEARTBEAT & "] (JobID)"

    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_SYNC_HEARTBEAT
End Sub


' ---------------------------------------------------------------------------
' tblSyncControl - Pause/Cancel-Flags je Job (v0.7)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblSyncControl()
    If TabelleExistiert(TBL_SYNC_CONTROL) Then
        Debug.Print "  [SKIP] " & TBL_SYNC_CONTROL & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_SYNC_CONTROL & "] ("
    sql = sql & "JobID LONG CONSTRAINT PK_SyncControl PRIMARY KEY, "
    sql = sql & "PauseRequested YESNO, "
    sql = sql & "CancelRequested YESNO, "
    sql = sql & "UpdatedAt DATETIME)"
    db.Execute sql

    SetzeDefaultWert db, TBL_SYNC_CONTROL, "PauseRequested", "False"
    SetzeDefaultWert db, TBL_SYNC_CONTROL, "CancelRequested", "False"

    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_SYNC_CONTROL
End Sub


' ---------------------------------------------------------------------------
' tblWorkerLease - Lease-Tabelle fuer Worker-Liveness (v0.7)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblWorkerLease()
    If TabelleExistiert(TBL_WORKER_LEASE) Then
        Debug.Print "  [SKIP] " & TBL_WORKER_LEASE & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_WORKER_LEASE & "] ("
    sql = sql & "WorkerId TEXT(100) CONSTRAINT PK_WorkerLease PRIMARY KEY, "
    sql = sql & "LeaseUntil DATETIME, "
    sql = sql & "UpdatedAt DATETIME, "
    sql = sql & "HostName TEXT(100), "
    sql = sql & "SessionUser TEXT(100))"
    db.Execute sql

    db.Execute "CREATE INDEX idx_WorkerLease_Until ON [" & TBL_WORKER_LEASE & "] (LeaseUntil)"
    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_WORKER_LEASE
End Sub


' ---------------------------------------------------------------------------
' tblWorkerTrace - gemeinsames Worker-Debug/Trace-Logging (v0.7)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblWorkerTrace()
    If TabelleExistiert(TBL_WORKER_TRACE) Then
        Debug.Print "  [SKIP] " & TBL_WORKER_TRACE & " (existiert bereits)"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim sql As String
    sql = "CREATE TABLE [" & TBL_WORKER_TRACE & "] ("
    sql = sql & "TraceID AUTOINCREMENT CONSTRAINT PK_WorkerTrace PRIMARY KEY, "
    sql = sql & "LoggedAt DATETIME, "
    sql = sql & "LogLevel INTEGER, "
    sql = sql & "LevelName TEXT(10), "
    sql = sql & "WorkerId TEXT(100), "
    sql = sql & "JobID LONG, "
    sql = sql & "Modul TEXT(100), "
    sql = sql & "Prozedur TEXT(100), "
    sql = sql & "Nachricht TEXT(255), "
    sql = sql & "Details LONGTEXT, "
    sql = sql & "HostName TEXT(100), "
    sql = sql & "SessionUser TEXT(100))"
    db.Execute sql

    db.Execute "CREATE INDEX idx_WorkerTrace_LoggedAt ON [" & TBL_WORKER_TRACE & "] (LoggedAt)"
    db.Execute "CREATE INDEX idx_WorkerTrace_Worker ON [" & TBL_WORKER_TRACE & "] (WorkerId, LoggedAt)"
    db.Execute "CREATE INDEX idx_WorkerTrace_Job ON [" & TBL_WORKER_TRACE & "] (JobID, LoggedAt)"
    db.Execute "CREATE INDEX idx_WorkerTrace_Level ON [" & TBL_WORKER_TRACE & "] (LogLevel, LoggedAt)"

    Set db = Nothing
    Debug.Print "  [OK  ] " & TBL_WORKER_TRACE
End Sub


' ---------------------------------------------------------------------------
' tblLog (Wrapper fuer modLogging.ErstelleLogTabelle)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblLog()
    ' DDL liegt in modLogging (Single Source of Truth)
    ErstelleLogTabelle
End Sub


' ===========================================================================
' INDIZES FUER HAUPTTABELLE
' ===========================================================================

Private Sub ErstelleIndizes()
    On Error Resume Next

    ' tblEmails - Performance-Indizes
    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Email_Hash ON [" & TBL_EMAILS & "] (UniqueHash)"
    CurrentDb.Execute "CREATE INDEX idx_Email_EntryID ON [" & TBL_EMAILS & "] (OutlookEntryID)"
    CurrentDb.Execute "CREATE INDEX idx_Email_ThreadID ON [" & TBL_EMAILS & "] (ThreadID)"
    CurrentDb.Execute "CREATE INDEX idx_Email_KontaktID ON [" & TBL_EMAILS & "] (KontaktID_Absender)"
    CurrentDb.Execute "CREATE INDEX idx_Email_OrdnerID ON [" & TBL_EMAILS & "] (OrdnerID)"
    CurrentDb.Execute "CREATE INDEX idx_Email_SyncLauf ON [" & TBL_EMAILS & "] (SyncLaufID)"
    CurrentDb.Execute "CREATE INDEX idx_Email_Datum ON [" & TBL_EMAILS & "] (EmpfangenAm)"
    CurrentDb.Execute "CREATE INDEX idx_Email_InternetMsgID ON [" & TBL_EMAILS & "] (InternetMessageID)"
    CurrentDb.Execute "CREATE INDEX idx_SyncJob_Status ON [" & TBL_SYNC_JOB & "] (Status)"
    CurrentDb.Execute "CREATE INDEX idx_SyncJob_CreatedAt ON [" & TBL_SYNC_JOB & "] (CreatedAt)"
    CurrentDb.Execute "CREATE INDEX idx_SyncHB_UpdatedAt ON [" & TBL_SYNC_HEARTBEAT & "] (UpdatedAt)"
    CurrentDb.Execute "CREATE INDEX idx_SyncHB_JobID ON [" & TBL_SYNC_HEARTBEAT & "] (JobID)"

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    Debug.Print "  [OK  ] Indizes erstellt"
End Sub


' ===========================================================================
' STANDARDKONFIGURATION
' ===========================================================================

Public Sub InitStandardConfig()
    Dim db As DAO.Database
    Set db = CurrentDb

    On Error Resume Next
    Call SetzeConfig(db, CFG_EXPORT_PFAD, Environ("USERPROFILE") & PATH_DEFAULT_FALLBACK, "Basis-Pfad fuer MSG- und Anhang-Export")
    Call SetzeConfig(db, CFG_MAX_MAILS, "500", "Maximale Anzahl Mails pro Sync-Durchlauf")
    Call SetzeConfig(db, CFG_ANHAENGE, "1", "Anhaenge auf Festplatte extrahieren (1=Ja / 0=Nein)")
    Call SetzeConfig(db, CFG_MSG_EXPORT, "1", "MSG-Dateien exportieren (1=Ja / 0=Nein)")
    Call SetzeConfig(db, CFG_SIGNATUR_FILTER, "1", "Versteckte Signatur-Bilder ueberspringen (1=Ja / 0=Nein)")
    Call SetzeConfig(db, CFG_LOG_LEVEL, "3", "Log-Level (0=Aus 1=Error 2=Warn 3=Info 4=Debug 5=Trace)")
    Call SetzeConfig(db, CFG_SCHEMA_VERSION, SCHEMA_VERSION, "Aktuelle Schema-Version")
    Call SetzeConfig(db, CFG_BACKEND_PFAD, "", "Pfad zur Backend-Datenbank (leer = lokal)")
    Call SetzeConfig(db, CFG_TEMP_PFAD, "", "Temp-Verzeichnis fuer Extraktion (leer = %TEMP%\OutlookSync\)")
    Call SetzeConfig(db, CFG_BUFFER_GROESSE, "25", "Anzahl Mails im Schreib-Puffer vor Flush (5-500)")
    Call SetzeConfig(db, CFG_NETZWERK_RETRIES, "3", "Anzahl Wiederholungsversuche bei Netzwerkfehlern")
    Call SetzeConfig(db, CFG_NETZWERK_PAUSE, "2000", "Millisekunden Pause zwischen Netzwerk-Retries")
    Call SetzeConfig(db, CFG_NETZWERK_TIMEOUT, "30", "Sekunden auf Netzwerk warten (Datei-Queue/Kopie)")
    Call SetzeConfig(db, CFG_RDO_PFAD, "", "Verzeichnis mit Redemption-DLLs (leer = Programmverzeichnis)")
    Call SetzeConfig(db, CFG_WORKER_POLL_MS, "3000", "Polling-Intervall Worker in Millisekunden")
    Call SetzeConfig(db, CFG_WORKER_HB_S, "10", "Heartbeat-Intervall Worker in Sekunden")
    Call SetzeConfig(db, CFG_WORKER_STALE_S, "90", "Schwelle fuer stale Worker/Jobs in Sekunden")
    On Error GoTo 0

    Debug.Print "  [OK  ] Standardkonfiguration gesetzt"
    Set db = Nothing
End Sub

Private Sub SetzeConfig(db As DAO.Database, strKey As String, strVal As String, strDesc As String)
    On Error Resume Next
    Dim lngCount As Long
    lngCount = DCount("*", TBL_CONFIG, "Schluessel='" & strKey & "'")
    If lngCount = 0 Then
        db.Execute "INSERT INTO [" & TBL_CONFIG & "] (Schluessel, Wert, Beschreibung) VALUES ('" & strKey & "', '" & Replace(strVal, "'", "''") & "', '" & Replace(strDesc, "'", "''") & "')"
    End If
    On Error GoTo 0
End Sub


' ===========================================================================
' HILFSFUNKTIONEN
' ===========================================================================

' Erstellt alle Backend-Tabellen + Indizes in einer externen DB
' Wird von modBackend.ErstelleBackendDB aufgerufen (Single Source of Truth)
' v0.5.2: DEFAULT entfernt -> SetzeDefaultWert per DAO
Public Function ErstelleBackendTabellenInDB(db As DAO.Database) As Boolean
    On Error GoTo ErrHandler

    Dim sql As String
    Dim lngOK As Long, lngSkip As Long, lngFail As Long
    lngOK = 0: lngSkip = 0: lngFail = 0

    ' Zugriff testen
    Debug.Print "    DB: " & db.Name
    db.TableDefs.Refresh

    ' --- tblSyncLauf ---
    If Not TabelleExistiertInDB(db, TBL_SYNC_LAUF) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_SYNC_LAUF & "] (" & _
              "SyncLaufID AUTOINCREMENT CONSTRAINT PK_SyncLauf PRIMARY KEY, " & _
              "StartZeit DATETIME, " & _
              "EndeZeit DATETIME, " & _
              "Status TEXT(20), " & _
              "AnzahlGelesen LONG, " & _
              "AnzahlNeu LONG, " & _
              "AnzahlDuplikate LONG, " & _
              "AnzahlFehler LONG, " & _
              "OrdnerPfad TEXT(255), " & _
              "Projekt TEXT(100), " & _
              "Phase TEXT(100))"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_SYNC_LAUF & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_SYNC_LAUF, "Status", "'Gestartet'"
            SetzeDefaultWert db, TBL_SYNC_LAUF, "AnzahlGelesen", "0"
            SetzeDefaultWert db, TBL_SYNC_LAUF, "AnzahlNeu", "0"
            SetzeDefaultWert db, TBL_SYNC_LAUF, "AnzahlDuplikate", "0"
            SetzeDefaultWert db, TBL_SYNC_LAUF, "AnzahlFehler", "0"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_SYNC_LAUF
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_SYNC_LAUF
    End If

    ' --- tblKontakte ---
    If Not TabelleExistiertInDB(db, TBL_KONTAKTE) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_KONTAKTE & "] (" & _
              "KontaktID AUTOINCREMENT CONSTRAINT PK_Kontakte PRIMARY KEY, " & _
              "Anzeigename TEXT(255), " & _
              "Email TEXT(255), " & _
              "EmailTyp TEXT(10), " & _
              "Vorname TEXT(100), " & _
              "Nachname TEXT(100), " & _
              "Titel TEXT(50), " & _
              "Namenszusatz TEXT(100), " & _
              "Institution TEXT(255), " & _
              "Sortiername TEXT(255), " & _
              "KontaktTyp TEXT(20), " & _
              "ErstelltAm DATETIME, " & _
              "AktualisiertAm DATETIME)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_KONTAKTE & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_KONTAKTE
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_KONTAKTE
    End If

    ' --- tblOutlookOrdner ---
    If Not TabelleExistiertInDB(db, TBL_OUTLOOK_ORDNER) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_OUTLOOK_ORDNER & "] (" & _
              "OrdnerID AUTOINCREMENT CONSTRAINT PK_Ordner PRIMARY KEY, " & _
              "OrdnerName TEXT(255), " & _
              "OrdnerPfad TEXT(255), " & _
              "ParentID LONG, " & _
              "PostfachName TEXT(255), " & _
              "StoreID TEXT(255), " & _
              "ElementAnzahl LONG, " & _
              "LetzterSync DATETIME)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_OUTLOOK_ORDNER & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_OUTLOOK_ORDNER, "ParentID", "0"
            SetzeDefaultWert db, TBL_OUTLOOK_ORDNER, "ElementAnzahl", "0"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_OUTLOOK_ORDNER
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_OUTLOOK_ORDNER
    End If

    ' --- tblEmailThreads ---
    If Not TabelleExistiertInDB(db, TBL_EMAIL_THREADS) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_EMAIL_THREADS & "] (" & _
              "ThreadID AUTOINCREMENT CONSTRAINT PK_Threads PRIMARY KEY, " & _
              "ThreadBetreff TEXT(255), " & _
              "ThreadIdentifier TEXT(255), " & _
              "Antwortanzahl LONG, " & _
              "ErsterAbsender TEXT(255), " & _
              "ErstesMailDatum DATETIME, " & _
              "LetztesMailDatum DATETIME, " & _
              "ErstelltAm DATETIME)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_EMAIL_THREADS & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_EMAIL_THREADS, "Antwortanzahl", "1"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_EMAIL_THREADS
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_EMAIL_THREADS
    End If

    ' --- tblEmails ---
    If Not TabelleExistiertInDB(db, TBL_EMAILS) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_EMAILS & "] (" & _
              "EmailID AUTOINCREMENT CONSTRAINT PK_Emails PRIMARY KEY, " & _
              "OutlookEntryID TEXT(255), " & _
              "UniqueHash TEXT(64), " & _
              "ThreadID LONG, " & _
              "OrdnerID LONG, " & _
              "KontaktID_Absender LONG, " & _
              "SyncLaufID LONG, " & _
              "Betreff TEXT(255), " & _
              "BetreffBereinigt TEXT(255), " & _
              "AbsenderName TEXT(255), " & _
              "AbsenderEmail TEXT(255), " & _
              "EmpfangenAm DATETIME, " & _
              "GesendetAm DATETIME, " & _
              "Groesse LONG, " & _
              "Wichtigkeit SHORT, " & _
              "Gelesen YESNO, " & _
              "HatAnhaenge YESNO, " & _
              "AnhangAnzahl SHORT, " & _
              "MessageClass TEXT(50), " & _
              "InternetMessageID TEXT(255), " & _
              "MSGDateiPfad TEXT(255), " & _
              "Status TEXT(20), " & _
              "ErstelltAm DATETIME)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_EMAILS & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_EMAILS, "ThreadID", "0"
            SetzeDefaultWert db, TBL_EMAILS, "OrdnerID", "0"
            SetzeDefaultWert db, TBL_EMAILS, "KontaktID_Absender", "0"
            SetzeDefaultWert db, TBL_EMAILS, "SyncLaufID", "0"
            SetzeDefaultWert db, TBL_EMAILS, "Groesse", "0"
            SetzeDefaultWert db, TBL_EMAILS, "Wichtigkeit", "1"
            SetzeDefaultWert db, TBL_EMAILS, "AnhangAnzahl", "0"
            SetzeDefaultWert db, TBL_EMAILS, "Status", "'Neu'"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_EMAILS
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_EMAILS
    End If

    ' --- tblEmailContent ---
    If Not TabelleExistiertInDB(db, TBL_EMAIL_CONTENT) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_EMAIL_CONTENT & "] (" & _
              "ContentID AUTOINCREMENT CONSTRAINT PK_Content PRIMARY KEY, " & _
              "EmailID LONG NOT NULL, " & _
              "HTMLBody MEMO, " & _
              "PlainTextBody MEMO, " & _
              "HatHTML YESNO, " & _
              "GroesseHTML LONG, " & _
              "GroesseText LONG)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_EMAIL_CONTENT & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_EMAIL_CONTENT, "GroesseHTML", "0"
            SetzeDefaultWert db, TBL_EMAIL_CONTENT, "GroesseText", "0"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_EMAIL_CONTENT
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_EMAIL_CONTENT
    End If

    ' --- tblEmailEmpfaenger ---
    If Not TabelleExistiertInDB(db, TBL_EMAIL_EMPFAENGER) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_EMAIL_EMPFAENGER & "] (" & _
              "EmpfaengerID AUTOINCREMENT CONSTRAINT PK_Empfaenger PRIMARY KEY, " & _
              "EmailID LONG NOT NULL, " & _
              "KontaktID LONG, " & _
              "Typ TEXT(5), " & _
              "Anzeigename TEXT(255), " & _
              "Email TEXT(255))"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_EMAIL_EMPFAENGER & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_EMAIL_EMPFAENGER, "KontaktID", "0"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_EMAIL_EMPFAENGER
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_EMAIL_EMPFAENGER
    End If

    ' --- tblEmailAnhaenge ---
    If Not TabelleExistiertInDB(db, TBL_EMAIL_ANHAENGE) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_EMAIL_ANHAENGE & "] (" & _
              "AnhangID AUTOINCREMENT CONSTRAINT PK_Anhaenge PRIMARY KEY, " & _
              "EmailID LONG NOT NULL, " & _
              "Dateiname TEXT(255), " & _
              "DateinameBereinigt TEXT(255), " & _
              "Erweiterung TEXT(20), " & _
              "Groesse LONG, " & _
              "MimeType TEXT(100), " & _
              "AnhangTyp SHORT, " & _
              "IstVersteckt YESNO, " & _
              "IstGespeichert YESNO, " & _
              "DateiPfad TEXT(255), " & _
              "ErstelltAm DATETIME)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_EMAIL_ANHAENGE & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_EMAIL_ANHAENGE, "Groesse", "0"
            SetzeDefaultWert db, TBL_EMAIL_ANHAENGE, "AnhangTyp", "1"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_EMAIL_ANHAENGE
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_EMAIL_ANHAENGE
    End If

    ' --- tblEmailStatus ---
    If Not TabelleExistiertInDB(db, TBL_EMAIL_STATUS) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_EMAIL_STATUS & "] (" & _
              "StatusID AUTOINCREMENT CONSTRAINT PK_EmailStatus PRIMARY KEY, " & _
              "EmailID LONG NOT NULL, " & _
              "Status TEXT(50), " & _
              "GeaendertVon TEXT(100), " & _
              "Bemerkung TEXT(255), " & _
              "GeaendertAm DATETIME)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_EMAIL_STATUS & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_EMAIL_STATUS
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_EMAIL_STATUS
    End If

    ' --- tblProjekte (v0.6) ---
    If Not TabelleExistiertInDB(db, TBL_PROJEKTE) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_PROJEKTE & "] (" & _
              "ProjektID AUTOINCREMENT CONSTRAINT PK_Projekte PRIMARY KEY, " & _
              "Name TEXT(100) NOT NULL, " & _
              "Kuerzel TEXT(20), " & _
              "Beschreibung TEXT(255), " & _
              "Phase TEXT(100), " & _
              "Status TEXT(20), " & _
              "Farbe TEXT(7), " & _
              "SortierNr LONG, " & _
              "ErstelltVon TEXT(100), " & _
              "ErstelltAm DATETIME, " & _
              "AktualisiertAm DATETIME)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_PROJEKTE & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_PROJEKTE, "Status", "'Aktiv'"
            SetzeDefaultWert db, TBL_PROJEKTE, "SortierNr", "0"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_PROJEKTE
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_PROJEKTE
    End If

    ' --- tblEmailProjekt (v0.6) ---
    If Not TabelleExistiertInDB(db, TBL_EMAIL_PROJEKT) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_EMAIL_PROJEKT & "] (" & _
              "EmailProjektID AUTOINCREMENT CONSTRAINT PK_EmailProjekt PRIMARY KEY, " & _
              "EmailID LONG NOT NULL, " & _
              "ProjektID LONG NOT NULL, " & _
              "Quelle TEXT(20), " & _
              "ZugeordnetVon TEXT(100), " & _
              "ZugeordnetAm DATETIME)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_EMAIL_PROJEKT & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_EMAIL_PROJEKT, "Quelle", "'Manuell'"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_EMAIL_PROJEKT
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_EMAIL_PROJEKT
    End If

    ' --- tblSyncJob (v0.7) ---
    If Not TabelleExistiertInDB(db, TBL_SYNC_JOB) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_SYNC_JOB & "] (" & _
              "JobID AUTOINCREMENT CONSTRAINT PK_SyncJob PRIMARY KEY, " & _
              "CreatedAt DATETIME, " & _
              "CreatedBy TEXT(100), " & _
              "RequestedFolderPath TEXT(255), " & _
              "RequestedMaxMails LONG, " & _
              "RequestedSubfolders YESNO, " & _
              "Status TEXT(30), " & _
              "WorkerId TEXT(100), " & _
              "StartedAt DATETIME, " & _
              "FinishedAt DATETIME, " & _
              "LastError MEMO, " & _
              "Priority INTEGER)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_SYNC_JOB & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_SYNC_JOB, "Status", "'" & JOB_STATUS_QUEUED & "'"
            SetzeDefaultWert db, TBL_SYNC_JOB, "Priority", "100"
            SetzeDefaultWert db, TBL_SYNC_JOB, "RequestedMaxMails", "500"
            SetzeDefaultWert db, TBL_SYNC_JOB, "RequestedSubfolders", "False"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_SYNC_JOB
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_SYNC_JOB
    End If

    ' --- tblSyncHeartbeat (v0.7) ---
    If Not TabelleExistiertInDB(db, TBL_SYNC_HEARTBEAT) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_SYNC_HEARTBEAT & "] (" & _
              "WorkerId TEXT(100) CONSTRAINT PK_SyncHeartbeat PRIMARY KEY, " & _
              "JobID LONG, " & _
              "Stage TEXT(50), " & _
              "CurrentItem LONG, " & _
              "TotalItems LONG, " & _
              "COMRetries LONG, " & _
              "COMReconnects LONG, " & _
              "UpdatedAt DATETIME, " & _
              "LastMessage MEMO)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_SYNC_HEARTBEAT & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_SYNC_HEARTBEAT, "CurrentItem", "0"
            SetzeDefaultWert db, TBL_SYNC_HEARTBEAT, "TotalItems", "0"
            SetzeDefaultWert db, TBL_SYNC_HEARTBEAT, "COMRetries", "0"
            SetzeDefaultWert db, TBL_SYNC_HEARTBEAT, "COMReconnects", "0"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_SYNC_HEARTBEAT
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_SYNC_HEARTBEAT
    End If

    ' --- tblSyncControl (v0.7) ---
    If Not TabelleExistiertInDB(db, TBL_SYNC_CONTROL) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_SYNC_CONTROL & "] (" & _
              "JobID LONG CONSTRAINT PK_SyncControl PRIMARY KEY, " & _
              "PauseRequested YESNO, " & _
              "CancelRequested YESNO, " & _
              "UpdatedAt DATETIME)"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_SYNC_CONTROL & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            SetzeDefaultWert db, TBL_SYNC_CONTROL, "PauseRequested", "False"
            SetzeDefaultWert db, TBL_SYNC_CONTROL, "CancelRequested", "False"
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_SYNC_CONTROL
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_SYNC_CONTROL
    End If

    ' --- tblWorkerLease (v0.7) ---
    If Not TabelleExistiertInDB(db, TBL_WORKER_LEASE) Then
        On Error Resume Next
        sql = "CREATE TABLE [" & TBL_WORKER_LEASE & "] (" & _
              "WorkerId TEXT(100) CONSTRAINT PK_WorkerLease PRIMARY KEY, " & _
              "LeaseUntil DATETIME, " & _
              "UpdatedAt DATETIME, " & _
              "HostName TEXT(100), " & _
              "SessionUser TEXT(100))"
        db.Execute sql
        If Err.Number <> 0 Then
            Debug.Print "    [FAIL] " & TBL_WORKER_LEASE & " - " & Err.Description
            lngFail = lngFail + 1: Err.Clear
        Else
            db.TableDefs.Refresh
            If Err.Number <> 0 Then Err.Clear
            lngOK = lngOK + 1
            Debug.Print "    [OK  ] " & TBL_WORKER_LEASE
        End If
        On Error GoTo 0
    Else
        lngSkip = lngSkip + 1
        Debug.Print "    [SKIP] " & TBL_WORKER_LEASE
    End If

    ' --- Indizes sicherstellen (idempotent, auch fuer bereits vorhandene Tabellen) ---
    On Error Resume Next
    db.Execute "CREATE INDEX idx_Kontakte_Email ON [" & TBL_KONTAKTE & "] (Email)"
    db.Execute "CREATE UNIQUE INDEX idx_Ordner_Pfad ON [" & TBL_OUTLOOK_ORDNER & "] (OrdnerPfad)"
    db.Execute "CREATE UNIQUE INDEX idx_Thread_Ident ON [" & TBL_EMAIL_THREADS & "] (ThreadIdentifier)"
    db.Execute "CREATE UNIQUE INDEX idx_Email_Hash ON [" & TBL_EMAILS & "] (UniqueHash)"
    db.Execute "CREATE INDEX idx_Email_EntryID ON [" & TBL_EMAILS & "] (OutlookEntryID)"
    db.Execute "CREATE INDEX idx_Email_ThreadID ON [" & TBL_EMAILS & "] (ThreadID)"
    db.Execute "CREATE INDEX idx_Email_KontaktID ON [" & TBL_EMAILS & "] (KontaktID_Absender)"
    db.Execute "CREATE INDEX idx_Email_OrdnerID ON [" & TBL_EMAILS & "] (OrdnerID)"
    db.Execute "CREATE INDEX idx_Email_SyncLauf ON [" & TBL_EMAILS & "] (SyncLaufID)"
    db.Execute "CREATE INDEX idx_Email_Datum ON [" & TBL_EMAILS & "] (EmpfangenAm)"
    db.Execute "CREATE UNIQUE INDEX idx_Content_EmailID ON [" & TBL_EMAIL_CONTENT & "] (EmailID)"
    db.Execute "CREATE INDEX idx_Empf_EmailID ON [" & TBL_EMAIL_EMPFAENGER & "] (EmailID)"
    db.Execute "CREATE INDEX idx_Anh_EmailID ON [" & TBL_EMAIL_ANHAENGE & "] (EmailID)"
    db.Execute "CREATE INDEX idx_Status_EmailID ON [" & TBL_EMAIL_STATUS & "] (EmailID)"
    db.Execute "CREATE INDEX idx_Email_InternetMsgID ON [" & TBL_EMAILS & "] (InternetMessageID)"
    db.Execute "CREATE UNIQUE INDEX idx_Projekt_Name ON [" & TBL_PROJEKTE & "] (Name)"
    db.Execute "CREATE INDEX idx_Projekt_Status ON [" & TBL_PROJEKTE & "] (Status)"
    db.Execute "CREATE INDEX idx_EP_EmailID ON [" & TBL_EMAIL_PROJEKT & "] (EmailID)"
    db.Execute "CREATE INDEX idx_EP_ProjektID ON [" & TBL_EMAIL_PROJEKT & "] (ProjektID)"
    db.Execute "CREATE UNIQUE INDEX idx_EP_Unique ON [" & TBL_EMAIL_PROJEKT & "] (EmailID, ProjektID)"
    db.Execute "CREATE INDEX idx_SyncJob_Status ON [" & TBL_SYNC_JOB & "] (Status)"
    db.Execute "CREATE INDEX idx_SyncJob_CreatedAt ON [" & TBL_SYNC_JOB & "] (CreatedAt)"
    db.Execute "CREATE INDEX idx_SyncJob_PrioCreated ON [" & TBL_SYNC_JOB & "] (Priority, CreatedAt)"
    db.Execute "CREATE INDEX idx_SyncHB_UpdatedAt ON [" & TBL_SYNC_HEARTBEAT & "] (UpdatedAt)"
    db.Execute "CREATE INDEX idx_SyncHB_JobID ON [" & TBL_SYNC_HEARTBEAT & "] (JobID)"
    db.Execute "CREATE INDEX idx_WorkerLease_Until ON [" & TBL_WORKER_LEASE & "] (LeaseUntil)"
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    Debug.Print "    [OK  ] Indizes sichergestellt"

    ' --- Zusammenfassung ---
    Dim strMsg As String
    strMsg = "Backend-Tabellen: " & lngOK & " erstellt, " & lngSkip & " uebersprungen"
    If lngFail > 0 Then strMsg = strMsg & ", " & lngFail & " Fehler"
    LogInfo strMsg, "SCHEMA"
    Debug.Print "    " & strMsg

    ErstelleBackendTabellenInDB = (lngFail = 0)
    Exit Function

ErrHandler:
    Debug.Print "    [ERROR] ErstelleBackendTabellenInDB: " & Err.Number & " - " & Err.Description
    LogError "ErstelleBackendTabellenInDB: " & Err.Number & " - " & Err.Description, "SCHEMA"
    ErstelleBackendTabellenInDB = False
End Function



' ---------------------------------------------------------------------------
' SetzeDefaultWert - Setzt DefaultValue per DAO (Access DDL kennt kein DEFAULT)
' ---------------------------------------------------------------------------
' Access SQL (CurrentDb.Execute "CREATE TABLE ...") unterstuetzt KEIN
' "DEFAULT x" in der Spaltendefinition -> Error 3290.
' Stattdessen: Tabelle zuerst ohne DEFAULT erstellen, dann:
'   db.TableDefs("tblName").Fields("FeldName").DefaultValue = "Wert"
'
' strDefault: Exakt wie in DAO erwartet:
'   Numerisch   -> "0"
'   Text        -> "'Gestartet'"  (mit einfachen Quotes!)
'   Boolean     -> "True" / "False"
' ---------------------------------------------------------------------------
Private Sub SetzeDefaultWert(db As DAO.Database, ByVal strTabelle As String, _
                             ByVal strFeld As String, ByVal strDefault As String)
    db.TableDefs(strTabelle).Fields(strFeld).DefaultValue = strDefault
End Sub



' Prueft ob eine Tabelle in einer beliebigen DB existiert (nicht nur CurrentDb)
' Wird von ErstelleBackendTabellenInDB benoetigt (db kann externe DB sein)
Private Function TabelleExistiertInDB(db As DAO.Database, ByVal strName As String) As Boolean
    On Error GoTo ErrHandler
    Dim td As DAO.TableDef
    TabelleExistiertInDB = False
    For Each td In db.TableDefs
        If td.Name = strName Then
            TabelleExistiertInDB = True
            Exit For
        End If
    Next td
    Exit Function

ErrHandler:
    Debug.Print "    [ERROR] TabelleExistiertInDB(" & strName & "): " & Err.Number & " - " & Err.Description
    TabelleExistiertInDB = False
End Function


' Konfigurationswert lesen
Public Function LeseConfig(ByVal strKey As String, Optional ByVal strDefault As String = "") As String
    On Error Resume Next
    Dim varVal As Variant
    varVal = DLookup("Wert", TBL_CONFIG, "Schluessel='" & strKey & "'")
    If IsNull(varVal) Or Err.Number <> 0 Then
        LeseConfig = strDefault
    Else
        LeseConfig = CStr(varVal)
    End If
    On Error GoTo 0
End Function

' Konfigurationswert schreiben (UPSERT: Insert wenn Key nicht existiert)
Public Sub SchreibeConfig(ByVal strKey As String, ByVal strVal As String)
    On Error Resume Next
    Dim strSafeVal As String
    strSafeVal = Replace(strVal, "'", "''")

    If DCount("*", TBL_CONFIG, "Schluessel='" & strKey & "'") = 0 Then
        CurrentDb.Execute "INSERT INTO [" & TBL_CONFIG & "] (Schluessel, Wert) VALUES ('" & _
                          strKey & "', '" & strSafeVal & "')"
    Else
        CurrentDb.Execute "UPDATE [" & TBL_CONFIG & "] SET Wert='" & strSafeVal & _
                          "' WHERE Schluessel='" & strKey & "'"
    End If
    On Error GoTo 0
End Sub


