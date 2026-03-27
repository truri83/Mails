Attribute VB_Name = "modSchema"
Option Compare Database
Option Explicit

' ===========================================================================
' modSchema - Tabellenschema-Verwaltung fuer OutlookSync
' ===========================================================================
' Erstellt alle 12 Tabellen + Indizes + Standardkonfiguration.
' Bestehende Tabellen werden NICHT ueberschrieben, fehlende Spalten
' koennen ueber SchemaAktualisieren() nachgezogen werden.
'
' Alle DDL-Operationen laufen ueber modDDL (Backend-transparent).
'
' Oeffentliche API:
'   ErstelleAlleTabellen       ' Einmalig im Direktbereich (Strg+G)
'   LoescheAlleTabellen        ' VORSICHT: Loescht alle Daten!
'   SchemaAktualisieren        ' Fehlende Spalten/Indizes nachziehen
'   InitStandardConfig         ' Config-Defaults neu setzen
'
' Abhaengigkeiten: modDDL, modLogging
' ===========================================================================

Private Const SCHEMA_VERSION As String = "0.5"

' ---------------------------------------------------------------------------
' Tabellennamen (zentral, typsicher - oeffentlich fuer andere Module)
' ---------------------------------------------------------------------------
Public Const TBL_CONFIG         As String = "tblConfig"
Public Const TBL_SYNC_LAUF      As String = "tblSyncLauf"
Public Const TBL_KONTAKTE       As String = "tblKontakte"
Public Const TBL_ORDNER         As String = "tblOutlookOrdner"
Public Const TBL_THREADS        As String = "tblEmailThreads"
Public Const TBL_EMAILS         As String = "tblEmails"
Public Const TBL_CONTENT        As String = "tblEmailContent"
Public Const TBL_EMPFAENGER     As String = "tblEmailEmpfaenger"
Public Const TBL_ANHAENGE       As String = "tblEmailAnhaenge"
Public Const TBL_EMAIL_STATUS   As String = "tblEmailStatus"
Public Const TBL_PROFIL         As String = "tblSyncProfil"
Public Const TBL_PROFIL_ORDNER  As String = "tblSyncProfilOrdner"

' Alle Tabellen als Array (Reihenfolge: abhaengige zuerst fuer DROP)
Public Function AlleTabellen() As Variant
    AlleTabellen = Array(TBL_PROFIL_ORDNER, TBL_PROFIL, TBL_EMAIL_STATUS, TBL_ANHAENGE, TBL_EMPFAENGER, TBL_CONTENT, TBL_EMAILS, TBL_THREADS, TBL_ORDNER, TBL_KONTAKTE, TBL_SYNC_LAUF, TBL_CONFIG)
End Function


' ===========================================================================
' HAUPTROUTINEN
' ===========================================================================

Public Sub ErstelleAlleTabellen()
    Debug.Print String(70, "=")
    Debug.Print "=== Schema-Erstellung v" & SCHEMA_VERSION & " - " & Now() & " ==="
    Debug.Print String(70, "=")

    Erstelle_tblConfig
    Erstelle_tblSyncLauf
    Erstelle_tblKontakte
    Erstelle_tblOutlookOrdner
    Erstelle_tblEmailThreads
    Erstelle_tblEmails
    Erstelle_tblEmailContent
    Erstelle_tblEmailEmpfaenger
    Erstelle_tblEmailAnhaenge
    Erstelle_tblEmailStatus
    Erstelle_tblSyncProfil
    Erstelle_tblSyncProfilOrdner

    ErstelleAlleIndizes
    InitStandardConfig

    Debug.Print String(70, "=")
    Debug.Print "=== Schema-Erstellung abgeschlossen ==="
    Debug.Print String(70, "=")
End Sub

Public Sub LoescheAlleTabellen()
    Debug.Print String(70, "=")
    Debug.Print "=== LOESCHE alle Tabellen ==="

    Dim arr As Variant
    arr = AlleTabellen()
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If DDL_LoescheTabelle(CStr(arr(i))) Then
            Debug.Print "  [DROP] " & arr(i)
        End If
    Next i

    Debug.Print "=== Loeschen abgeschlossen ==="
    Debug.Print String(70, "=")
End Sub

' ---------------------------------------------------------------------------
' Schema-Migration: Fehlende Spalten, Defaults und Indizes nachziehen
' ---------------------------------------------------------------------------
Public Sub SchemaAktualisieren()
    Debug.Print String(70, "=")
    Debug.Print "=== Schema-Aktualisierung v" & SCHEMA_VERSION & " - " & Now() & " ==="

    ' --- Neue Spalten in kuenftigen Versionen hier einfuegen ---
    ' DDL_SichereSpalte TBL_EMAILS, "NeuesFeld", "TEXT(100)"
    ' DDL_SichereSpalte TBL_KONTAKTE, "Telefon", "TEXT(50)"

    ' Alle Defaults sicherstellen (idempotent, auch bei bestehenden Tabellen)
    SichereFeldDefaults

    ' Indizes nachziehen
    ErstelleAlleIndizes

    ' Config-Version aktualisieren
    SchreibeConfig "SchemaVersion", SCHEMA_VERSION

    Debug.Print "=== Schema-Aktualisierung abgeschlossen ==="
    Debug.Print String(70, "=")
End Sub


' ===========================================================================
' TABELLENERSTELLUNG (Private)
' ===========================================================================

Private Sub Erstelle_tblConfig()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_CONFIG & " ("
    sql = sql & "ConfigID AUTOINCREMENT CONSTRAINT PK_Config PRIMARY KEY, "
    sql = sql & "Schluessel TEXT(100) NOT NULL, "
    sql = sql & "Wert TEXT(255), "
    sql = sql & "Beschreibung TEXT(255))"

    If DDL_ErstelleTabelle(TBL_CONFIG, sql) Then
        DDL_SichererIndex TBL_CONFIG, "idx_Config_Key", "Schluessel", True
    End If
End Sub

Private Sub Erstelle_tblSyncLauf()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_SYNC_LAUF & " ("
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

    DDL_ErstelleTabelle TBL_SYNC_LAUF, sql
End Sub

Private Sub Erstelle_tblKontakte()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_KONTAKTE & " ("
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

    If DDL_ErstelleTabelle(TBL_KONTAKTE, sql) Then
        DDL_SichererIndex TBL_KONTAKTE, "idx_Kontakte_Email", "Email"
    End If
End Sub

Private Sub Erstelle_tblOutlookOrdner()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_ORDNER & " ("
    sql = sql & "OrdnerID AUTOINCREMENT CONSTRAINT PK_Ordner PRIMARY KEY, "
    sql = sql & "OrdnerName TEXT(255), "
    sql = sql & "OrdnerPfad TEXT(255), "
    sql = sql & "ParentID LONG, "
    sql = sql & "PostfachName TEXT(255), "
    sql = sql & "StoreID TEXT(255), "
    sql = sql & "ElementAnzahl LONG, "
    sql = sql & "LetzterSync DATETIME)"

    If DDL_ErstelleTabelle(TBL_ORDNER, sql) Then
        DDL_SichererIndex TBL_ORDNER, "idx_Ordner_Pfad", "OrdnerPfad", True
    End If
End Sub

Private Sub Erstelle_tblEmailThreads()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_THREADS & " ("
    sql = sql & "ThreadID AUTOINCREMENT CONSTRAINT PK_Threads PRIMARY KEY, "
    sql = sql & "ThreadBetreff TEXT(255), "
    sql = sql & "ThreadIdentifier TEXT(255), "
    sql = sql & "Antwortanzahl LONG, "
    sql = sql & "ErsterAbsender TEXT(255), "
    sql = sql & "ErstesMailDatum DATETIME, "
    sql = sql & "LetztesMailDatum DATETIME, "
    sql = sql & "ErstelltAm DATETIME)"

    If DDL_ErstelleTabelle(TBL_THREADS, sql) Then
        DDL_SichererIndex TBL_THREADS, "idx_Thread_Ident", "ThreadIdentifier", True
    End If
End Sub

Private Sub Erstelle_tblEmails()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_EMAILS & " ("
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

    DDL_ErstelleTabelle TBL_EMAILS, sql
End Sub

Private Sub Erstelle_tblEmailContent()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_CONTENT & " ("
    sql = sql & "ContentID AUTOINCREMENT CONSTRAINT PK_Content PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "HTMLBody MEMO, "
    sql = sql & "PlainTextBody MEMO, "
    sql = sql & "HatHTML YESNO, "
    sql = sql & "GroesseHTML LONG, "
    sql = sql & "GroesseText LONG)"

    If DDL_ErstelleTabelle(TBL_CONTENT, sql) Then
        DDL_SichererIndex TBL_CONTENT, "idx_Content_EmailID", "EmailID", True
    End If
End Sub

Private Sub Erstelle_tblEmailEmpfaenger()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_EMPFAENGER & " ("
    sql = sql & "EmpfaengerID AUTOINCREMENT CONSTRAINT PK_Empfaenger PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "KontaktID LONG, "
    sql = sql & "Typ TEXT(5), "
    sql = sql & "Anzeigename TEXT(255), "
    sql = sql & "Email TEXT(255))"

    If DDL_ErstelleTabelle(TBL_EMPFAENGER, sql) Then
        DDL_SichererIndex TBL_EMPFAENGER, "idx_Empf_EmailID", "EmailID"
    End If
End Sub

Private Sub Erstelle_tblEmailAnhaenge()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_ANHAENGE & " ("
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

    If DDL_ErstelleTabelle(TBL_ANHAENGE, sql) Then
        DDL_SichererIndex TBL_ANHAENGE, "idx_Anh_EmailID", "EmailID"
    End If
End Sub

Private Sub Erstelle_tblEmailStatus()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_EMAIL_STATUS & " ("
    sql = sql & "StatusID AUTOINCREMENT CONSTRAINT PK_EmailStatus PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "Status TEXT(50), "
    sql = sql & "GeaendertVon TEXT(100), "
    sql = sql & "Bemerkung TEXT(255), "
    sql = sql & "GeaendertAm DATETIME)"

    If DDL_ErstelleTabelle(TBL_EMAIL_STATUS, sql) Then
        DDL_SichererIndex TBL_EMAIL_STATUS, "idx_Status_EmailID", "EmailID"
    End If
End Sub

Private Sub Erstelle_tblSyncProfil()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_PROFIL & " ("
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

    If DDL_ErstelleTabelle(TBL_PROFIL, sql) Then
        DDL_SichererIndex TBL_PROFIL, "idx_Profil_Name", "ProfilName", True
    End If
End Sub

Private Sub Erstelle_tblSyncProfilOrdner()
    Dim sql As String
    sql = "CREATE TABLE " & TBL_PROFIL_ORDNER & " ("
    sql = sql & "ID AUTOINCREMENT CONSTRAINT PK_SyncProfilOrdner PRIMARY KEY, "
    sql = sql & "ProfilID LONG NOT NULL, "
    sql = sql & "OrdnerPfad TEXT(255) NOT NULL, "
    sql = sql & "PostfachName TEXT(255), "
    sql = sql & "IstAktiv YESNO)"

    If DDL_ErstelleTabelle(TBL_PROFIL_ORDNER, sql) Then
        DDL_SichererIndex TBL_PROFIL_ORDNER, "idx_ProfilOrdner_Profil", "ProfilID"
    End If
End Sub


' ===========================================================================
' DEFAULTS VIA DAO (nicht per CREATE TABLE DEFAULT, das geht nur bei Neuerstellung)
' ===========================================================================

' Setzt alle Feld-Defaults zentral via DAO (idempotent, Backend-transparent)
' Wird nach ErstelleAlleTabellen UND bei SchemaAktualisieren aufgerufen.
Private Sub SichereFeldDefaults()
    ' --- tblSyncLauf ---
    DDL_SetzeFeldDefault TBL_SYNC_LAUF, "Status", """Gestartet"""
    DDL_SetzeFeldDefault TBL_SYNC_LAUF, "AnzahlGelesen", "0"
    DDL_SetzeFeldDefault TBL_SYNC_LAUF, "AnzahlNeu", "0"
    DDL_SetzeFeldDefault TBL_SYNC_LAUF, "AnzahlDuplikate", "0"
    DDL_SetzeFeldDefault TBL_SYNC_LAUF, "AnzahlFehler", "0"

    ' --- tblOutlookOrdner ---
    DDL_SetzeFeldDefault TBL_ORDNER, "ParentID", "0"
    DDL_SetzeFeldDefault TBL_ORDNER, "ElementAnzahl", "0"

    ' --- tblEmailThreads ---
    DDL_SetzeFeldDefault TBL_THREADS, "Antwortanzahl", "1"

    ' --- tblEmails ---
    DDL_SetzeFeldDefault TBL_EMAILS, "ThreadID", "0"
    DDL_SetzeFeldDefault TBL_EMAILS, "OrdnerID", "0"
    DDL_SetzeFeldDefault TBL_EMAILS, "KontaktID_Absender", "0"
    DDL_SetzeFeldDefault TBL_EMAILS, "SyncLaufID", "0"
    DDL_SetzeFeldDefault TBL_EMAILS, "Groesse", "0"
    DDL_SetzeFeldDefault TBL_EMAILS, "Wichtigkeit", "1"
    DDL_SetzeFeldDefault TBL_EMAILS, "AnhangAnzahl", "0"
    DDL_SetzeFeldDefault TBL_EMAILS, "Status", """Neu"""

    ' --- tblEmailContent ---
    DDL_SetzeFeldDefault TBL_CONTENT, "GroesseHTML", "0"
    DDL_SetzeFeldDefault TBL_CONTENT, "GroesseText", "0"

    ' --- tblEmailEmpfaenger ---
    DDL_SetzeFeldDefault TBL_EMPFAENGER, "KontaktID", "0"

    ' --- tblEmailAnhaenge ---
    DDL_SetzeFeldDefault TBL_ANHAENGE, "Groesse", "0"
    DDL_SetzeFeldDefault TBL_ANHAENGE, "AnhangTyp", "1"

    ' --- tblSyncProfil ---
    DDL_SetzeFeldDefault TBL_PROFIL, "MaxMailsProOrdner", "500"
    DDL_SetzeFeldDefault TBL_PROFIL, "MaxTiefe", "5"

    Debug.Print "  [OK  ] Feld-Defaults gesetzt (via DAO)"
End Sub


' ===========================================================================
' INDIZES (zentral)
' ===========================================================================

Private Sub ErstelleAlleIndizes()
    DDL_SichererIndex TBL_EMAILS, "idx_Email_Hash", "UniqueHash", True
    DDL_SichererIndex TBL_EMAILS, "idx_Email_EntryID", "OutlookEntryID"
    DDL_SichererIndex TBL_EMAILS, "idx_Email_ThreadID", "ThreadID"
    DDL_SichererIndex TBL_EMAILS, "idx_Email_KontaktID", "KontaktID_Absender"
    DDL_SichererIndex TBL_EMAILS, "idx_Email_OrdnerID", "OrdnerID"
    DDL_SichererIndex TBL_EMAILS, "idx_Email_SyncLauf", "SyncLaufID"
    DDL_SichererIndex TBL_EMAILS, "idx_Email_Datum", "EmpfangenAm"

    Debug.Print "  [OK  ] Indizes erstellt/geprueft"
End Sub


' ===========================================================================
' STANDARDKONFIGURATION
' ===========================================================================

Public Sub InitStandardConfig()
    Dim db As DAO.Database
    Set db = CurrentDb

    On Error Resume Next
    SetzeConfig db, "ExportBasisPfad", Environ("USERPROFILE") & "\OutlookSync\", "Basis-Pfad fuer MSG- und Anhang-Export"
    SetzeConfig db, "MaxMailsProSync", "500", "Maximale Anzahl Mails pro Sync-Durchlauf"
    SetzeConfig db, "AnhaengeExtrahieren", "1", "Anhaenge auf Festplatte extrahieren (1=Ja / 0=Nein)"
    SetzeConfig db, "MSGExportieren", "1", "MSG-Dateien exportieren (1=Ja / 0=Nein)"
    SetzeConfig db, "SignaturBilderFiltern", "1", "Versteckte Signatur-Bilder ueberspringen (1=Ja / 0=Nein)"
    SetzeConfig db, "LogLevel", "3", "Log-Level (0=Aus 1=Error 2=Warn 3=Info 4=Debug 5=Trace)"
    SetzeConfig db, "SchemaVersion", SCHEMA_VERSION, "Aktuelle Schema-Version"
    SetzeConfig db, "BackendPfad", "", "Pfad zur Backend-Datenbank (leer = lokal)"
    SetzeConfig db, "TempPfad", "", "Temp-Verzeichnis fuer Extraktion (leer = %TEMP%\OutlookSync\)"
    SetzeConfig db, "BufferGroesse", "25", "Anzahl Mails im Schreib-Puffer vor Flush (5-500)"
    SetzeConfig db, "NetzwerkRetries", "3", "Anzahl Wiederholungsversuche bei Netzwerkfehlern"
    SetzeConfig db, "NetzwerkRetryPause", "2000", "Millisekunden Pause zwischen Netzwerk-Retries"
    On Error GoTo 0

    ' Defaults fuer alle Tabellen via DAO (Backend-transparent)
    SichereFeldDefaults

    Debug.Print "  [OK  ] Standardkonfiguration gesetzt"
    Set db = Nothing
End Sub

Private Sub SetzeConfig(db As DAO.Database, strKey As String, strVal As String, strDesc As String)
    On Error Resume Next
    Dim lngCount As Long
    lngCount = DCount("*", TBL_CONFIG, "Schluessel='" & strKey & "'")
    If lngCount = 0 Then
        db.Execute "INSERT INTO " & TBL_CONFIG & " (Schluessel, Wert, Beschreibung) VALUES ('" & strKey & "', '" & Replace(strVal, "'", "''") & "', '" & Replace(strDesc, "'", "''") & "')"
    End If
    On Error GoTo 0
End Sub


' ===========================================================================
' CONFIG LESEN/SCHREIBEN (nutzt modDDL nicht, da tblConfig immer lokal)
' ===========================================================================

Public Function LeseConfig(ByVal strKey As String, Optional ByVal strDefault As String = "") As String
    On Error Resume Next
    Dim varVal As Variant
    varVal = DLookup("Wert", TBL_CONFIG, "Schluessel='" & strKey & "'")
    If IsNull(varVal) Or Err.Number <> 0 Then
        LeseConfig = strDefault
    Else
        LeseConfig = CStr(varVal)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Public Sub SchreibeConfig(ByVal strKey As String, ByVal strVal As String)
    On Error Resume Next
    CurrentDb.Execute "UPDATE " & TBL_CONFIG & " SET Wert='" & Replace(strVal, "'", "''") & "' WHERE Schluessel='" & strKey & "'"
    Err.Clear
    On Error GoTo 0
End Sub
