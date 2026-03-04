Attribute VB_Name = "modSchema"
Option Compare Database
Option Explicit

' ===========================================================================
' modSchema - Tabellenschema-Verwaltung fuer OutlookSync
' ===========================================================================
' Erstellt alle 10 Tabellen + Indizes + Standardkonfiguration.
' Bestehende Tabellen werden NICHT ueberschrieben.
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

Private Const SCHEMA_VERSION As String = "0.3"

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
    tblNames = Array("tblSyncProfilOrdner", "tblSyncProfil", "tblEmailStatus", "tblEmailAnhaenge", "tblEmailEmpfaenger", "tblEmailContent", "tblEmails", "tblEmailThreads", "tblOutlookOrdner", "tblKontakte", "tblSyncLauf", "tblConfig")

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
    If TabelleExistiert("tblConfig") Then
        Debug.Print "  [SKIP] tblConfig (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblConfig ("
    sql = sql & "ConfigID AUTOINCREMENT CONSTRAINT PK_Config PRIMARY KEY, "
    sql = sql & "Schluessel TEXT(100) NOT NULL, "
    sql = sql & "Wert TEXT(255), "
    sql = sql & "Beschreibung TEXT(255))"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Config_Key ON tblConfig (Schluessel)"
    Debug.Print "  [OK  ] tblConfig"
End Sub

Private Sub Erstelle_tblSyncLauf()
    If TabelleExistiert("tblSyncLauf") Then
        Debug.Print "  [SKIP] tblSyncLauf (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblSyncLauf ("
    sql = sql & "SyncLaufID AUTOINCREMENT CONSTRAINT PK_SyncLauf PRIMARY KEY, "
    sql = sql & "StartZeit DATETIME, "
    sql = sql & "EndeZeit DATETIME, "
    sql = sql & "Status TEXT(20) DEFAULT 'Gestartet', "
    sql = sql & "AnzahlGelesen LONG DEFAULT 0, "
    sql = sql & "AnzahlNeu LONG DEFAULT 0, "
    sql = sql & "AnzahlDuplikate LONG DEFAULT 0, "
    sql = sql & "AnzahlFehler LONG DEFAULT 0, "
    sql = sql & "OrdnerPfad TEXT(255), "
    sql = sql & "Projekt TEXT(100), "
    sql = sql & "Phase TEXT(100))"
    CurrentDb.Execute sql

    Debug.Print "  [OK  ] tblSyncLauf"
End Sub

Private Sub Erstelle_tblKontakte()
    If TabelleExistiert("tblKontakte") Then
        Debug.Print "  [SKIP] tblKontakte (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblKontakte ("
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

    CurrentDb.Execute "CREATE INDEX idx_Kontakte_Email ON tblKontakte (Email)"
    Debug.Print "  [OK  ] tblKontakte"
End Sub

Private Sub Erstelle_tblOutlookOrdner()
    If TabelleExistiert("tblOutlookOrdner") Then
        Debug.Print "  [SKIP] tblOutlookOrdner (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblOutlookOrdner ("
    sql = sql & "OrdnerID AUTOINCREMENT CONSTRAINT PK_Ordner PRIMARY KEY, "
    sql = sql & "OrdnerName TEXT(255), "
    sql = sql & "OrdnerPfad TEXT(255), "
    sql = sql & "ParentID LONG DEFAULT 0, "
    sql = sql & "PostfachName TEXT(255), "
    sql = sql & "StoreID TEXT(255), "
    sql = sql & "ElementAnzahl LONG DEFAULT 0, "
    sql = sql & "LetzterSync DATETIME)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Ordner_Pfad ON tblOutlookOrdner (OrdnerPfad)"
    Debug.Print "  [OK  ] tblOutlookOrdner"
End Sub

Private Sub Erstelle_tblEmailThreads()
    If TabelleExistiert("tblEmailThreads") Then
        Debug.Print "  [SKIP] tblEmailThreads (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblEmailThreads ("
    sql = sql & "ThreadID AUTOINCREMENT CONSTRAINT PK_Threads PRIMARY KEY, "
    sql = sql & "ThreadBetreff TEXT(255), "
    sql = sql & "ThreadIdentifier TEXT(255), "
    sql = sql & "Antwortanzahl LONG DEFAULT 1, "
    sql = sql & "ErsterAbsender TEXT(255), "
    sql = sql & "ErstesMailDatum DATETIME, "
    sql = sql & "LetztesMailDatum DATETIME, "
    sql = sql & "ErstelltAm DATETIME)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Thread_Ident ON tblEmailThreads (ThreadIdentifier)"
    Debug.Print "  [OK  ] tblEmailThreads"
End Sub

Private Sub Erstelle_tblEmails()
    If TabelleExistiert("tblEmails") Then
        Debug.Print "  [SKIP] tblEmails (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblEmails ("
    sql = sql & "EmailID AUTOINCREMENT CONSTRAINT PK_Emails PRIMARY KEY, "
    sql = sql & "OutlookEntryID TEXT(255), "
    sql = sql & "UniqueHash TEXT(64), "
    sql = sql & "ThreadID LONG DEFAULT 0, "
    sql = sql & "OrdnerID LONG DEFAULT 0, "
    sql = sql & "KontaktID_Absender LONG DEFAULT 0, "
    sql = sql & "SyncLaufID LONG DEFAULT 0, "
    sql = sql & "Betreff TEXT(255), "
    sql = sql & "BetreffBereinigt TEXT(255), "
    sql = sql & "AbsenderName TEXT(255), "
    sql = sql & "AbsenderEmail TEXT(255), "
    sql = sql & "EmpfangenAm DATETIME, "
    sql = sql & "GesendetAm DATETIME, "
    sql = sql & "Groesse LONG DEFAULT 0, "
    sql = sql & "Wichtigkeit SHORT DEFAULT 1, "
    sql = sql & "Gelesen YESNO, "
    sql = sql & "HatAnhaenge YESNO, "
    sql = sql & "AnhangAnzahl SHORT DEFAULT 0, "
    sql = sql & "MessageClass TEXT(50), "
    sql = sql & "InternetMessageID TEXT(255), "
    sql = sql & "MSGDateiPfad TEXT(255), "
    sql = sql & "Status TEXT(20) DEFAULT 'Neu', "
    sql = sql & "ErstelltAm DATETIME)"
    CurrentDb.Execute sql

    Debug.Print "  [OK  ] tblEmails"
End Sub

Private Sub Erstelle_tblEmailContent()
    If TabelleExistiert("tblEmailContent") Then
        Debug.Print "  [SKIP] tblEmailContent (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblEmailContent ("
    sql = sql & "ContentID AUTOINCREMENT CONSTRAINT PK_Content PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "HTMLBody MEMO, "
    sql = sql & "PlainTextBody MEMO, "
    sql = sql & "HatHTML YESNO, "
    sql = sql & "GroesseHTML LONG DEFAULT 0, "
    sql = sql & "GroesseText LONG DEFAULT 0)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Content_EmailID ON tblEmailContent (EmailID)"
    Debug.Print "  [OK  ] tblEmailContent"
End Sub

Private Sub Erstelle_tblEmailEmpfaenger()
    If TabelleExistiert("tblEmailEmpfaenger") Then
        Debug.Print "  [SKIP] tblEmailEmpfaenger (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblEmailEmpfaenger ("
    sql = sql & "EmpfaengerID AUTOINCREMENT CONSTRAINT PK_Empfaenger PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "KontaktID LONG DEFAULT 0, "
    sql = sql & "Typ TEXT(5), "
    sql = sql & "Anzeigename TEXT(255), "
    sql = sql & "Email TEXT(255))"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE INDEX idx_Empf_EmailID ON tblEmailEmpfaenger (EmailID)"
    Debug.Print "  [OK  ] tblEmailEmpfaenger"
End Sub

Private Sub Erstelle_tblEmailAnhaenge()
    If TabelleExistiert("tblEmailAnhaenge") Then
        Debug.Print "  [SKIP] tblEmailAnhaenge (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblEmailAnhaenge ("
    sql = sql & "AnhangID AUTOINCREMENT CONSTRAINT PK_Anhaenge PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "Dateiname TEXT(255), "
    sql = sql & "DateinameBereinigt TEXT(255), "
    sql = sql & "Erweiterung TEXT(20), "
    sql = sql & "Groesse LONG DEFAULT 0, "
    sql = sql & "MimeType TEXT(100), "
    sql = sql & "AnhangTyp SHORT DEFAULT 1, "
    sql = sql & "IstVersteckt YESNO, "
    sql = sql & "IstGespeichert YESNO, "
    sql = sql & "DateiPfad TEXT(255), "
    sql = sql & "ErstelltAm DATETIME)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE INDEX idx_Anh_EmailID ON tblEmailAnhaenge (EmailID)"
    Debug.Print "  [OK  ] tblEmailAnhaenge"
End Sub

Private Sub Erstelle_tblEmailStatus()
    If TabelleExistiert("tblEmailStatus") Then
        Debug.Print "  [SKIP] tblEmailStatus (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblEmailStatus ("
    sql = sql & "StatusID AUTOINCREMENT CONSTRAINT PK_EmailStatus PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, "
    sql = sql & "Status TEXT(50), "
    sql = sql & "GeaendertVon TEXT(100), "
    sql = sql & "Bemerkung TEXT(255), "
    sql = sql & "GeaendertAm DATETIME)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE INDEX idx_Status_EmailID ON tblEmailStatus (EmailID)"
    Debug.Print "  [OK  ] tblEmailStatus"
End Sub


' ---------------------------------------------------------------------------
' tblSyncProfil - Gespeicherte Sync-Profile (selektiver Sync)
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblSyncProfil()
    If TabelleExistiert("tblSyncProfil") Then
        Debug.Print "  [SKIP] tblSyncProfil (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblSyncProfil ("
    sql = sql & "ProfilID AUTOINCREMENT CONSTRAINT PK_SyncProfil PRIMARY KEY, "
    sql = sql & "ProfilName TEXT(100) NOT NULL, "
    sql = sql & "Beschreibung TEXT(255), "
    sql = sql & "IstAktiv YESNO, "
    sql = sql & "Projekt TEXT(100), "
    sql = sql & "Phase TEXT(100), "
    sql = sql & "MaxMailsProOrdner LONG DEFAULT 500, "
    sql = sql & "MaxTiefe SHORT DEFAULT 5, "
    sql = sql & "ExportPfad TEXT(255), "
    sql = sql & "ErstelltAm DATETIME)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Profil_Name ON tblSyncProfil (ProfilName)"
    Debug.Print "  [OK  ] tblSyncProfil"
End Sub


' ---------------------------------------------------------------------------
' tblSyncProfilOrdner - Ordnerauswahl je Profil
' ---------------------------------------------------------------------------
Private Sub Erstelle_tblSyncProfilOrdner()
    If TabelleExistiert("tblSyncProfilOrdner") Then
        Debug.Print "  [SKIP] tblSyncProfilOrdner (existiert bereits)"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE tblSyncProfilOrdner ("
    sql = sql & "ID AUTOINCREMENT CONSTRAINT PK_SyncProfilOrdner PRIMARY KEY, "
    sql = sql & "ProfilID LONG NOT NULL, "
    sql = sql & "OrdnerPfad TEXT(255) NOT NULL, "
    sql = sql & "PostfachName TEXT(255), "
    sql = sql & "IstAktiv YESNO)"
    CurrentDb.Execute sql

    CurrentDb.Execute "CREATE INDEX idx_ProfilOrdner_Profil ON tblSyncProfilOrdner (ProfilID)"
    Debug.Print "  [OK  ] tblSyncProfilOrdner"
End Sub


' ===========================================================================
' INDIZES FUER HAUPTTABELLE
' ===========================================================================

Private Sub ErstelleIndizes()
    On Error Resume Next

    ' tblEmails - Performance-Indizes
    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Email_Hash ON tblEmails (UniqueHash)"
    CurrentDb.Execute "CREATE INDEX idx_Email_EntryID ON tblEmails (OutlookEntryID)"
    CurrentDb.Execute "CREATE INDEX idx_Email_ThreadID ON tblEmails (ThreadID)"
    CurrentDb.Execute "CREATE INDEX idx_Email_KontaktID ON tblEmails (KontaktID_Absender)"
    CurrentDb.Execute "CREATE INDEX idx_Email_OrdnerID ON tblEmails (OrdnerID)"
    CurrentDb.Execute "CREATE INDEX idx_Email_SyncLauf ON tblEmails (SyncLaufID)"
    CurrentDb.Execute "CREATE INDEX idx_Email_Datum ON tblEmails (EmpfangenAm)"

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
    Call SetzeConfig(db, "ExportBasisPfad", Environ("USERPROFILE") & "\OutlookSync\", "Basis-Pfad fuer MSG- und Anhang-Export")
    Call SetzeConfig(db, "MaxMailsProSync", "500", "Maximale Anzahl Mails pro Sync-Durchlauf")
    Call SetzeConfig(db, "AnhaengeExtrahieren", "1", "Anhaenge auf Festplatte extrahieren (1=Ja / 0=Nein)")
    Call SetzeConfig(db, "MSGExportieren", "1", "MSG-Dateien exportieren (1=Ja / 0=Nein)")
    Call SetzeConfig(db, "SignaturBilderFiltern", "1", "Versteckte Signatur-Bilder ueberspringen (1=Ja / 0=Nein)")
    Call SetzeConfig(db, "LogLevel", "3", "Log-Level (0=Aus 1=Error 2=Warn 3=Info 4=Debug 5=Trace)")
    Call SetzeConfig(db, "SchemaVersion", SCHEMA_VERSION, "Aktuelle Schema-Version")
    Call SetzeConfig(db, "BackendPfad", "", "Pfad zur Backend-Datenbank (leer = lokal)")
    Call SetzeConfig(db, "TempPfad", "", "Temp-Verzeichnis fuer Extraktion (leer = %TEMP%\OutlookSync\)")
    Call SetzeConfig(db, "BufferGroesse", "25", "Anzahl Mails im Schreib-Puffer vor Flush (5-500)")
    Call SetzeConfig(db, "NetzwerkRetries", "3", "Anzahl Wiederholungsversuche bei Netzwerkfehlern")
    Call SetzeConfig(db, "NetzwerkRetryPause", "2000", "Millisekunden Pause zwischen Netzwerk-Retries")
    On Error GoTo 0

    Debug.Print "  [OK  ] Standardkonfiguration gesetzt"
    Set db = Nothing
End Sub

Private Sub SetzeConfig(db As DAO.Database, strKey As String, strVal As String, strDesc As String)
    On Error Resume Next
    Dim lngCount As Long
    lngCount = DCount("*", "tblConfig", "Schluessel='" & strKey & "'")
    If lngCount = 0 Then
        db.Execute "INSERT INTO tblConfig (Schluessel, Wert, Beschreibung) VALUES ('" & strKey & "', '" & Replace(strVal, "'", "''") & "', '" & Replace(strDesc, "'", "''") & "')"
    End If
    On Error GoTo 0
End Sub


' ===========================================================================
' HILFSFUNKTIONEN
' ===========================================================================

' Prueft ob eine Tabelle in der Datenbank existiert
Public Function TabelleExistiert(ByVal strName As String) As Boolean
    Dim td As DAO.TableDef
    TabelleExistiert = False
    For Each td In CurrentDb.TableDefs
        If td.Name = strName Then
            TabelleExistiert = True
            Exit For
        End If
    Next td
End Function

' Konfigurationswert lesen
Public Function LeseConfig(ByVal strKey As String, Optional ByVal strDefault As String = "") As String
    On Error Resume Next
    Dim varVal As Variant
    varVal = DLookup("Wert", "tblConfig", "Schluessel='" & strKey & "'")
    If IsNull(varVal) Or Err.Number <> 0 Then
        LeseConfig = strDefault
    Else
        LeseConfig = CStr(varVal)
    End If
    On Error GoTo 0
End Function

' Konfigurationswert schreiben
Public Sub SchreibeConfig(ByVal strKey As String, ByVal strVal As String)
    On Error Resume Next
    CurrentDb.Execute "UPDATE tblConfig SET Wert='" & Replace(strVal, "'", "''") & "' WHERE Schluessel='" & strKey & "'"
    On Error GoTo 0
End Sub
