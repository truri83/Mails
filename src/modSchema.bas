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
    tblNames = Array("tblSyncProfilOrdner", "tblSyncProfil", _
                     "tblEmailStatus", "tblEmailAnhaenge", "tblEmailEmpfaenger", _
                     "tblEmailContent", "tblEmails", "tblEmailThreads", _
                     "tblOutlookOrdner", "tblKontakte", "tblSyncLauf", "tblConfig")

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

    CurrentDb.Execute _
        "CREATE TABLE tblConfig (" & _
        "  ConfigID AUTOINCREMENT CONSTRAINT PK_Config PRIMARY KEY," & _
        "  Schluessel TEXT(100) NOT NULL," & _
        "  Wert TEXT(255)," & _
        "  Beschreibung TEXT(255)" & _
        ")"

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Config_Key ON tblConfig (Schluessel)"
    Debug.Print "  [OK  ] tblConfig"
End Sub

Private Sub Erstelle_tblSyncLauf()
    If TabelleExistiert("tblSyncLauf") Then
        Debug.Print "  [SKIP] tblSyncLauf (existiert bereits)"
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE tblSyncLauf (" & _
        "  SyncLaufID AUTOINCREMENT CONSTRAINT PK_SyncLauf PRIMARY KEY," & _
        "  StartZeit DATETIME," & _
        "  EndeZeit DATETIME," & _
        "  Status TEXT(20) DEFAULT 'Gestartet'," & _
        "  AnzahlGelesen LONG DEFAULT 0," & _
        "  AnzahlNeu LONG DEFAULT 0," & _
        "  AnzahlDuplikate LONG DEFAULT 0," & _
        "  AnzahlFehler LONG DEFAULT 0," & _
        "  OrdnerPfad TEXT(255)," & _
        "  Projekt TEXT(100)," & _
        "  Phase TEXT(100)" & _
        ")"

    Debug.Print "  [OK  ] tblSyncLauf"
End Sub

Private Sub Erstelle_tblKontakte()
    If TabelleExistiert("tblKontakte") Then
        Debug.Print "  [SKIP] tblKontakte (existiert bereits)"
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE tblKontakte (" & _
        "  KontaktID AUTOINCREMENT CONSTRAINT PK_Kontakte PRIMARY KEY," & _
        "  Anzeigename TEXT(255)," & _
        "  Email TEXT(255)," & _
        "  EmailTyp TEXT(10)," & _
        "  Vorname TEXT(100)," & _
        "  Nachname TEXT(100)," & _
        "  Titel TEXT(50)," & _
        "  Namenszusatz TEXT(100)," & _
        "  Institution TEXT(255)," & _
        "  Sortiername TEXT(255)," & _
        "  KontaktTyp TEXT(20)," & _
        "  ErstelltAm DATETIME," & _
        "  AktualisiertAm DATETIME" & _
        ")"

    CurrentDb.Execute "CREATE INDEX idx_Kontakte_Email ON tblKontakte (Email)"
    Debug.Print "  [OK  ] tblKontakte"
End Sub

Private Sub Erstelle_tblOutlookOrdner()
    If TabelleExistiert("tblOutlookOrdner") Then
        Debug.Print "  [SKIP] tblOutlookOrdner (existiert bereits)"
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE tblOutlookOrdner (" & _
        "  OrdnerID AUTOINCREMENT CONSTRAINT PK_Ordner PRIMARY KEY," & _
        "  OrdnerName TEXT(255)," & _
        "  OrdnerPfad TEXT(255)," & _
        "  ParentID LONG DEFAULT 0," & _
        "  PostfachName TEXT(255)," & _
        "  StoreID TEXT(255)," & _
        "  ElementAnzahl LONG DEFAULT 0," & _
        "  LetzterSync DATETIME" & _
        ")"

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Ordner_Pfad ON tblOutlookOrdner (OrdnerPfad)"
    Debug.Print "  [OK  ] tblOutlookOrdner"
End Sub

Private Sub Erstelle_tblEmailThreads()
    If TabelleExistiert("tblEmailThreads") Then
        Debug.Print "  [SKIP] tblEmailThreads (existiert bereits)"
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE tblEmailThreads (" & _
        "  ThreadID AUTOINCREMENT CONSTRAINT PK_Threads PRIMARY KEY," & _
        "  ThreadBetreff TEXT(255)," & _
        "  ThreadIdentifier TEXT(255)," & _
        "  Antwortanzahl LONG DEFAULT 1," & _
        "  ErsterAbsender TEXT(255)," & _
        "  ErstesMailDatum DATETIME," & _
        "  LetztesMailDatum DATETIME," & _
        "  ErstelltAm DATETIME" & _
        ")"

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Thread_Ident ON tblEmailThreads (ThreadIdentifier)"
    Debug.Print "  [OK  ] tblEmailThreads"
End Sub

Private Sub Erstelle_tblEmails()
    If TabelleExistiert("tblEmails") Then
        Debug.Print "  [SKIP] tblEmails (existiert bereits)"
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE tblEmails (" & _
        "  EmailID AUTOINCREMENT CONSTRAINT PK_Emails PRIMARY KEY," & _
        "  OutlookEntryID TEXT(255)," & _
        "  UniqueHash TEXT(64)," & _
        "  ThreadID LONG DEFAULT 0," & _
        "  OrdnerID LONG DEFAULT 0," & _
        "  KontaktID_Absender LONG DEFAULT 0," & _
        "  SyncLaufID LONG DEFAULT 0," & _
        "  Betreff TEXT(255)," & _
        "  BetreffBereinigt TEXT(255)," & _
        "  AbsenderName TEXT(255)," & _
        "  AbsenderEmail TEXT(255)," & _
        "  EmpfangenAm DATETIME," & _
        "  GesendetAm DATETIME," & _
        "  Groesse LONG DEFAULT 0," & _
        "  Wichtigkeit SHORT DEFAULT 1," & _
        "  Gelesen YESNO," & _
        "  HatAnhaenge YESNO," & _
        "  AnhangAnzahl SHORT DEFAULT 0," & _
        "  MessageClass TEXT(50)," & _
        "  InternetMessageID TEXT(255)," & _
        "  MSGDateiPfad TEXT(255)," & _
        "  Status TEXT(20) DEFAULT 'Neu'," & _
        "  ErstelltAm DATETIME" & _
        ")"

    Debug.Print "  [OK  ] tblEmails"
End Sub

Private Sub Erstelle_tblEmailContent()
    If TabelleExistiert("tblEmailContent") Then
        Debug.Print "  [SKIP] tblEmailContent (existiert bereits)"
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE tblEmailContent (" & _
        "  ContentID AUTOINCREMENT CONSTRAINT PK_Content PRIMARY KEY," & _
        "  EmailID LONG NOT NULL," & _
        "  HTMLBody MEMO," & _
        "  PlainTextBody MEMO," & _
        "  HatHTML YESNO," & _
        "  GroesseHTML LONG DEFAULT 0," & _
        "  GroesseText LONG DEFAULT 0" & _
        ")"

    CurrentDb.Execute "CREATE UNIQUE INDEX idx_Content_EmailID ON tblEmailContent (EmailID)"
    Debug.Print "  [OK  ] tblEmailContent"
End Sub

Private Sub Erstelle_tblEmailEmpfaenger()
    If TabelleExistiert("tblEmailEmpfaenger") Then
        Debug.Print "  [SKIP] tblEmailEmpfaenger (existiert bereits)"
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE tblEmailEmpfaenger (" & _
        "  EmpfaengerID AUTOINCREMENT CONSTRAINT PK_Empfaenger PRIMARY KEY," & _
        "  EmailID LONG NOT NULL," & _
        "  KontaktID LONG DEFAULT 0," & _
        "  Typ TEXT(5)," & _
        "  Anzeigename TEXT(255)," & _
        "  Email TEXT(255)" & _
        ")"

    CurrentDb.Execute "CREATE INDEX idx_Empf_EmailID ON tblEmailEmpfaenger (EmailID)"
    Debug.Print "  [OK  ] tblEmailEmpfaenger"
End Sub

Private Sub Erstelle_tblEmailAnhaenge()
    If TabelleExistiert("tblEmailAnhaenge") Then
        Debug.Print "  [SKIP] tblEmailAnhaenge (existiert bereits)"
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE tblEmailAnhaenge (" & _
        "  AnhangID AUTOINCREMENT CONSTRAINT PK_Anhaenge PRIMARY KEY," & _
        "  EmailID LONG NOT NULL," & _
        "  Dateiname TEXT(255)," & _
        "  DateinameBereinigt TEXT(255)," & _
        "  Erweiterung TEXT(20)," & _
        "  Groesse LONG DEFAULT 0," & _
        "  MimeType TEXT(100)," & _
        "  AnhangTyp SHORT DEFAULT 1," & _
        "  IstVersteckt YESNO," & _
        "  IstGespeichert YESNO," & _
        "  DateiPfad TEXT(255)," & _
        "  ErstelltAm DATETIME" & _
        ")"

    CurrentDb.Execute "CREATE INDEX idx_Anh_EmailID ON tblEmailAnhaenge (EmailID)"
    Debug.Print "  [OK  ] tblEmailAnhaenge"
End Sub

Private Sub Erstelle_tblEmailStatus()
    If TabelleExistiert("tblEmailStatus") Then
        Debug.Print "  [SKIP] tblEmailStatus (existiert bereits)"
        Exit Sub
    End If

    CurrentDb.Execute _
        "CREATE TABLE tblEmailStatus (" & _
        "  StatusID AUTOINCREMENT CONSTRAINT PK_EmailStatus PRIMARY KEY," & _
        "  EmailID LONG NOT NULL," & _
        "  Status TEXT(50)," & _
        "  GeaendertVon TEXT(100)," & _
        "  Bemerkung TEXT(255)," & _
        "  GeaendertAm DATETIME" & _
        ")"

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

    CurrentDb.Execute _
        "CREATE TABLE tblSyncProfil (" & _
        "  ProfilID AUTOINCREMENT CONSTRAINT PK_SyncProfil PRIMARY KEY," & _
        "  ProfilName TEXT(100) NOT NULL," & _
        "  Beschreibung TEXT(255)," & _
        "  IstAktiv YESNO," & _
        "  Projekt TEXT(100)," & _
        "  Phase TEXT(100)," & _
        "  MaxMailsProOrdner LONG DEFAULT 500," & _
        "  MaxTiefe SHORT DEFAULT 5," & _
        "  ExportPfad TEXT(255)," & _
        "  ErstelltAm DATETIME" & _
        ")"

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

    CurrentDb.Execute _
        "CREATE TABLE tblSyncProfilOrdner (" & _
        "  ID AUTOINCREMENT CONSTRAINT PK_SyncProfilOrdner PRIMARY KEY," & _
        "  ProfilID LONG NOT NULL," & _
        "  OrdnerPfad TEXT(255) NOT NULL," & _
        "  PostfachName TEXT(255)," & _
        "  IstAktiv YESNO" & _
        ")"

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
    Call SetzeConfig(db, "ExportBasisPfad", Environ("USERPROFILE") & "\OutlookSync\", _
                    "Basis-Pfad fuer MSG- und Anhang-Export")
    Call SetzeConfig(db, "MaxMailsProSync", "500", _
                    "Maximale Anzahl Mails pro Sync-Durchlauf")
    Call SetzeConfig(db, "AnhaengeExtrahieren", "1", _
                    "Anhaenge auf Festplatte extrahieren (1=Ja / 0=Nein)")
    Call SetzeConfig(db, "MSGExportieren", "1", _
                    "MSG-Dateien exportieren (1=Ja / 0=Nein)")
    Call SetzeConfig(db, "SignaturBilderFiltern", "1", _
                    "Versteckte Signatur-Bilder ueberspringen (1=Ja / 0=Nein)")
    Call SetzeConfig(db, "LogLevel", "3", _
                    "Log-Level (0=Aus 1=Error 2=Warn 3=Info 4=Debug 5=Trace)")
    Call SetzeConfig(db, "SchemaVersion", SCHEMA_VERSION, _
                    "Aktuelle Schema-Version")
    Call SetzeConfig(db, "BackendPfad", "", _
                    "Pfad zur Backend-Datenbank (leer = lokal)")
    Call SetzeConfig(db, "TempPfad", "", _
                    "Temp-Verzeichnis fuer Extraktion (leer = %TEMP%\OutlookSync\)")
    Call SetzeConfig(db, "BufferGroesse", "25", _
                    "Anzahl Mails im Schreib-Puffer vor Flush (5-500)")
    Call SetzeConfig(db, "NetzwerkRetries", "3", _
                    "Anzahl Wiederholungsversuche bei Netzwerkfehlern")
    Call SetzeConfig(db, "NetzwerkRetryPause", "2000", _
                    "Millisekunden Pause zwischen Netzwerk-Retries")
    On Error GoTo 0

    Debug.Print "  [OK  ] Standardkonfiguration gesetzt"
    Set db = Nothing
End Sub

Private Sub SetzeConfig(db As DAO.Database, strKey As String, strVal As String, strDesc As String)
    On Error Resume Next
    Dim lngCount As Long
    lngCount = DCount("*", "tblConfig", "Schluessel='" & strKey & "'")
    If lngCount = 0 Then
        db.Execute "INSERT INTO tblConfig (Schluessel, Wert, Beschreibung) VALUES ('" & _
                   strKey & "', '" & Replace(strVal, "'", "''") & "', '" & Replace(strDesc, "'", "''") & "')"
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
    CurrentDb.Execute "UPDATE tblConfig SET Wert='" & Replace(strVal, "'", "''") & _
                      "' WHERE Schluessel='" & strKey & "'"
    On Error GoTo 0
End Sub
