Option Compare Database
Option Explicit

' ===========================================================================
' modErsteinrichtung - Ersteinrichtungs-Assistent fuer OutlookSync
' ===========================================================================
' Fuehrt die komplette Ersteinrichtung durch:
'   1. Frontend-Tabellen lokal erstellen (tblConfig, tblSyncProfil, ...)
'   2. Backend-DB erstellen (OutlookSync_BE.accdb)
'   3. Backend-Tabellen in der BE-DB anlegen
'   4. Alle BE-Tabellen verknuepfen
'   5. Standard-Konfiguration setzen
'   6. Schema-Integritaet pruefen
'
' AUFRUF (einmalig im Direktbereich oder per Button):
'   Ersteinrichtung                              ' Interaktiver Assistent
'   Ersteinrichtung "\\Server\Share\Daten\"      ' Mit vorgegebenem Pfad
'   ErsteinrichtungLokal                          ' Nur lokaler Modus (kein BE)
'   SchemaIntegritaetPruefen                      ' Nachtraegliche Pruefung
'
' Abhaengigkeiten:
'   modGlobals (Konstanten), modSchema (Tabellen-DDL),
'   modSchemaTools (Smart DDL), modBackend (Verknuepfung),
'   modLogging (Protokoll), modCache (Config-Cache)
' ===========================================================================

Private Const MODUL_NAME As String = "modErsteinrichtung"
Private Const BE_DATEINAME As String = "OutlookSync_BE.accdb"


' ===========================================================================
' HAUPTROUTINE: Interaktiver Ersteinrichtungs-Assistent
' ===========================================================================

' Komplette Ersteinrichtung mit optionalem Backend-Pfad.
' Wenn strBackendPfad leer: Benutzer wird gefragt (InputBox).
' Wenn Benutzer abbricht: Nur lokaler Modus (kein Backend).
'
' Beispiele:
'   Ersteinrichtung                              ? InputBox
'   Ersteinrichtung "\\Server\Share\Daten\"      ? Direkt mit Pfad
'   Ersteinrichtung "S:\Projekte\OutlookSync\"   ? Mapped Drive
Public Sub Ersteinrichtung(Optional ByVal strBackendPfad As String = "")
    On Error GoTo ErrHandler

    Dim dtStart As Double
    dtStart = Timer

    Debug.Print String(70, "=")
    Debug.Print "=== ERSTEINRICHTUNG OutlookSync v" & LeseSchemaVersion() & " ==="
    Debug.Print "=== " & Now() & " ==="
    Debug.Print String(70, "=")
    LogInfo "Ersteinrichtung gestartet", MODUL_NAME

    ' -----------------------------------------------------------------------
    ' SCHRITT 1: Frontend-Tabellen (lokal)
    ' -----------------------------------------------------------------------
    Debug.Print ""
    Debug.Print "--- Schritt 1/6: Frontend-Tabellen erstellen ---"
    Call ErstelleFrontendTabellen
    Debug.Print "    FE-Tabellen: OK"

    ' -----------------------------------------------------------------------
    ' SCHRITT 2: Standard-Konfiguration
    ' -----------------------------------------------------------------------
    Debug.Print ""
    Debug.Print "--- Schritt 2/6: Standard-Konfiguration ---"
    Call InitStandardConfig
    Debug.Print "    Config: OK"

    ' -----------------------------------------------------------------------
    ' SCHRITT 3: Backend-Pfad bestimmen
    ' -----------------------------------------------------------------------
    Debug.Print ""
    Debug.Print "--- Schritt 3/6: Backend-Konfiguration ---"

    If strBackendPfad = "" Then
        strBackendPfad = FrageBackendPfad()
    End If

    If strBackendPfad = "" Then
        ' Benutzer hat abgebrochen -> nur lokaler Modus
        Debug.Print "    LOKALER MODUS (kein Backend)"
        Debug.Print ""
        Debug.Print "--- Schritt 3b: Backend-Tabellen lokal erstellen ---"
        Call ErstelleBackendTabellenLokal
        Debug.Print "    BE-Tabellen lokal: OK"
        GoTo SchemaCheck
    End If

    ' Pfad normalisieren
    strBackendPfad = NormalisiereBEPfad(strBackendPfad)
    Debug.Print "    Backend-Pfad: " & strBackendPfad

    ' -----------------------------------------------------------------------
    ' SCHRITT 4: Backend erstellen + Tabellen anlegen + Verknuepfen
    ' -----------------------------------------------------------------------
    Debug.Print ""
    Debug.Print "--- Schritt 4/6: Backend erstellen + verknuepfen ---"

    If Not ErstelleUndVerknuepfeBackend(strBackendPfad) Then
        ' Fallback: Lokaler Modus
        LogWarn "Backend-Erstellung fehlgeschlagen, Fallback auf lokalen Modus", MODUL_NAME
        Debug.Print "    WARNUNG: Backend fehlgeschlagen -> lokaler Modus"
        Call ErstelleBackendTabellenLokal
        GoTo SchemaCheck
    End If

    Debug.Print "    Backend: OK"

SchemaCheck:
    ' -----------------------------------------------------------------------
    ' SCHRITT 5: Schema-Integritaet pruefen
    ' -----------------------------------------------------------------------
    Debug.Print ""
    Debug.Print "--- Schritt 5/6: Schema-Integritaet pruefen ---"
    Dim lngFehler As Long
    lngFehler = SchemaIntegritaetPruefen()

    ' -----------------------------------------------------------------------
    ' SCHRITT 6: Dual-Access Worker Setup (Queue + optionale Formulare)
    ' -----------------------------------------------------------------------
    Debug.Print ""
    Debug.Print "--- Schritt 6/6: Dual-Access Worker Setup ---"
    If Not SetupDualAccessNoAdmin(False, False) Then
        Debug.Print "    [WARN] Dual-Access Setup unvollstaendig (Basis laeuft weiter)"
    Else
        Debug.Print "    [OK  ] Dual-Access Basis bereit (Formulare optional)"
        Debug.Print "    [HINW] Formulare optional spaeter: ? EnsureDualAccessForms()"
    End If

    Debug.Print ""
    Debug.Print "--- Finaler Schema-Recheck nach Schritt 6 ---"
    lngFehler = SchemaIntegritaetPruefen()

    Dim dblDauer As Double
    dblDauer = Timer - dtStart

    Debug.Print ""
    Debug.Print String(70, "=")
    If lngFehler = 0 Then
        Debug.Print "=== ERSTEINRICHTUNG ERFOLGREICH (" & Format$(dblDauer, "0.0") & "s) ==="
        LogInfo "Ersteinrichtung abgeschlossen (" & Format$(dblDauer, "0.0") & "s)", MODUL_NAME
    Else
        Debug.Print "=== ERSTEINRICHTUNG MIT " & lngFehler & " WARNUNG(EN) (" & Format$(dblDauer, "0.0") & "s) ==="
        LogWarn "Ersteinrichtung mit " & lngFehler & " Warnungen abgeschlossen", MODUL_NAME
    End If
    Debug.Print String(70, "=")

    ' Ergebnis anzeigen
    MsgBox "Ersteinrichtung abgeschlossen!" & vbCrLf & vbCrLf & _
           IIf(CacheGetConfig(CFG_BACKEND_PFAD, "") <> "", _
               "Backend: " & CacheGetConfig(CFG_BACKEND_PFAD, ""), _
               "Modus: Lokal (kein Backend)") & vbCrLf & _
           "Dauer: " & Format$(dblDauer, "0.0") & "s" & vbCrLf & _
           IIf(lngFehler > 0, "Warnungen: " & lngFehler, "Keine Fehler"), _
           vbInformation, "OutlookSync - Ersteinrichtung"

    Exit Sub

ErrHandler:
    LogError "Ersteinrichtung fehlgeschlagen: " & Err.Description, MODUL_NAME
    MsgBox "Ersteinrichtung fehlgeschlagen:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "OutlookSync - Fehler"
End Sub


' ===========================================================================
' LOKALER MODUS: Alles lokal (ohne Backend)
' ===========================================================================

' Ersteinrichtung im reinen Lokal-Modus (kein Backend, kein Netzwerk noetig)
Public Sub ErsteinrichtungLokal()
    On Error GoTo ErrHandler

    Debug.Print String(70, "=")
    Debug.Print "=== ERSTEINRICHTUNG (LOKAL) ==="
    Debug.Print String(70, "=")

    ' Alle 12 Tabellen lokal erstellen
    ErstelleAlleTabellen
    Debug.Print "  Alle Tabellen lokal erstellt."

    ' Config setzen
    InitStandardConfig

    ' Pruefen
    Dim lngFehler As Long
    lngFehler = SchemaIntegritaetPruefen()

    Debug.Print String(70, "=")
    Debug.Print "=== LOKALE ERSTEINRICHTUNG ABGESCHLOSSEN ==="
    Debug.Print String(70, "=")
    Exit Sub

ErrHandler:
    HandleError MODUL_NAME, "ErsteinrichtungLokal"
End Sub


' ===========================================================================
' SCHEMA-INTEGRITAET: Alle Tabellen + Felder + Indizes pruefen
' ===========================================================================

' Prueft ob alle erwarteten Tabellen, Felder und Indizes vorhanden sind.
' Gibt Anzahl Probleme zurueck. Schreibt Ergebnis nach Debug.Print + Log.
Public Function SchemaIntegritaetPruefen() As Long
    On Error GoTo ErrHandler

    Dim lngFehler As Long
    lngFehler = 0

    Debug.Print "  Schema-Pruefung gestartet..."

    ' TableDefs-Cache aktualisieren (nach Verknuepfung noetig)
    CurrentDb.TableDefs.Refresh

    ' --- Alle Tabellen pruefen ---
    Dim arrAlleTabellen As Variant
    arrAlleTabellen = Array(TBL_CONFIG, TBL_SYNC_LAUF, TBL_KONTAKTE, _
                            TBL_OUTLOOK_ORDNER, TBL_EMAIL_THREADS, TBL_EMAILS, _
                            TBL_EMAIL_CONTENT, TBL_EMAIL_EMPFAENGER, TBL_EMAIL_ANHAENGE, _
                            TBL_EMAIL_STATUS, TBL_SYNC_PROFIL, TBL_SYNC_PROFIL_ORDNER, _
                            TBL_PROJEKTE, TBL_EMAIL_PROJEKT, _
                            TBL_SYNC_JOB, TBL_SYNC_HEARTBEAT, TBL_SYNC_CONTROL, TBL_WORKER_LEASE)

    Dim i As Long
    For i = LBound(arrAlleTabellen) To UBound(arrAlleTabellen)
        If TabelleExistiert(CStr(arrAlleTabellen(i))) Then
            Debug.Print "    [OK  ] " & arrAlleTabellen(i) & " " & TabellenKontextText(CStr(arrAlleTabellen(i)))
        Else
            Debug.Print "    [FAIL] " & arrAlleTabellen(i) & " FEHLT! " & TabellenSollKontextText(CStr(arrAlleTabellen(i)))
            LogWarn "Schema-Pruefung: Tabelle fehlt: " & arrAlleTabellen(i), MODUL_NAME
            lngFehler = lngFehler + 1
        End If
    Next i

    ' --- Kritische Felder in Haupttabellen pruefen ---
    lngFehler = lngFehler + PruefeFelder(TBL_EMAILS, Array( _
        "EmailID", "OutlookEntryID", "UniqueHash", "ThreadID", "OrdnerID", _
        "KontaktID_Absender", "SyncLaufID", "Betreff", "AbsenderEmail", _
        "EmpfangenAm", "Status", "ErstelltAm"))

    lngFehler = lngFehler + PruefeFelder(TBL_SYNC_LAUF, Array( _
        "SyncLaufID", "StartZeit", "Status", "AnzahlGelesen", "AnzahlNeu"))

    lngFehler = lngFehler + PruefeFelder(TBL_KONTAKTE, Array( _
        "KontaktID", "Email", "Anzeigename"))

    lngFehler = lngFehler + PruefeFelder(TBL_CONFIG, Array( _
        "ConfigID", "Schluessel", "Wert"))

    ' --- Kritische Indizes pruefen ---
    lngFehler = lngFehler + PruefeIndex(TBL_EMAILS, "idx_Email_Hash")
    lngFehler = lngFehler + PruefeIndex(TBL_CONFIG, "idx_Config_Key")
    lngFehler = lngFehler + PruefeIndex(TBL_KONTAKTE, "idx_Kontakte_Email")
    lngFehler = lngFehler + PruefeIndex(TBL_EMAIL_THREADS, "idx_Thread_Ident")
    lngFehler = lngFehler + PruefeIndex(TBL_OUTLOOK_ORDNER, "idx_Ordner_Pfad")
    lngFehler = lngFehler + PruefeIndex(TBL_SYNC_JOB, "idx_SyncJob_Status")
    lngFehler = lngFehler + PruefeIndex(TBL_SYNC_JOB, "idx_SyncJob_CreatedAt")
    lngFehler = lngFehler + PruefeIndex(TBL_SYNC_HEARTBEAT, "idx_SyncHB_UpdatedAt")
    lngFehler = lngFehler + PruefeIndex(TBL_SYNC_HEARTBEAT, "idx_SyncHB_JobID")

    ' --- Neue Tabellen (v0.6) ---
    lngFehler = lngFehler + PruefeFelder(TBL_PROJEKTE, Array( _
        "ProjektID", "Name", "Status", "ErstelltAm"))
    lngFehler = lngFehler + PruefeFelder(TBL_EMAIL_PROJEKT, Array( _
        "EmailProjektID", "EmailID", "ProjektID", "ZugeordnetAm"))
    lngFehler = lngFehler + PruefeIndex(TBL_PROJEKTE, "idx_Projekt_Name")
    lngFehler = lngFehler + PruefeIndex(TBL_EMAIL_PROJEKT, "idx_EP_Unique")

    ' --- Backend-Verknuepfung pruefen (wenn konfiguriert) ---
    Dim strBEPfad As String
    strBEPfad = CacheGetConfig(CFG_BACKEND_PFAD, "")
    If strBEPfad <> "" Then
        lngFehler = lngFehler + PruefeBackendLinks(strBEPfad)
    End If

    ' Zusammenfassung
    If lngFehler = 0 Then
        Debug.Print "  Schema-Pruefung: ALLES OK (18 Tabellen, Felder, Indizes)"
    Else
        Debug.Print "  Schema-Pruefung: " & lngFehler & " Problem(e) gefunden!"
    End If

    SchemaIntegritaetPruefen = lngFehler
    Exit Function

ErrHandler:
    HandleError MODUL_NAME, "SchemaIntegritaetPruefen"
    SchemaIntegritaetPruefen = -1
End Function


' ===========================================================================
' SCHEMA-MIGRATION: Bestehendes Schema auf neue Version bringen
' ===========================================================================

' Fuehrt Schema-Migrationen durch (fehlende Spalten/Indizes nachruesten).
' Kann gefahrlos wiederholt aufgerufen werden (idempotent).
'
' Rufe auf nach einem Update, um neue Felder/Indizes hinzuzufuegen,
' OHNE bestehende Daten zu verlieren.
Public Sub SchemaMigration()
    On Error GoTo ErrHandler

    Debug.Print String(70, "=")
    Debug.Print "=== SCHEMA-MIGRATION ==="
    Debug.Print String(70, "=")

    ' --- tblEmails: Sicherstellen dass alle Felder da sind ---
    EnsureColumn TBL_EMAILS, "MSGDateiPfad", "TEXT(255)"
    EnsureColumn TBL_EMAILS, "InternetMessageID", "TEXT(255)"
    EnsureColumn TBL_EMAILS, "MessageClass", "TEXT(50)"

    ' --- tblSyncLauf: Projekt/Phase (ab v0.5) ---
    EnsureColumn TBL_SYNC_LAUF, "Projekt", "TEXT(100)"
    EnsureColumn TBL_SYNC_LAUF, "Phase", "TEXT(100)"

    ' --- tblKontakte: Erweiterte Felder ---
    EnsureColumn TBL_KONTAKTE, "Titel", "TEXT(50)"
    EnsureColumn TBL_KONTAKTE, "Namenszusatz", "TEXT(100)"
    EnsureColumn TBL_KONTAKTE, "Institution", "TEXT(255)"
    EnsureColumn TBL_KONTAKTE, "Sortiername", "TEXT(255)"

    ' --- Performance-Indizes nachziehen ---
    EnsureIndex TBL_EMAILS, "idx_Email_SyncLauf", "SyncLaufID"
    EnsureIndex TBL_EMAILS, "idx_Email_Datum", "EmpfangenAm"
    EnsureIndex TBL_EMAILS, "idx_Email_OrdnerID", "OrdnerID"

    ' --- Defaults sicherstellen ---
    EnsureDefaults TBL_EMAILS, Array( _
        "ThreadID|0", "OrdnerID|0", "KontaktID_Absender|0", _
        "SyncLaufID|0", "Groesse|0", "Wichtigkeit|1", _
        "AnhangAnzahl|0", "Status|'Neu'")

    EnsureDefaults TBL_SYNC_LAUF, Array( _
        "Status|'Gestartet'", "AnzahlGelesen|0", "AnzahlNeu|0", _
        "AnzahlDuplikate|0", "AnzahlFehler|0")

    ' =================================================================
    ' v0.6: Projekt-Tagging + Dedup-Verbesserung
    ' =================================================================

    ' --- Neue Tabellen sicherstellen (idempotent) ---
    Debug.Print "  v0.6: Projekt-Tabellen..."
    If Not TabelleExistiert(TBL_PROJEKTE) Then
        Erstelle_tblProjekte_Migration
    End If
    If Not TabelleExistiert(TBL_EMAIL_PROJEKT) Then
        Erstelle_tblEmailProjekt_Migration
    End If

    ' --- tblSyncProfil: ProjektID-Feld ---
    EnsureColumn TBL_SYNC_PROFIL, "ProjektID", "LONG"
    EnsureDefaults TBL_SYNC_PROFIL, Array("ProjektID|0")
    EnsureIndex TBL_SYNC_PROFIL, "idx_Profil_ProjektID", "ProjektID"

    ' --- tblEmails: InternetMessageID-Index fuer Dedup ---
    EnsureIndex TBL_EMAILS, "idx_Email_InternetMsgID", "InternetMessageID"

    ' --- Projekt-Indizes sicherstellen ---
    EnsureUniqueIndex TBL_PROJEKTE, "idx_Projekt_Name", "Name"
    EnsureIndex TBL_PROJEKTE, "idx_Projekt_Status", "Status"
    EnsureIndex TBL_EMAIL_PROJEKT, "idx_EP_EmailID", "EmailID"
    EnsureIndex TBL_EMAIL_PROJEKT, "idx_EP_ProjektID", "ProjektID"
    EnsureUniqueIndex TBL_EMAIL_PROJEKT, "idx_EP_Unique", "EmailID, ProjektID"

    ' --- Daten-Migration (idempotent) ---
    Debug.Print "  v0.6: Daten-Migration..."
    MigriereProjektStammdaten
    MigriereEmailProjektZuordnung
    MigriereSyncProfilProjektID

    ' =================================================================
    ' v0.7: Dual-Access Queue/Worker
    ' =================================================================
    Debug.Print "  v0.7: Worker-Tabellen..."
    If Not TabelleExistiert(TBL_SYNC_JOB) Then
        Erstelle_tblSyncJob_Migration
    End If
    If Not TabelleExistiert(TBL_SYNC_HEARTBEAT) Then
        Erstelle_tblSyncHeartbeat_Migration
    End If
    If Not TabelleExistiert(TBL_SYNC_CONTROL) Then
        Erstelle_tblSyncControl_Migration
    End If
    If Not TabelleExistiert(TBL_WORKER_LEASE) Then
        Erstelle_tblWorkerLease_Migration
    End If

    EnsureIndex TBL_SYNC_JOB, "idx_SyncJob_Status", "Status"
    EnsureIndex TBL_SYNC_JOB, "idx_SyncJob_CreatedAt", "CreatedAt"
    EnsureIndex TBL_SYNC_HEARTBEAT, "idx_SyncHB_UpdatedAt", "UpdatedAt"
    EnsureIndex TBL_SYNC_HEARTBEAT, "idx_SyncHB_JobID", "JobID"
    EnsureIndex TBL_WORKER_LEASE, "idx_WorkerLease_Until", "LeaseUntil"

    Debug.Print "  Migration abgeschlossen."
    Debug.Print String(70, "=")
    LogInfo "Schema-Migration abgeschlossen", MODUL_NAME
    Exit Sub

ErrHandler:
    HandleError MODUL_NAME, "SchemaMigration"
End Sub


' ===========================================================================
' STATUS-REPORT: Uebersicht ueber aktuellen Zustand
' ===========================================================================

' Gibt eine vollstaendige Statusuebersicht aus (Debug.Print)
Public Sub StatusReport()
    On Error Resume Next

    Debug.Print String(70, "=")
    Debug.Print "=== OUTLOOKSYNC STATUS-REPORT ==="
    Debug.Print "=== " & Now() & " ==="
    Debug.Print String(70, "=")

    ' 1. Backend-Status
    Debug.Print ""
    Debug.Print "--- Backend ---"
    Debug.Print "  Status:  " & BackendStatus()
    Debug.Print "  Pfad:    " & IIf(GetBackendPfad() = "", "(lokal)", GetBackendPfad())
    Debug.Print "  Offline: " & IIf(g_blnBackendOffline, "JA (!)", "Nein")

    ' 2. Tabellen-Status
    Debug.Print ""
    Debug.Print "--- Tabellen ---"
    Dim arrAlle As Variant
    arrAlle = Array(TBL_CONFIG, TBL_SYNC_LAUF, TBL_KONTAKTE, _
                    TBL_OUTLOOK_ORDNER, TBL_EMAIL_THREADS, TBL_EMAILS, _
                    TBL_EMAIL_CONTENT, TBL_EMAIL_EMPFAENGER, TBL_EMAIL_ANHAENGE, _
                    TBL_EMAIL_STATUS, TBL_SYNC_PROFIL, TBL_SYNC_PROFIL_ORDNER, _
                    TBL_PROJEKTE, TBL_EMAIL_PROJEKT, TBL_SYNC_JOB, _
                    TBL_SYNC_HEARTBEAT, TBL_SYNC_CONTROL, TBL_WORKER_LEASE)

    Dim i As Long
    For i = LBound(arrAlle) To UBound(arrAlle)
        Dim strStatus As String
        Dim strTbl As String
        strTbl = CStr(arrAlle(i))

        If Not TabelleExistiert(strTbl) Then
            strStatus = "FEHLT"
        ElseIf IstLinkedTable(strTbl) Then
            strStatus = "BE-Link (" & AnzahlFelder(strTbl) & " Felder)"
        Else
            strStatus = "Lokal (" & AnzahlFelder(strTbl) & " Felder)"
        End If

        Debug.Print "  " & Left$(strTbl & String(28, " "), 28) & strStatus
    Next i

    ' 3. Datensatz-Statistik
    Debug.Print ""
    Debug.Print "--- Datensaetze ---"
    Dim lngCount As Long
    Dim arrStats As Variant
    arrStats = Array(TBL_EMAILS, TBL_KONTAKTE, TBL_EMAIL_THREADS, _
                     TBL_SYNC_LAUF, TBL_OUTLOOK_ORDNER, TBL_EMAIL_ANHAENGE)
    For i = LBound(arrStats) To UBound(arrStats)
        lngCount = 0
        On Error Resume Next
        lngCount = DCount("*", CStr(arrStats(i)))
        On Error GoTo 0
        Debug.Print "  " & Left$(CStr(arrStats(i)) & String(28, " "), 28) & Format$(lngCount, "#,##0")
    Next i

    ' 4. Config-Auszug
    Debug.Print ""
    Debug.Print "--- Konfiguration ---"
    Debug.Print "  Schema-Version: " & LeseConfig(CFG_SCHEMA_VERSION, "?")
    Debug.Print "  Log-Level:      " & LeseConfig(CFG_LOG_LEVEL, "?")
    Debug.Print "  Export-Pfad:    " & LeseConfig(CFG_EXPORT_PFAD, "(nicht gesetzt)")
    Debug.Print "  Buffer-Groesse: " & LeseConfig(CFG_BUFFER_GROESSE, "?")
    Debug.Print "  Max Mails:      " & LeseConfig(CFG_MAX_MAILS, "?")

    Debug.Print ""
    Debug.Print String(70, "=")
    On Error GoTo 0
End Sub


' ===========================================================================
' PRIVATE: Einzelschritte der Ersteinrichtung
' ===========================================================================

' Erstellt nur die Frontend-Tabellen (tblConfig, tblSyncProfil, tblSyncProfilOrdner)
Private Sub ErstelleFrontendTabellen()
    Dim db As DAO.Database
    Set db = CurrentDb

    ' tblConfig
    If Not TabelleExistiert(TBL_CONFIG) Then
        db.Execute "CREATE TABLE [" & TBL_CONFIG & "] (" & _
                   "ConfigID AUTOINCREMENT CONSTRAINT PK_Config PRIMARY KEY, " & _
                   "Schluessel TEXT(100) NOT NULL, " & _
                   "Wert TEXT(255), " & _
                   "Beschreibung TEXT(255))"
        db.Execute "CREATE UNIQUE INDEX idx_Config_Key ON [" & TBL_CONFIG & "] (Schluessel)"
        Debug.Print "    [OK  ] " & TBL_CONFIG
    Else
        Debug.Print "    [SKIP] " & TBL_CONFIG & " (existiert)"
    End If

    ' tblSyncProfil
    If Not TabelleExistiert(TBL_SYNC_PROFIL) Then
        db.Execute "CREATE TABLE [" & TBL_SYNC_PROFIL & "] (" & _
                   "ProfilID AUTOINCREMENT CONSTRAINT PK_SyncProfil PRIMARY KEY, " & _
                   "ProfilName TEXT(100) NOT NULL, " & _
                   "Beschreibung TEXT(255), " & _
                   "IstAktiv YESNO, " & _
                   "Projekt TEXT(100), " & _
                   "Phase TEXT(100), " & _
                   "MaxMailsProOrdner LONG, " & _
                   "MaxTiefe SHORT, " & _
                   "ExportPfad TEXT(255), " & _
                   "ErstelltAm DATETIME)"
        db.TableDefs(TBL_SYNC_PROFIL).Fields("MaxMailsProOrdner").DefaultValue = "500"
        db.TableDefs(TBL_SYNC_PROFIL).Fields("MaxTiefe").DefaultValue = "5"
        db.Execute "CREATE UNIQUE INDEX idx_Profil_Name ON [" & TBL_SYNC_PROFIL & "] (ProfilName)"
        Debug.Print "    [OK  ] " & TBL_SYNC_PROFIL
    Else
        Debug.Print "    [SKIP] " & TBL_SYNC_PROFIL & " (existiert)"
    End If

    ' tblSyncProfilOrdner
    If Not TabelleExistiert(TBL_SYNC_PROFIL_ORDNER) Then
        db.Execute "CREATE TABLE [" & TBL_SYNC_PROFIL_ORDNER & "] (" & _
                   "ID AUTOINCREMENT CONSTRAINT PK_SyncProfilOrdner PRIMARY KEY, " & _
                   "ProfilID LONG NOT NULL, " & _
                   "OrdnerPfad TEXT(255) NOT NULL, " & _
                   "PostfachName TEXT(255), " & _
                   "IstAktiv YESNO)"
        db.Execute "CREATE INDEX idx_ProfilOrdner_Profil ON [" & TBL_SYNC_PROFIL_ORDNER & "] (ProfilID)"
        Debug.Print "    [OK  ] " & TBL_SYNC_PROFIL_ORDNER
    Else
        Debug.Print "    [SKIP] " & TBL_SYNC_PROFIL_ORDNER & " (existiert)"
    End If

    ' Indizes sicherstellen (auch fuer bereits existierende Tabellen ohne Index)
    Debug.Print "    FE-Indizes sicherstellen..."
    On Error Resume Next
    db.Execute "CREATE UNIQUE INDEX idx_Config_Key ON [" & TBL_CONFIG & "] (Schluessel)"
    If Err.Number <> 0 Then Debug.Print "    idx_Config_Key: " & Err.Description: Err.Clear
    db.Execute "CREATE UNIQUE INDEX idx_Profil_Name ON [" & TBL_SYNC_PROFIL & "] (ProfilName)"
    If Err.Number <> 0 Then Debug.Print "    idx_Profil_Name: " & Err.Description: Err.Clear
    db.Execute "CREATE INDEX idx_ProfilOrdner_Profil ON [" & TBL_SYNC_PROFIL_ORDNER & "] (ProfilID)"
    If Err.Number <> 0 Then Debug.Print "    idx_ProfilOrdner: " & Err.Description: Err.Clear
    On Error GoTo 0
    Debug.Print "    FE-Indizes: OK"

    Set db = Nothing
End Sub


' Erstellt Backend-Tabellen LOKAL (wenn kein Backend konfiguriert)
' ErstelleBackendTabellenInDB ist idempotent: ueberspringt vorhandene Tabellen,
' erstellt fehlende, und stellt alle Indizes sicher.
Private Sub ErstelleBackendTabellenLokal()
    Dim db As DAO.Database
    Set db = CurrentDb

    If Not ErstelleBackendTabellenInDB(db) Then
        LogWarn "Lokale BE-Tabellen Erstellung hatte Probleme", MODUL_NAME
    End If

    Set db = Nothing
End Sub


' Erstellt Backend-DB + Tabellen + Verknuepfung (der volle Flow)
Private Function ErstelleUndVerknuepfeBackend(ByVal strPfad As String) As Boolean
    On Error GoTo ErrHandler

    ' 1. Verzeichnis erstellen
    Dim strDir As String
    strDir = Left$(strPfad, InStrRev(strPfad, "\"))
    If strDir <> "" Then
        ErstelleOrdner strDir
    End If

    ' 2. Backend-DB erstellen (wenn noetig)
    Dim dbBE As DAO.Database
    If Dir(strPfad) = "" Then
        Debug.Print "    Backend-DB wird erstellt: " & strPfad
        Set dbBE = DBEngine.CreateDatabase(strPfad, dbLangGeneral)
        LogInfo "Backend-DB erstellt: " & strPfad, MODUL_NAME
    Else
        Debug.Print "    [SKIP] Backend-DB existiert bereits"
        Set dbBE = DBEngine.OpenDatabase(strPfad)
    End If

    ' 3. Tabellen in Backend-DB sicherstellen (idempotent)
    Debug.Print "    Backend-Tabellen sicherstellen..."
    If Not ErstelleBackendTabellenInDB(dbBE) Then
        dbBE.Close: Set dbBE = Nothing
        LogError "Backend-Tabellen Erstellung fehlgeschlagen", MODUL_NAME
        ErstelleUndVerknuepfeBackend = False
        Exit Function
    End If
    dbBE.Close: Set dbBE = Nothing
    Debug.Print "    [OK  ] Backend-Tabellen komplett"

    ' 4. TableDefs-Cache aktualisieren (damit SchemaIntegritaetPruefen Links sieht)
    CurrentDb.TableDefs.Refresh

    ' 5. Tabellen verknuepfen (mit Datenmigration)
    If Not VerknuepfeBackend(strPfad) Then
        LogError "Backend-Verknuepfung fehlgeschlagen", MODUL_NAME
        ErstelleUndVerknuepfeBackend = False
        Exit Function
    End If

    ErstelleUndVerknuepfeBackend = True
    Exit Function

ErrHandler:
    Debug.Print "    [ERROR] ErstelleUndVerknuepfeBackend: " & Err.Number & " - " & Err.Description
    LogError "ErstelleUndVerknuepfeBackend: " & Err.Number & " - " & Err.Description, MODUL_NAME
    On Error Resume Next
    If Not dbBE Is Nothing Then dbBE.Close: Set dbBE = Nothing
    On Error GoTo 0
    ErstelleUndVerknuepfeBackend = False
End Function


' ===========================================================================
' PRIVATE: Benutzer-Interaktion
' ===========================================================================

' Fragt den Benutzer nach dem Backend-Pfad (InputBox + FileDialog Fallback)
' Gibt "" zurueck wenn abgebrochen
Private Function FrageBackendPfad() As String
    Dim strAntwort As String

    ' Erst fragen ob ueberhaupt Backend gewuenscht
    Dim intAntwort As VbMsgBoxResult
    intAntwort = MsgBox( _
        "Soll eine Backend-Datenbank auf einem Netzlaufwerk eingerichtet werden?" & vbCrLf & vbCrLf & _
        "JA = Backend auf Netzlaufwerk (Mehrbenutzerbetrieb)" & vbCrLf & _
        "NEIN = Nur lokal arbeiten (Einzelplatz)" & vbCrLf & vbCrLf & _
        "Das Backend kann auch spaeter mit" & vbCrLf & _
        "  VerknuepfeBackend ""\\Server\Pfad\""" & vbCrLf & _
        "eingerichtet werden.", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "OutlookSync - Backend einrichten?")

    If intAntwort = vbNo Then
        FrageBackendPfad = ""
        Exit Function
    End If

    ' Pfad per InputBox abfragen
    strAntwort = InputBox( _
        "Bitte den Pfad fuer die Backend-Datenbank eingeben:" & vbCrLf & vbCrLf & _
        "Beispiele:" & vbCrLf & _
        "  \\Server\Share\OutlookSync\" & vbCrLf & _
        "  S:\Projekte\OutlookSync\" & vbCrLf & vbCrLf & _
        "Der Dateiname '" & BE_DATEINAME & "' wird automatisch angehaengt.", _
        "OutlookSync - Backend-Pfad", _
        Environ("USERPROFILE") & "\OutlookSync\")

    If Trim$(strAntwort) = "" Then
        FrageBackendPfad = ""
    Else
        FrageBackendPfad = Trim$(strAntwort)
    End If
End Function


' ===========================================================================
' PRIVATE: Hilfsfunktionen
' ===========================================================================

' Normalisiert den Backend-Pfad (haengt Dateiname an wenn noetig)
Private Function NormalisiereBEPfad(ByVal strPfad As String) As String
    strPfad = Trim$(strPfad)
    If strPfad = "" Then NormalisiereBEPfad = "": Exit Function

    ' Backslash am Ende sicherstellen fuer Verzeichnisse
    If Right$(strPfad, 6) <> ".accdb" Then
        If Right$(strPfad, 1) <> "\" Then strPfad = strPfad & "\"
        strPfad = strPfad & BE_DATEINAME
    End If

    NormalisiereBEPfad = strPfad
End Function


' Liest die aktuelle Schema-Version aus modSchema
Private Function LeseSchemaVersion() As String
    On Error Resume Next
    LeseSchemaVersion = LeseConfig(CFG_SCHEMA_VERSION, "0.5.2")
    On Error GoTo 0
End Function


' Prueft ob bestimmte Felder in einer Tabelle existieren
' Gibt Anzahl fehlender Felder zurueck
Private Function PruefeFelder(ByVal strTabelle As String, arrFelder As Variant) As Long
    If Not TabelleExistiert(strTabelle) Then
        PruefeFelder = 0  ' Tabelle fehlt insgesamt -> wurde schon oben gezaehlt
        Exit Function
    End If

    Dim lngFehler As Long
    lngFehler = 0

    Dim i As Long
    For i = LBound(arrFelder) To UBound(arrFelder)
        If Not FeldExistiert(strTabelle, CStr(arrFelder(i))) Then
            Debug.Print "    [WARN] " & strTabelle & "." & arrFelder(i) & " FEHLT! " & TabellenKontextText(strTabelle)
            lngFehler = lngFehler + 1
        End If
    Next i

    PruefeFelder = lngFehler
End Function


' Prueft ob ein bestimmter Index existiert
' Gibt 0 oder 1 zurueck
Private Function PruefeIndex(ByVal strTabelle As String, ByVal strIndex As String) As Long
    If Not TabelleExistiert(strTabelle) Then
        PruefeIndex = 0
        Exit Function
    End If

    If IndexExistiert(strTabelle, strIndex) Then
        PruefeIndex = 0
    Else
        Debug.Print "    [WARN] Index fehlt: " & strIndex & " ON " & strTabelle & " " & TabellenKontextText(strTabelle)
        PruefeIndex = 1
    End If
End Function

Private Function TabellenKontextText(ByVal strTabelle As String) As String
    If Not TabelleExistiert(strTabelle) Then
        TabellenKontextText = TabellenSollKontextText(strTabelle)
    ElseIf IstLinkedTable(strTabelle) Then
        TabellenKontextText = "(BE-Link)"
    Else
        TabellenKontextText = "(FE/Lokal)"
    End If
End Function

Private Function TabellenSollKontextText(ByVal strTabelle As String) As String
    Select Case strTabelle
        Case TBL_CONFIG, TBL_SYNC_PROFIL, TBL_SYNC_PROFIL_ORDNER
            TabellenSollKontextText = "(erwartet FE/Lokal)"
        Case Else
            TabellenSollKontextText = "(erwartet BE/Link)"
    End Select
End Function


' Prueft ob alle Backend-Tabellen korrekt verlinkt sind
Private Function PruefeBackendLinks(ByVal strBEPfad As String) As Long
    Dim lngFehler As Long
    lngFehler = 0

    Dim arrBE As Variant
    arrBE = Array(TBL_SYNC_LAUF, TBL_KONTAKTE, TBL_OUTLOOK_ORDNER, _
                  TBL_EMAIL_THREADS, TBL_EMAILS, TBL_EMAIL_CONTENT, _
                  TBL_EMAIL_EMPFAENGER, TBL_EMAIL_ANHAENGE, TBL_EMAIL_STATUS, _
                  TBL_PROJEKTE, TBL_EMAIL_PROJEKT, _
                  TBL_SYNC_JOB, TBL_SYNC_HEARTBEAT, TBL_SYNC_CONTROL, TBL_WORKER_LEASE)

    Dim i As Long
    For i = LBound(arrBE) To UBound(arrBE)
        If TabelleExistiert(CStr(arrBE(i))) Then
            If Not IstLinkedTable(CStr(arrBE(i))) Then
                Debug.Print "    [WARN] " & arrBE(i) & " sollte verlinkt sein, ist aber lokal!"
                lngFehler = lngFehler + 1
            End If
        End If
    Next i

    PruefeBackendLinks = lngFehler
End Function


' ===================================================================
' v0.6 MIGRATION: Projekt-Tabellen + Daten-Migration
' ===================================================================

' Erstellt tblProjekte via RunDDL_Smart (funktioniert lokal + Backend)
Private Sub Erstelle_tblProjekte_Migration()
    On Error GoTo ErrHandler
    Dim sql As String
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
    RunDDL_Smart TBL_PROJEKTE, sql
    SetzeDefaultSmart TBL_PROJEKTE, "Status", "'Aktiv'"
    SetzeDefaultSmart TBL_PROJEKTE, "SortierNr", "0"
    Debug.Print "    [OK  ] " & TBL_PROJEKTE & " erstellt (Migration)"
    Exit Sub
ErrHandler:
    Debug.Print "    [FAIL] " & TBL_PROJEKTE & " Migration: " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub


' Erstellt tblEmailProjekt via RunDDL_Smart
Private Sub Erstelle_tblEmailProjekt_Migration()
    On Error GoTo ErrHandler
    Dim sql As String
    sql = "CREATE TABLE [" & TBL_EMAIL_PROJEKT & "] (" & _
          "EmailProjektID AUTOINCREMENT CONSTRAINT PK_EmailProjekt PRIMARY KEY, " & _
          "EmailID LONG NOT NULL, " & _
          "ProjektID LONG NOT NULL, " & _
          "Quelle TEXT(20), " & _
          "ZugeordnetVon TEXT(100), " & _
          "ZugeordnetAm DATETIME)"
    RunDDL_Smart TBL_EMAIL_PROJEKT, sql
    SetzeDefaultSmart TBL_EMAIL_PROJEKT, "Quelle", "'Manuell'"
    Debug.Print "    [OK  ] " & TBL_EMAIL_PROJEKT & " erstellt (Migration)"
    Exit Sub
ErrHandler:
    Debug.Print "    [FAIL] " & TBL_EMAIL_PROJEKT & " Migration: " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub


' Erstellt tblSyncJob via RunDDL_Smart
Private Sub Erstelle_tblSyncJob_Migration()
    On Error GoTo ErrHandler
    Dim sql As String
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
    RunDDL_Smart TBL_SYNC_JOB, sql
    SetzeDefaultSmart TBL_SYNC_JOB, "Status", "'" & JOB_STATUS_QUEUED & "'"
    SetzeDefaultSmart TBL_SYNC_JOB, "Priority", "100"
    SetzeDefaultSmart TBL_SYNC_JOB, "RequestedMaxMails", "500"
    SetzeDefaultSmart TBL_SYNC_JOB, "RequestedSubfolders", "False"
    Debug.Print "    [OK  ] " & TBL_SYNC_JOB & " erstellt (Migration)"
    Exit Sub
ErrHandler:
    Debug.Print "    [FAIL] " & TBL_SYNC_JOB & " Migration: " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub


' Erstellt tblSyncHeartbeat via RunDDL_Smart
Private Sub Erstelle_tblSyncHeartbeat_Migration()
    On Error GoTo ErrHandler
    Dim sql As String
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
    RunDDL_Smart TBL_SYNC_HEARTBEAT, sql
    SetzeDefaultSmart TBL_SYNC_HEARTBEAT, "CurrentItem", "0"
    SetzeDefaultSmart TBL_SYNC_HEARTBEAT, "TotalItems", "0"
    SetzeDefaultSmart TBL_SYNC_HEARTBEAT, "COMRetries", "0"
    SetzeDefaultSmart TBL_SYNC_HEARTBEAT, "COMReconnects", "0"
    Debug.Print "    [OK  ] " & TBL_SYNC_HEARTBEAT & " erstellt (Migration)"
    Exit Sub
ErrHandler:
    Debug.Print "    [FAIL] " & TBL_SYNC_HEARTBEAT & " Migration: " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub


' Erstellt tblSyncControl via RunDDL_Smart
Private Sub Erstelle_tblSyncControl_Migration()
    On Error GoTo ErrHandler
    Dim sql As String
    sql = "CREATE TABLE [" & TBL_SYNC_CONTROL & "] (" & _
          "JobID LONG CONSTRAINT PK_SyncControl PRIMARY KEY, " & _
          "PauseRequested YESNO, " & _
          "CancelRequested YESNO, " & _
          "UpdatedAt DATETIME)"
    RunDDL_Smart TBL_SYNC_CONTROL, sql
    SetzeDefaultSmart TBL_SYNC_CONTROL, "PauseRequested", "False"
    SetzeDefaultSmart TBL_SYNC_CONTROL, "CancelRequested", "False"
    Debug.Print "    [OK  ] " & TBL_SYNC_CONTROL & " erstellt (Migration)"
    Exit Sub
ErrHandler:
    Debug.Print "    [FAIL] " & TBL_SYNC_CONTROL & " Migration: " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub


' Erstellt tblWorkerLease via RunDDL_Smart
Private Sub Erstelle_tblWorkerLease_Migration()
    On Error GoTo ErrHandler
    Dim sql As String
    sql = "CREATE TABLE [" & TBL_WORKER_LEASE & "] (" & _
          "WorkerId TEXT(100) CONSTRAINT PK_WorkerLease PRIMARY KEY, " & _
          "LeaseUntil DATETIME, " & _
          "UpdatedAt DATETIME, " & _
          "HostName TEXT(100), " & _
          "SessionUser TEXT(100))"
    RunDDL_Smart TBL_WORKER_LEASE, sql
    Debug.Print "    [OK  ] " & TBL_WORKER_LEASE & " erstellt (Migration)"
    Exit Sub
ErrHandler:
    Debug.Print "    [FAIL] " & TBL_WORKER_LEASE & " Migration: " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub


' Befuellt tblProjekte aus existierenden Freitext-Projekt-Werten
Private Sub MigriereProjektStammdaten()
    On Error GoTo ErrHandler
    If Not TabelleExistiert(TBL_PROJEKTE) Then Exit Sub

    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb
    Dim lngCount As Long: lngCount = 0

    ' 1. Aus tblSyncProfil (Frontend-Tabelle)
    If TabelleExistiert(TBL_SYNC_PROFIL) Then
        Set rs = db.OpenRecordset( _
            "SELECT DISTINCT Projekt FROM [" & TBL_SYNC_PROFIL & "] " & _
            "WHERE Projekt IS NOT NULL AND Projekt <> ''", dbOpenSnapshot)
        Do While Not rs.EOF
            Dim strP As String
            strP = Nz(rs!Projekt, "")
            If strP <> "" And HoleProjektID(strP) = 0 Then
                ErstelleProjekt strP, , "Migriert aus SyncProfil"
                lngCount = lngCount + 1
            End If
            rs.MoveNext
        Loop
        rs.Close: Set rs = Nothing
    End If

    ' 2. Aus tblSyncLauf (Backend-Tabelle, kann weitere Werte haben)
    If TabelleExistiert(TBL_SYNC_LAUF) Then
        Set rs = db.OpenRecordset( _
            "SELECT DISTINCT Projekt FROM [" & TBL_SYNC_LAUF & "] " & _
            "WHERE Projekt IS NOT NULL AND Projekt <> ''", dbOpenSnapshot)
        Do While Not rs.EOF
            strP = Nz(rs!Projekt, "")
            If strP <> "" And HoleProjektID(strP) = 0 Then
                ErstelleProjekt strP, , "Migriert aus SyncLauf"
                lngCount = lngCount + 1
            End If
            rs.MoveNext
        Loop
        rs.Close: Set rs = Nothing
    End If

    Set db = Nothing
    If lngCount > 0 Then
        Debug.Print "    " & lngCount & " Projekte aus Bestandsdaten migriert"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "    [WARN] MigriereProjektStammdaten: " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub


' Befuellt tblEmailProjekt: Verknuepft Emails mit Projekten via SyncLauf
Private Sub MigriereEmailProjektZuordnung()
    On Error GoTo ErrHandler
    If Not TabelleExistiert(TBL_EMAIL_PROJEKT) Then Exit Sub
    If Not TabelleExistiert(TBL_PROJEKTE) Then Exit Sub
    If Not TabelleExistiert(TBL_EMAILS) Then Exit Sub
    If Not TabelleExistiert(TBL_SYNC_LAUF) Then Exit Sub

    ' Nur Emails die noch KEINE Zuordnung haben
    Dim strSql As String
    strSql = "INSERT INTO [" & TBL_EMAIL_PROJEKT & "] " & _
             "(EmailID, ProjektID, Quelle, ZugeordnetVon, ZugeordnetAm) " & _
             "SELECT e.EmailID, p.ProjektID, '" & EP_QUELLE_MIGRATION & "', 'System', Now() " & _
             "FROM ([" & TBL_EMAILS & "] e " & _
             "INNER JOIN [" & TBL_SYNC_LAUF & "] sl ON e.SyncLaufID = sl.SyncLaufID) " & _
             "INNER JOIN [" & TBL_PROJEKTE & "] p ON sl.Projekt = p.Name " & _
             "WHERE sl.Projekt IS NOT NULL AND sl.Projekt <> '' " & _
             "AND e.EmailID NOT IN (" & _
             "  SELECT ep2.EmailID FROM [" & TBL_EMAIL_PROJEKT & "] ep2 " & _
             "  WHERE ep2.ProjektID = p.ProjektID)"

    Dim db As DAO.Database
    Dim lngTry As Long
    Dim lngMaxTry As Long
    Set db = CurrentDb

    lngMaxTry = 4
    For lngTry = 1 To lngMaxTry
        On Error GoTo WriteErr
        db.Execute strSql, dbFailOnError
        Exit For

WriteErr:
        If IsLockError(Err.Number) And lngTry < lngMaxTry Then
            Err.Clear
            Sleep 200 * lngTry
            DoEvents
        ElseIf Err.Number <> 0 Then
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Next lngTry

    Dim lngInserted As Long
    lngInserted = db.RecordsAffected
    Set db = Nothing

    If lngInserted > 0 Then
        Debug.Print "    " & lngInserted & " Email-Projekt-Zuordnungen migriert"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "    [WARN] MigriereEmailProjektZuordnung: " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub


' Setzt ProjektID in tblSyncProfil basierend auf dem Freitext-Feld Projekt
Private Sub MigriereSyncProfilProjektID()
    On Error GoTo ErrHandler
    If Not TabelleExistiert(TBL_SYNC_PROFIL) Then Exit Sub
    If Not TabelleExistiert(TBL_PROJEKTE) Then Exit Sub
    If Not FeldExistiert(TBL_SYNC_PROFIL, "ProjektID") Then Exit Sub

    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb
    Dim lngCount As Long: lngCount = 0

    Set rs = db.OpenRecordset( _
        "SELECT ProfilID, Projekt, ProjektID FROM [" & TBL_SYNC_PROFIL & "] " & _
        "WHERE Projekt IS NOT NULL AND Projekt <> '' AND (ProjektID = 0 OR ProjektID IS NULL)", _
        dbOpenDynaset)

    Do While Not rs.EOF
        Dim lngPid As Long
        lngPid = HoleProjektID(Nz(rs!Projekt, ""))
        If lngPid > 0 Then
            rs.Edit
            rs!ProjektID = lngPid
            rs.Update
            lngCount = lngCount + 1
        End If
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    Set db = Nothing

    If lngCount > 0 Then
        Debug.Print "    " & lngCount & " SyncProfil-ProjektIDs gesetzt"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "    [WARN] MigriereSyncProfilProjektID: " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub


