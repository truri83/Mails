Option Compare Database
Option Explicit

' ===========================================================================
' modBackend - Frontend/Backend Architektur-Verwaltung
' ===========================================================================
' Verwaltet die Verknuepfung zwischen Frontend (.accdb mit Code/Forms)
' und Backend (.accdb auf Netzlaufwerk mit Datentabellen).
'
' Frontend-Tabellen (bleiben lokal):
'   tblConfig, tblSyncProfil, tblSyncProfilOrdner
'
' Backend-Tabellen (auf Netzlaufwerk):
'   tblSyncLauf, tblKontakte, tblOutlookOrdner, tblEmailThreads,
'   tblEmails, tblEmailContent, tblEmailEmpfaenger, tblEmailAnhaenge,
'   tblEmailStatus, tblProjekte, tblEmailProjekt
'
' WICHTIG: Wenn kein Backend konfiguriert ist, arbeitet alles lokal.
'          Die Verknuepfung ist OPTIONAL und kann jederzeit hergestellt werden.
'
' NETZWERK-SCHUTZ (v0.5.2):
'   - Proaktiver Health-Check VOR jedem DB-Zugriff (PruefeBackendVorZugriff)
'   - Globales Offline-Flag (g_blnBackendOffline) verhindert DB-Zugriffe
'   - Jet/ACE Timeout-Optimierung (BackendOptimierTimeouts)
'   - Netzwerk-Fehlercode-Erkennung (IstNetzwerkFehler)
'   - Watchdog-Timer (StartBackendWatchdog/StoppBackendWatchdog)
'   ? Verhindert den Jet-Dialog "Netzwerkzugriff unterbrochen"
'
' Aufruf:
'   VerknuepfeBackend "\\Server\Share\OutlookSync_BE.accdb"
'   TrenneBackend
'   ? IstBackendVerfuegbar()
'   ? BackendStatus()
'   BackendOptimierTimeouts   ' Bei Init aufrufen!
'   StartBackendWatchdog       ' Periodische Pruefung starten
'
' Abhaengigkeiten: modSchema (TabelleExistiert, ErstelleAlleTabellen),
'                  modStringUtils (NormalisierePfad, ErstelleOrdner),
'                  modLogging
' ===========================================================================


' ---------------------------------------------------------------------------
' NETZWERK-SCHUTZ: Globaler Status + Watchdog
' ---------------------------------------------------------------------------

' Globales Offline-Flag: Wird True wenn Backend nicht erreichbar
' Alle DAO-Operationen pruefen dieses Flag VOR dem Zugriff
Public g_blnBackendOffline As Boolean

' Watchdog-Status
Private m_blnWatchdogAktiv  As Boolean
Private Const WATCHDOG_INTERVALL_SEK As Long = 30  ' Pruef-Intervall in Sekunden

' Windows API Timer (Access hat KEIN Application.OnTime wie Excel)
#If VBA7 Then
    Private Declare PtrSafe Function apiSetTimer Lib "user32" Alias "SetTimer" ( _
        ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function apiKillTimer Lib "user32" Alias "KillTimer" ( _
        ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private m_lngTimerID As LongPtr
#Else
    Private Declare Function apiSetTimer Lib "user32" Alias "SetTimer" ( _
        ByVal hWnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function apiKillTimer Lib "user32" Alias "KillTimer" ( _
        ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    Private m_lngTimerID As Long
#End If

' Watchdog-Form: Name des versteckten Formulars (Form-Timer > API-Timer)
Private Const WATCHDOG_FORM_NAME As String = "frmBackendWatchdog"
Private m_blnFormWatchdogAktiv As Boolean

' Recovery-Konstanten (Keys in tblConfig, lokal = ueberlebt Crash)
Private Const CFG_SYNC_AKTIV        As String = "SyncAktiv"          ' "1" wenn Sync laeuft
Private Const CFG_SYNC_LAUF_ID      As String = "SyncAktivLaufID"    ' Aktuelle SyncLaufID
Private Const CFG_SYNC_ORDNER       As String = "SyncAktivOrdner"    ' Ordnerpfad
Private Const CFG_SYNC_POSITION     As String = "SyncAktivPosition"  ' Letzte verarbeitete Mail-Nr
Private Const CFG_SYNC_TIMESTAMP    As String = "SyncAktivTimestamp" ' Letztes Update (Crash-Erkennung)


' ---------------------------------------------------------------------------
' TABELLEN-KLASSIFIKATION
' ---------------------------------------------------------------------------

' Backend-Tabellen (werden auf Netzlaufwerk verknuepft)
Private Function GetBackendTabellen() As Variant
    GetBackendTabellen = Array(TBL_SYNC_LAUF, TBL_KONTAKTE, TBL_OUTLOOK_ORDNER, _
                                TBL_EMAIL_THREADS, TBL_EMAILS, TBL_EMAIL_CONTENT, _
                                TBL_EMAIL_EMPFAENGER, TBL_EMAIL_ANHAENGE, TBL_EMAIL_STATUS, _
                                TBL_PROJEKTE, TBL_EMAIL_PROJEKT, _
                                TBL_SYNC_JOB, TBL_SYNC_HEARTBEAT, TBL_SYNC_CONTROL, TBL_WORKER_LEASE, _
                                TBL_WORKER_TRACE)
End Function

' Frontend-Tabellen (bleiben immer lokal)
Private Function GetFrontendTabellen() As Variant
    GetFrontendTabellen = Array(TBL_CONFIG, TBL_SYNC_PROFIL, TBL_SYNC_PROFIL_ORDNER)
End Function


' ===========================================================================
' BACKEND VERKNUEPFEN
' ===========================================================================

' Verknuepft alle Datentabellen mit einer Backend-Datenbank auf Netzlaufwerk.
' Erstellt die Backend-DB wenn sie nicht existiert.
' Migriert bestehende lokale Daten automatisch ins Backend.
'
' Parameter:
'   strBackendPfad - Vollstaendiger Pfad zur Backend-.accdb
'                    z.B. "\\Server\Share\OutlookSync_BE.accdb"
'                    oder "S:\Daten\OutlookSync_BE.accdb"
Public Function VerknuepfeBackend(ByVal strBackendPfad As String) As Boolean
    On Error GoTo ErrHandler

    If Trim(strBackendPfad) = "" Then
        LogError "VerknuepfeBackend: Kein Pfad angegeben", "BACKEND"
        VerknuepfeBackend = False
        Exit Function
    End If

    ' Pfad normalisieren
    If Right(strBackendPfad, 1) = "\" Then
        strBackendPfad = strBackendPfad & "OutlookSync_BE.accdb"
    End If
    If LCase(Right(strBackendPfad, 6)) <> ".accdb" Then
        strBackendPfad = strBackendPfad & ".accdb"
    End If

    Debug.Print String(70, "=")
    Debug.Print "=== BACKEND VERKNUEPFEN ==="
    Debug.Print "    Pfad: " & strBackendPfad
    Debug.Print String(70, "=")

    ' 1. Backend-DB erstellen (wenn noetig)
    If Dir(strBackendPfad) = "" Then
        Debug.Print "  Backend-DB existiert nicht, wird erstellt..."
        If Not ErstelleBackendDB(strBackendPfad) Then
            LogError "Backend-DB konnte nicht erstellt werden: " & strBackendPfad, "BACKEND"
            VerknuepfeBackend = False
            Exit Function
        End If
    End If

    ' 2. Pruefen ob Backend erreichbar
    If Not PruefeDatei(strBackendPfad) Then
        LogError "Backend nicht erreichbar: " & strBackendPfad, "BACKEND"
        VerknuepfeBackend = False
        Exit Function
    End If

    ' 3. Tabellen verknuepfen
    Dim arrTabellen As Variant
    Dim i As Long
    Dim lngOK As Long, lngFail As Long
    arrTabellen = GetBackendTabellen()
    lngOK = 0: lngFail = 0

    For i = LBound(arrTabellen) To UBound(arrTabellen)
        If VerknuepfeTabelle(CStr(arrTabellen(i)), strBackendPfad) Then
            lngOK = lngOK + 1
        Else
            lngFail = lngFail + 1
        End If
    Next i

    ' 4. Pfad in Config speichern (Cache + DB)
    CacheSetConfig CFG_BACKEND_PFAD, strBackendPfad

    Debug.Print "  Verknuepft: " & lngOK & " OK, " & lngFail & " Fehler"
    Debug.Print String(70, "=")

    LogInfo "Backend verknuepft: " & strBackendPfad & " (" & lngOK & " Tabellen)", "BACKEND"
    VerknuepfeBackend = (lngFail = 0)
    Exit Function

ErrHandler:
    HandleError "modBackend", "VerknuepfeBackend"
    VerknuepfeBackend = False
End Function


' ===========================================================================
' BACKEND TRENNEN
' ===========================================================================

' Entfernt alle Verknuepfungen und stellt lokale Tabellen wieder her.
' ACHTUNG: Die Daten bleiben im Backend, lokal werden leere Tabellen erstellt.
Public Sub TrenneBackend()
    On Error GoTo ErrHandler

    Dim arrTabellen As Variant
    Dim i As Long
    arrTabellen = GetBackendTabellen()

    Debug.Print String(70, "=")
    Debug.Print "=== BACKEND TRENNEN ==="

    For i = LBound(arrTabellen) To UBound(arrTabellen)
        Call EntferneVerknuepfung(CStr(arrTabellen(i)))
    Next i

    ' Backend-Pfad aus Config entfernen (Cache + DB)
    CacheSetConfig CFG_BACKEND_PFAD, ""

    ' Lokale Tabellen neu erstellen
    Debug.Print "  Erstelle lokale Tabellen..."
    ErstelleAlleTabellen

    Debug.Print "=== Backend getrennt ==="
    Debug.Print String(70, "=")

    LogInfo "Backend getrennt, lokale Tabellen wiederhergestellt", "BACKEND"
    Exit Sub

ErrHandler:
    HandleError "modBackend", "TrenneBackend"
End Sub


' ===========================================================================
' BACKEND-STATUS PRUEFEN
' ===========================================================================

' Prueft ob das konfigurierte Backend erreichbar ist
' Wenn kein Backend konfiguriert: TRUE (lokal = immer verfuegbar)
Public Function IstBackendVerfuegbar() As Boolean
    On Error Resume Next
    Dim strPfad As String
    strPfad = CacheGetConfig(CFG_BACKEND_PFAD, "")

    If strPfad = "" Then
        ' Kein Backend konfiguriert = lokaler Modus = immer verfuegbar
        IstBackendVerfuegbar = True
        Exit Function
    End If

    IstBackendVerfuegbar = PruefeDatei(strPfad)
    On Error GoTo 0
End Function


' Gibt eine menschenlesbare Status-Zusammenfassung zurueck
Public Function BackendStatus() As String
    Dim strPfad As String
    strPfad = CacheGetConfig(CFG_BACKEND_PFAD, "")

    If strPfad = "" Then
        BackendStatus = "Lokal (kein Backend konfiguriert)"
        Exit Function
    End If

    If PruefeDatei(strPfad) Then
        BackendStatus = "Backend OK: " & strPfad
    Else
        BackendStatus = "Backend NICHT ERREICHBAR: " & strPfad
    End If
End Function


' Prueft ob aktuell ein Backend verknuepft ist.
' Prueft zusaetzlich ob tatsaechlich Linked Tables existieren,
' nicht nur ob ein Pfad in der Config steht.
Public Function IstBackendVerknuepft() As Boolean
    ' Schnell-Check: Ist ueberhaupt ein Pfad konfiguriert?
    If CacheGetConfig(CFG_BACKEND_PFAD, "") = "" Then
        IstBackendVerknuepft = False
        Exit Function
    End If

    ' Sicherheits-Check: Gibt es tatsaechlich verlinkte Tabellen?
    ' WICHTIG: CurrentDb in Variable halten - ohne Variable kann VBA
    ' die temporaere DB-Referenz mid-loop per GC freigeben (Access-Gotcha).
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If Len(td.Connect) > 0 And Left(td.Connect, 10) = ";DATABASE=" Then
            IstBackendVerknuepft = True
            Set db = Nothing
            Err.Clear
            Exit Function
        End If
    Next td
    Set db = Nothing
    Err.Clear
    On Error GoTo 0

    ' Pfad konfiguriert aber keine Links -> Config ist veraltet
    IstBackendVerknuepft = False
End Function


' Gibt den konfigurierten Backend-Pfad zurueck (leer wenn lokal)
Public Function GetBackendPfad() As String
    GetBackendPfad = CacheGetConfig(CFG_BACKEND_PFAD, "")
End Function


' Prueft ob eine Tabelle eine Frontend-Tabelle ist (bleibt lokal)
Public Function IstFETabelle(ByVal strTabelle As String) As Boolean
    Dim arr As Variant, v As Variant
    arr = GetFrontendTabellen()
    For Each v In arr
        If LCase(CStr(v)) = LCase(strTabelle) Then
            IstFETabelle = True
            Exit Function
        End If
    Next v
    IstFETabelle = False
End Function


' ===========================================================================
' RECONNECT (nach Netzwerk-Unterbrechung)
' ===========================================================================

' Versucht die Backend-Verbindung wiederherzustellen
' (Linked Table Links refreshen + Netzwerk-Recovery)
' v0.5.1: Mit Retry-Logik und VPN-Warteschleife
Public Function ReconnectBackend() As Boolean
    On Error GoTo ErrHandler

    Dim strPfad As String
    strPfad = CacheGetConfig(CFG_BACKEND_PFAD, "")

    If strPfad = "" Then
        ReconnectBackend = True  ' Lokal = immer OK
        Exit Function
    End If

    ' Netzwerk-Pruefung mit Wartelogik (VPN/LAN)
    Dim intRetries As Integer
    Dim lngPause As Long
    intRetries = CInt(CacheGetConfig(CFG_NETZWERK_RETRIES, "3"))
    lngPause = CLng(CacheGetConfig(CFG_NETZWERK_PAUSE, "2000"))

    Dim intVersuch As Integer
    For intVersuch = 1 To intRetries
        If PruefeDatei(strPfad) Then GoTo Reconnect_Links

        LogWarn "Backend nicht erreichbar (Versuch " & intVersuch & "/" & _
                intRetries & ") - warte " & lngPause & "ms...", "BACKEND"
        Sleep lngPause
        DoEvents

        ' Exponentielles Backoff (max 10s)
        If lngPause < 10000 Then lngPause = lngPause * 2
    Next intVersuch

    ' Alle Versuche gescheitert
    LogError "Backend nach " & intRetries & " Versuchen nicht erreichbar: " & strPfad, "BACKEND"
    ReconnectBackend = False
    Exit Function

Reconnect_Links:
    ' Links refreshen
    Dim arrTabellen As Variant, i As Long
    arrTabellen = GetBackendTabellen()

    For i = LBound(arrTabellen) To UBound(arrTabellen)
        Call RefreshLink(CStr(arrTabellen(i)), strPfad)
    Next i

    LogInfo "Backend Reconnect erfolgreich: " & strPfad, "BACKEND"
    ReconnectBackend = True
    Exit Function

ErrHandler:
    HandleError "modBackend", "ReconnectBackend"
    ReconnectBackend = False
End Function


' Health-Check: Prueft ob ein SELECT auf eine Backend-Tabelle funktioniert.
' Erkennt kaputte Linked-Table-Verbindungen (z.B. nach VPN-Reconnect).
' Versuch automatischer Reparatur bei Fehler.
' v0.5.1
Public Function BackendHealthCheck() As Boolean
    On Error GoTo ErrHandler

    Dim strPfad As String
    strPfad = CacheGetConfig(CFG_BACKEND_PFAD, "")

    ' Kein Backend = lokal = immer OK
    If strPfad = "" Then
        BackendHealthCheck = True
        Exit Function
    End If

    ' Test-Query auf leichteste Tabelle
    Dim db As DAO.Database
    Set db = CurrentDb

    On Error GoTo HealthFailed
    Dim lngTest As Long
    lngTest = DCount("*", TBL_SYNC_LAUF)
    Set db = Nothing

    BackendHealthCheck = True
    Exit Function

HealthFailed:
    ' Link kaputt -> Reconnect versuchen
    LogWarn "Backend HealthCheck fehlgeschlagen: " & Err.Description & _
            " -> versuche Reconnect", "BACKEND"
    Err.Clear
    On Error GoTo ErrHandler

    BackendHealthCheck = ReconnectBackend()
    Exit Function

ErrHandler:
    HandleError "modBackend", "BackendHealthCheck"
    BackendHealthCheck = False
End Function


' ===========================================================================
' PRIVATE: TABELLEN-VERKNUEPFUNG
' ===========================================================================

' Einzelne Tabelle mit Backend verknuepfen
' Migriert bestehende lokale Daten automatisch
Private Function VerknuepfeTabelle(ByVal strTabelle As String, _
                                    ByVal strBackendDB As String) As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    ' 1. Wenn lokale (nicht-verknuepfte) Tabelle existiert: Daten migrieren
    If TabelleExistiert(strTabelle) Then
        Dim td As DAO.TableDef
        Set td = db.TableDefs(strTabelle)

        If td.Connect = "" Then
            ' Lokale Tabelle -> Daten ins Backend kopieren
            Call MigriereTabellenDaten(strTabelle, strBackendDB)
            ' Lokale Tabelle loeschen
            db.Execute "DROP TABLE [" & strTabelle & "]"
            db.TableDefs.Refresh
            Debug.Print "  [MIGRIERT] " & strTabelle
        Else
            ' Bereits verknuepft -> nur Link aktualisieren
            Call EntferneVerknuepfung(strTabelle)
        End If
    End If

    ' 2. Verknuepfung erstellen
    Dim tdNew As DAO.TableDef
    Set tdNew = db.CreateTableDef(strTabelle)
    tdNew.Connect = ";DATABASE=" & strBackendDB
    tdNew.SourceTableName = strTabelle
    db.TableDefs.Append tdNew
    db.TableDefs.Refresh

    Debug.Print "  [LINK  ] " & strTabelle & " -> " & strBackendDB
    Set db = Nothing
    VerknuepfeTabelle = True
    Exit Function

ErrHandler:
    Debug.Print "  [FAIL  ] " & strTabelle & " - " & Err.Description
    LogWarn "Verknuepfung fehlgeschlagen: " & strTabelle & " - " & Err.Description, "BACKEND"
    VerknuepfeTabelle = False
End Function


' Daten einer lokalen Tabelle ins Backend migrieren
Private Sub MigriereTabellenDaten(ByVal strTabelle As String, _
                                   ByVal strBackendDB As String)
    On Error GoTo ErrHandler

    ' Pruefen ob lokale Tabelle Daten hat
    Dim lngCount As Long
    lngCount = DCount("*", strTabelle)
    If lngCount = 0 Then Exit Sub

    ' Pruefen ob Backend-Tabelle bereits Daten hat
    Dim dbBE As DAO.Database
    Set dbBE = DBEngine.OpenDatabase(strBackendDB)

    Dim lngBECount As Long
    On Error Resume Next
    lngBECount = dbBE.OpenRecordset("SELECT COUNT(*) FROM [" & strTabelle & "]").Fields(0)
    If Err.Number <> 0 Then lngBECount = 0: Err.Clear
    On Error GoTo ErrHandler

    If lngBECount = 0 Then
        ' Backend-Tabelle leer -> Daten kopieren
        CurrentDb.Execute "INSERT INTO [" & strTabelle & "] IN '" & strBackendDB & "' " & _
                          "SELECT * FROM [" & strTabelle & "]"
        LogInfo "Migriert: " & lngCount & " Datensaetze von " & strTabelle, "BACKEND"
        Debug.Print "    -> " & lngCount & " Datensaetze migriert"
    Else
        LogWarn "Backend-Tabelle " & strTabelle & " hat bereits " & lngBECount & _
                " Datensaetze, Migration uebersprungen", "BACKEND"
    End If

    dbBE.Close: Set dbBE = Nothing
    Exit Sub

ErrHandler:
    LogWarn "Datenmigration fehlgeschlagen fuer " & strTabelle & ": " & _
            Err.Description, "BACKEND"
End Sub


' Verknuepfung einer Tabelle entfernen
Private Sub EntferneVerknuepfung(ByVal strTabelle As String)
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb

    If TabelleExistiert(strTabelle) Then
        db.TableDefs.Delete strTabelle
        db.TableDefs.Refresh
        Debug.Print "  [UNLINK] " & strTabelle
    End If

    Set db = Nothing
    On Error GoTo 0
End Sub


' Link einer verknuepften Tabelle aktualisieren
Private Sub RefreshLink(ByVal strTabelle As String, ByVal strBackendDB As String)
    On Error Resume Next
    Dim td As DAO.TableDef
    Set td = CurrentDb.TableDefs(strTabelle)

    If Not td Is Nothing Then
        td.Connect = ";DATABASE=" & strBackendDB
        td.RefreshLink
        Debug.Print "  [REFRESH] " & strTabelle
    End If

    On Error GoTo 0
End Sub


' ===========================================================================
' PRIVATE: BACKEND-DB ERSTELLEN
' ===========================================================================

' Erstellt eine neue Access-Datenbank mit allen BE-Tabellen + Indizes
' v0.5.1: Nutzt modSchema.ErstelleBackendTabellenInDB (Single Source of Truth)
Private Function ErstelleBackendDB(ByVal strPfad As String) As Boolean
    On Error GoTo ErrHandler

    ' Verzeichnis erstellen
    Dim strDir As String
    strDir = Left(strPfad, InStrRev(strPfad, "\"))
    If strDir <> "" Then ErstelleOrdner strDir

    ' Neue Access-DB erstellen
    Dim dbNew As DAO.Database
    Set dbNew = DBEngine.CreateDatabase(strPfad, dbLangGeneral)

    ' Schema aus zentraler Quelle erstellen (modSchema)
    If Not ErstelleBackendTabellenInDB(dbNew) Then
        dbNew.Close: Set dbNew = Nothing
        LogError "Backend-Tabellen konnten nicht erstellt werden: " & strPfad, "BACKEND"
        ErstelleBackendDB = False
        Exit Function
    End If

    dbNew.Close
    Set dbNew = Nothing

    Debug.Print "  [OK  ] Backend-DB erstellt: " & strPfad
    LogInfo "Backend-DB erstellt: " & strPfad, "BACKEND"
    ErstelleBackendDB = True
    Exit Function

ErrHandler:
    Debug.Print "  [FAIL] Backend-DB Erstellung: " & Err.Description
    HandleError "modBackend", "ErstelleBackendDB"
    ErstelleBackendDB = False
End Function


' Prueft ob eine Datei existiert und zugreifbar ist
Private Function PruefeDatei(ByVal strPfad As String) As Boolean
    On Error Resume Next
    PruefeDatei = (Dir(strPfad) <> "")
    If Err.Number <> 0 Then PruefeDatei = False: Err.Clear
    On Error GoTo 0
End Function


' ===========================================================================
' NETZWERK-SCHUTZ (v0.5.2)
' ===========================================================================
' Verhindert den Jet/ACE-Dialog "Ihr Netzwerkzugriff wurde unterbrochen.
' Schliessen Sie die Datenbank..." durch:
'   1. Proaktive Pruefung VOR jedem DB-Zugriff
'   2. Reduzierte Jet-Timeouts (schnellerer VBA-Error statt Dialog)
'   3. Netzwerk-Fehlercode-Klassifikation
'   4. Watchdog-Timer fuer periodische Pruefung
'   5. Globales Offline-Flag als Schutzschalter
'
' HINTERGRUND:
' Der Jet-Dialog erscheint auf System-Ebene BEVOR VBA-Error-Handler
' greifen koennen. Die einzige zuverlaessige Praevention ist,
' gar nicht erst auf kaputte Linked Tables zuzugreifen.
' ===========================================================================


' ---------------------------------------------------------------------------
' NETZWERK-FEHLER ERKENNEN
' ---------------------------------------------------------------------------

' Prueft ob ein VBA-Fehlercode auf Netzwerk/Dateisystem-Probleme hinweist.
' Diese Fehler koennen bei Linked-Table-Zugriffen auftreten wenn das
' Backend-Netzlaufwerk nicht erreichbar ist.
Public Function IstNetzwerkFehler(ByVal lngErrNum As Long) As Boolean
    Select Case lngErrNum
        Case 3043   ' Datentraeger- oder Netzwerkfehler
            IstNetzwerkFehler = True
        Case 3044   ' Pfad ist ungueltig
            IstNetzwerkFehler = True
        Case 3045   ' Datei konnte nicht geoeffnet werden
            IstNetzwerkFehler = True
        Case 3024   ' Datei nicht gefunden
            IstNetzwerkFehler = True
        Case 3051   ' Datenbank-Engine kann Datei nicht oeffnen
            IstNetzwerkFehler = True
        Case 3055   ' Kein gueltiger Dateiname
            IstNetzwerkFehler = True
        Case 3197   ' Daten wurden geaendert; Vorgang abgebrochen
            IstNetzwerkFehler = True
        Case 3353   ' Vorgang konnte nicht ausgefuehrt werden
            IstNetzwerkFehler = True
        Case 3356   ' Keine Leseberechtigung fuer Datei
            IstNetzwerkFehler = True
        Case 3021   ' Kein aktueller Datensatz (kann bei Verbindungsverlust kommen)
            IstNetzwerkFehler = True
        Case Else
            IstNetzwerkFehler = False
    End Select
End Function


' ---------------------------------------------------------------------------
' JET/ACE TIMEOUT-OPTIMIERUNG
' ---------------------------------------------------------------------------

' Optimiert die Jet/ACE-Engine-Timeouts um bei Netzwerkverlust
' schneller einen VBA-Error auszuloesen statt den System-Dialog zu zeigen.
'
' MUSS einmal beim Start aufgerufen werden (z.B. in InitGlobals).
' Ohne diese Einstellung wartet Jet bis zu 5-10 Sekunden und zeigt
' dann den nicht-abfangbaren Dialog.
Public Sub BackendOptimierTimeouts()
    On Error Resume Next

    ' PageTimeout: Wie lange Jet auf eine I/O-Seite wartet (in ms)
    ' Default: 5000ms (5 Sek) -> Reduziert auf 2000ms (2 Sek)
    ' Niedrigerer Wert = schnellerer Fehler = eher VBA-Error statt Dialog
    DBEngine.SetOption dbPageTimeout, 2000

    ' SharedAsyncDelay: Verzögerung fuer Shared-Modus-Refresh
    ' Default: 0  -> Kein Einfluss, aber sauber setzen
    DBEngine.SetOption dbSharedAsyncDelay, 50

    ' ExclusiveAsyncDelay: Verzögerung im Exclusive-Modus
    DBEngine.SetOption dbExclusiveAsyncDelay, 50

    ' FlushTransactionTimeout: Wie lang Jet wartet bevor es Transaktions-
    ' Daten auf Datentraeger schreibt (ms). Niedriger = weniger Datenverlust.
    DBEngine.SetOption dbFlushTransactionTimeout, 500

    On Error GoTo 0

    LogDebug "Jet/ACE Timeouts optimiert (PageTimeout=2000ms)", "BACKEND"
End Sub


' ---------------------------------------------------------------------------
' PROAKTIVE BACKEND-PRUEFUNG (VOR DB-ZUGRIFF)
' ---------------------------------------------------------------------------

' Prueft BEVOR eine DAO-Operation ausgefuehrt wird, ob das Backend
' erreichbar ist. Setzt g_blnBackendOffline und versucht bei Bedarf
' einen Reconnect.
'
' AUFRUF: Am Anfang jeder DAO-Funktion die auf Backend-Tabellen zugreift:
'   If Not PruefeBackendVorZugriff() Then Exit Function
'
' Rueckgabe:
'   True  = Backend OK (oder lokal, kein Backend konfiguriert)
'   False = Backend offline, kein Reconnect moeglich -> KEIN DB-Zugriff!
Public Function PruefeBackendVorZugriff() As Boolean
    On Error Resume Next

    ' Kein Backend konfiguriert = lokal = immer OK
    If Not IstBackendVerknuepft() Then
        g_blnBackendOffline = False
        PruefeBackendVorZugriff = True
        Err.Clear
        Exit Function
    End If

    ' Schnell-Pruefung: Ist Backend-Datei erreichbar?
    Dim strPfad As String
    strPfad = CacheGetConfig(CFG_BACKEND_PFAD, "")

    If PruefeDatei(strPfad) Then
        ' Backend erreichbar -> Flag zuruecksetzen
        If g_blnBackendOffline Then
            LogInfo "Backend wieder erreichbar: " & strPfad, "BACKEND"
        End If
        g_blnBackendOffline = False
        PruefeBackendVorZugriff = True
        Err.Clear
        Exit Function
    End If

    ' Backend NICHT erreichbar!
    If Not g_blnBackendOffline Then
        ' Erstmalige Erkennung -> Loggen
        LogWarn "Backend-Verbindung verloren: " & strPfad, "BACKEND"
    End If
    g_blnBackendOffline = True

    ' Reconnect-Versuch (mit Retry-Logik aus ReconnectBackend)
    LogInfo "Versuche Backend-Reconnect...", "BACKEND"
    Err.Clear
    ' WICHTIG: On Error Resume Next beibehalten!
    ' On Error GoTo 0 wuerde Fehler aus ReconnectBackend zum Caller
    ' propagieren und dessen ErrHandler ausloesen (stiller Abbruch).

    If ReconnectBackend() Then
        g_blnBackendOffline = False
        LogInfo "Backend-Reconnect erfolgreich", "BACKEND"
        PruefeBackendVorZugriff = True
    Else
        LogError "Backend-Reconnect fehlgeschlagen - DB-Zugriffe blockiert", "BACKEND"
        PruefeBackendVorZugriff = False
    End If
    Err.Clear
End Function


' Behandelt einen Netzwerkfehler der waehrend einer DAO-Operation aufgetreten ist.
' Setzt das Offline-Flag und loggt den Fehler.
' Gibt True zurueck wenn es ein Netzwerkfehler war (Caller sollte abbrechen).
'
' AUFRUF: Im ErrHandler jeder DAO-Funktion:
'   If BehandleNetzwerkFehler(Err.Number, "modDAO", "MeineFunktion") Then
'       Exit Function  ' Sauber abbrechen statt Jet-Dialog riskieren
'   End If
Public Function BehandleNetzwerkFehler(ByVal lngErrNum As Long, _
                                        ByVal strModul As String, _
                                        ByVal strProzedur As String) As Boolean
    If Not IstNetzwerkFehler(lngErrNum) Then
        BehandleNetzwerkFehler = False
        Exit Function
    End If

    ' Netzwerkfehler erkannt!
    g_blnBackendOffline = True
    LogError strModul & "." & strProzedur & ": Netzwerkfehler [" & lngErrNum & "] " & _
             Err.Description & " -> Backend als offline markiert", "NETZWERK"

    BehandleNetzwerkFehler = True
End Function


' ---------------------------------------------------------------------------
' BACKEND-WATCHDOG (Periodische Pruefung)
' ---------------------------------------------------------------------------

' Startet die periodische Backend-Pruefung.
' Nutzt Windows API SetTimer (Access hat kein Application.OnTime).
'
' Der Watchdog prueft alle WATCHDOG_INTERVALL_SEK Sekunden ob das Backend
' erreichbar ist und setzt g_blnBackendOffline entsprechend.
' Bei Offline-Erkennung wird automatisch ein Reconnect versucht.
Public Sub StartBackendWatchdog()
    On Error Resume Next

    ' Nur starten wenn Backend konfiguriert
    If Not IstBackendVerknuepft() Then
        LogDebug "Watchdog nicht gestartet (kein Backend konfiguriert)", "BACKEND"
        Exit Sub
    End If

    ' Alten Timer stoppen falls vorhanden
    StoppBackendWatchdog

    m_blnWatchdogAktiv = True
    ' Timer via Windows API starten (feuert periodisch)
    m_lngTimerID = apiSetTimer(0, 0, WATCHDOG_INTERVALL_SEK * 1000, AddressOf WatchdogTimerCallback)

    If m_lngTimerID = 0 Then
        LogWarn "Watchdog-Timer konnte nicht erstellt werden", "BACKEND"
        m_blnWatchdogAktiv = False
    Else
        LogInfo "Backend-Watchdog gestartet (Intervall: " & WATCHDOG_INTERVALL_SEK & "s)", "BACKEND"
    End If
    On Error GoTo 0
End Sub


' Stoppt den Watchdog-Timer
Public Sub StoppBackendWatchdog()
    On Error Resume Next

    m_blnWatchdogAktiv = False
    ' Timer stoppen
    If m_lngTimerID <> 0 Then
        apiKillTimer 0, m_lngTimerID
        m_lngTimerID = 0
    End If

    LogInfo "Backend-Watchdog gestoppt", "BACKEND"
    On Error GoTo 0
End Sub


' Timer-Callback: Wird von der Windows API aufgerufen.
' ACHTUNG: Unbehandelte Fehler in dieser Routine crashen Access!
#If VBA7 Then
Public Sub WatchdogTimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
                                  ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub WatchdogTimerCallback(ByVal hwnd As Long, ByVal uMsg As Long, _
                                  ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    On Error Resume Next  ' KRITISCH: Muss IMMER aktiv sein!
    BackendWatchdogTick
    On Error GoTo 0
End Sub


' Watchdog-Tick: Wird periodisch vom API-Timer aufgerufen.
' MUSS Public sein fuer WatchdogTimerCallback.
Public Sub BackendWatchdogTick()
    On Error Resume Next

    ' Watchdog noch aktiv?
    If Not m_blnWatchdogAktiv Then Exit Sub

    ' Backend-Verfuegbarkeit pruefen (leichtgewichtig via Dir())
    Dim strPfad As String
    strPfad = CacheGetConfig(CFG_BACKEND_PFAD, "")

    If strPfad = "" Then
        m_blnWatchdogAktiv = False
        Exit Sub
    End If

    Dim blnErreichbar As Boolean
    blnErreichbar = PruefeDatei(strPfad)

    If blnErreichbar Then
        ' Backend OK
        If g_blnBackendOffline Then
            ' War offline, jetzt wieder da -> Reconnect Links
            LogInfo "Watchdog: Backend wieder erreichbar -> Reconnect", "BACKEND"
            g_blnBackendOffline = False
            Err.Clear
            On Error Resume Next
            ReconnectBackend
        End If
    Else
        ' Backend nicht erreichbar
        If Not g_blnBackendOffline Then
            LogWarn "Watchdog: Backend-Verbindung verloren! DB-Zugriffe blockiert.", "BACKEND"
            g_blnBackendOffline = True
        End If
    End If

    ' SetTimer feuert automatisch periodisch — keine Neuplanung noetig
    On Error GoTo 0
End Sub


' ===========================================================================
' WATCHDOG-FORM (v0.5.2) - Alternative zum API-Timer
' ===========================================================================
' Ein verstecktes Formular mit Timer-Control ist eine Alternative zum
' SetTimer-API-Watchdog, weil:
'   - Timer-Events ueber die Windows Message Pump laufen
'   - Sie feuern auch waehrend langer VBA-Operationen (zwischen DAO-Calls)
'   - Application.OnTime wird erst nach dem aktuellen Code-Block verarbeitet
'
' Das Formular "frmBackendWatchdog" muss einmalig erstellt werden:
'   ErstelleWatchdogForm     ' Erstellt Form + Timer programmatisch
'
' Danach kann der Form-Watchdog verwendet werden:
'   StarteFormWatchdog       ' Oeffnet Form versteckt, Timer laeuft
'   StoppeFormWatchdog       ' Schliesst Form
'
' FORM-TIMER EVENT:
'   Der Timer ruft alle 5 Sekunden WatchdogFormTick auf.
'   Diese Routine ist das "Herzstueck" - sie merkt Netzwerkverlust
'   bevor die naechste DAO-Operation den Jet-Dialog ausloesen kann.
' ===========================================================================


' Erstellt das Watchdog-Formular programmatisch (einmalig ausfuehren).
' Das Formular hat nur ein Timer-Control, keine sichtbaren Elemente.
Public Sub ErstelleWatchdogForm()
    On Error GoTo ErrHandler

    ' Pruefen ob Form bereits existiert
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = WATCHDOG_FORM_NAME Then
            Debug.Print "Watchdog-Form existiert bereits: " & WATCHDOG_FORM_NAME
            Exit Sub
        End If
    Next obj

    ' Form programmatisch erstellen
    Dim frm As Form
    Set frm = CreateForm()

    ' Minimale Groesse, kein Rand, kein Datensatz, versteckbar
    frm.Caption = "Backend Watchdog"
    frm.DefaultView = 0          ' Einzelformular
    frm.ScrollBars = 0           ' Keine Scrollbars
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.Width = 3000             ' Minimal (ca. 2cm)
    frm.Section(0).Height = 1000

    ' Timer-Intervall setzen (5000ms = 5 Sekunden)
    frm.TimerInterval = 5000

    ' Label als Status-Anzeige (optional sichtbar)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acLabel, acDetail, , , 100, 100, 2800, 400)
    ctl.Name = "lblStatus"
    ctl.Caption = "Backend: Pruefe..."
    ctl.FontSize = 8

    ' Speichern unter dem gewuenschten Namen
    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name
    DoCmd.Rename WATCHDOG_FORM_NAME, acForm, frm.Name

    Debug.Print "[OK] Watchdog-Form erstellt: " & WATCHDOG_FORM_NAME
    Debug.Print "     Timer-Code muss noch eingefuegt werden!"
    Debug.Print "     -> Im Form-Modul: Private Sub Form_Timer()"
    Debug.Print "        WatchdogFormTick Me"
    Debug.Print "        End Sub"

    LogInfo "Watchdog-Form erstellt: " & WATCHDOG_FORM_NAME, "BACKEND"
    Exit Sub

ErrHandler:
    Debug.Print "[FAIL] Watchdog-Form Erstellung: " & Err.Description
    HandleError "modBackend", "ErstelleWatchdogForm"
End Sub


' Startet den Form-basierten Watchdog (oeffnet Form versteckt)
Public Sub StarteFormWatchdog()
    On Error GoTo ErrHandler

    ' Nur starten wenn Backend konfiguriert
    If Not IstBackendVerknuepft() Then
        LogDebug "Form-Watchdog nicht gestartet (kein Backend)", "BACKEND"
        Exit Sub
    End If

    ' Pruefen ob Form existiert
    Dim blnExistiert As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = WATCHDOG_FORM_NAME Then
            blnExistiert = True
            Exit For
        End If
    Next obj

    If Not blnExistiert Then
        ' Kein Form -> Fallback auf API-Timer Watchdog
        LogDebug "Watchdog-Form nicht vorhanden, nutze SetTimer API", "BACKEND"
        StartBackendWatchdog
        Exit Sub
    End If

    ' Form versteckt oeffnen (acHidden)
    DoCmd.OpenForm WATCHDOG_FORM_NAME, , , , , acHidden
    m_blnFormWatchdogAktiv = True

    ' API-Timer Watchdog stoppen (nicht beide parallel)
    StoppBackendWatchdog

    LogInfo "Form-Watchdog gestartet (5s Intervall, zuverlaessiger)", "BACKEND"
    Exit Sub

ErrHandler:
    ' Fallback auf API-Timer
    LogWarn "Form-Watchdog konnte nicht gestartet werden: " & Err.Description & _
            " -> Fallback auf SetTimer API", "BACKEND"
    StartBackendWatchdog
End Sub


' Stoppt den Form-basierten Watchdog
Public Sub StoppeFormWatchdog()
    On Error Resume Next

    If m_blnFormWatchdogAktiv Then
        DoCmd.Close acForm, WATCHDOG_FORM_NAME
        m_blnFormWatchdogAktiv = False
        LogInfo "Form-Watchdog gestoppt", "BACKEND"
    End If

    On Error GoTo 0
End Sub


' Timer-Tick: Wird vom Form_Timer-Event aufgerufen.
' Diese Routine legt man in modBackend (nicht im Form-Modul) damit
' die Logik zentral bleibt. Im Form-Modul steht nur:
'   Private Sub Form_Timer()
'       WatchdogFormTick Me
'   End Sub
Public Sub WatchdogFormTick(frm As Form)
    On Error Resume Next

    Dim strPfad As String
    strPfad = CacheGetConfig(CFG_BACKEND_PFAD, "")
    If strPfad = "" Then Exit Sub

    ' Leichtgewichtige Pruefung (Dir, kein DB-Zugriff!)
    Dim blnOK As Boolean
    blnOK = (Dir(strPfad) <> "")
    If Err.Number <> 0 Then blnOK = False: Err.Clear

    ' Status-Label aktualisieren (wenn vorhanden)
    Dim lblStatus As Control
    Set lblStatus = frm.Controls("lblStatus")
    If Not lblStatus Is Nothing Then
        If blnOK Then
            lblStatus.Caption = "Backend: OK  " & Format(Now, "hh:nn:ss")
            lblStatus.ForeColor = RGB(0, 128, 0)  ' Gruen
        Else
            lblStatus.Caption = "Backend: OFFLINE  " & Format(Now, "hh:nn:ss")
            lblStatus.ForeColor = RGB(255, 0, 0)   ' Rot
        End If
    End If
    Set lblStatus = Nothing

    ' Zustandswechsel-Logik (identisch zum Application.OnTime Watchdog)
    If blnOK Then
        If g_blnBackendOffline Then
            LogInfo "Form-Watchdog: Backend wieder erreichbar -> Reconnect", "BACKEND"
            g_blnBackendOffline = False
            Err.Clear
            ReconnectBackend
        End If
    Else
        If Not g_blnBackendOffline Then
            LogWarn "Form-Watchdog: Backend-Verbindung verloren!", "BACKEND"
            g_blnBackendOffline = True
        End If
    End If

    On Error GoTo 0
End Sub


' ===========================================================================
' CRASH-RECOVERY (v0.5.2)
' ===========================================================================
' Speichert den Sync-Zustand laufend in die LOKALE tblConfig.
' Da tblConfig eine Frontend-Tabelle ist, ueberlebt diese Information
' einen Netzwerkverlust und sogar ein erzwungenes Schliessen von Access.
'
' Ablauf:
'   1. SyncFolder ruft SyncZustandMarkieren("1", ...) am Start
'   2. Alle 25 Mails: SyncZustandAktualisieren(i) speichert Position
'   3. Am Ende: SyncZustandLoeschen() -> sauberes Ende
'   4. Beim naechsten Start: PruefeDirtyShutdown() erkennt Crash
'      -> HoleSyncZustand() liefert letzten bekannten Stand
' ===========================================================================


' Markiert: "Es laeuft gerade ein Sync" (Dirty-Flag setzen)
Public Sub SyncZustandMarkieren(ByVal lngSyncLaufID As Long, _
                                 ByVal strOrdnerPfad As String)
    On Error Resume Next

    CacheSetConfig CFG_SYNC_AKTIV, "1"
    CacheSetConfig CFG_SYNC_LAUF_ID, CStr(lngSyncLaufID)
    CacheSetConfig CFG_SYNC_ORDNER, strOrdnerPfad
    CacheSetConfig CFG_SYNC_POSITION, "0"
    CacheSetConfig CFG_SYNC_TIMESTAMP, Format(Now, "yyyy-mm-dd hh:nn:ss")

    LogDebug "Sync-Zustand markiert: LaufID=" & lngSyncLaufID, "RECOVERY"
    On Error GoTo 0
End Sub


' Aktualisiert die aktuelle Position im Sync (nach jeder N-ten Mail)
' Leichtgewichtig: Nur lokale tblConfig, kein Backend-Zugriff.
Public Sub SyncZustandAktualisieren(ByVal lngPosition As Long)
    On Error Resume Next

    CacheSetConfig CFG_SYNC_POSITION, CStr(lngPosition)
    CacheSetConfig CFG_SYNC_TIMESTAMP, Format(Now, "yyyy-mm-dd hh:nn:ss")

    On Error GoTo 0
End Sub


' Loescht den Sync-Zustand (sauberes Ende)
Public Sub SyncZustandLoeschen()
    On Error Resume Next

    CacheSetConfig CFG_SYNC_AKTIV, "0"
    CacheSetConfig CFG_SYNC_LAUF_ID, ""
    CacheSetConfig CFG_SYNC_ORDNER, ""
    CacheSetConfig CFG_SYNC_POSITION, ""
    CacheSetConfig CFG_SYNC_TIMESTAMP, ""

    LogDebug "Sync-Zustand geloescht (sauberes Ende)", "RECOVERY"
    On Error GoTo 0
End Sub


' Prueft beim Start ob der letzte Sync unsauber beendet wurde.
' Gibt True zurueck wenn ein Crash/Dirty-Shutdown erkannt wurde.
Public Function PruefeDirtyShutdown() As Boolean
    On Error Resume Next

    Dim strAktiv As String
    strAktiv = CacheGetConfig(CFG_SYNC_AKTIV, "0")

    PruefeDirtyShutdown = (strAktiv = "1")

    If PruefeDirtyShutdown Then
        Dim strTS As String
        strTS = CacheGetConfig(CFG_SYNC_TIMESTAMP, "")
        LogWarn "DIRTY SHUTDOWN erkannt! Letzter Sync-Zustand von " & strTS, "RECOVERY"
    End If

    On Error GoTo 0
End Function


' Holt den gespeicherten Sync-Zustand nach einem Crash.
' Gibt ein Dictionary mit den Recovery-Informationen zurueck.
Public Function HoleSyncZustand() As Object
    On Error Resume Next

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    dict("SyncLaufID") = CLng("0" & CacheGetConfig(CFG_SYNC_LAUF_ID, "0"))
    dict("OrdnerPfad") = CacheGetConfig(CFG_SYNC_ORDNER, "")
    dict("LetztePosition") = CLng("0" & CacheGetConfig(CFG_SYNC_POSITION, "0"))
    dict("Timestamp") = CacheGetConfig(CFG_SYNC_TIMESTAMP, "")

    Set HoleSyncZustand = dict
    On Error GoTo 0
End Function


' Zeigt dem Benutzer den Recovery-Status und fragt ob fortgesetzt werden soll.
' Gibt True zurueck wenn der Benutzer fortsetzen moechte.
Public Function ZeigeCrashRecovery() As Boolean
    If Not PruefeDirtyShutdown() Then
        ZeigeCrashRecovery = False
        Exit Function
    End If

    Dim dict As Object
    Set dict = HoleSyncZustand()

    Dim strMsg As String
    strMsg = "Der letzte Sync wurde nicht sauber beendet!" & vbCrLf & vbCrLf & _
             "Ordner: " & dict("OrdnerPfad") & vbCrLf & _
             "Position: Mail " & dict("LetztePosition") & vbCrLf & _
             "Zeitpunkt: " & dict("Timestamp") & vbCrLf & vbCrLf & _
             "Moegliche Ursache: Netzwerkverlust oder Access-Absturz." & vbCrLf & vbCrLf & _
             "Der Sync kann ab der letzten Position fortgesetzt werden." & vbCrLf & _
             "(Duplikate werden automatisch erkannt und uebersprungen.)" & vbCrLf & vbCrLf & _
             "Sync-Zustand zuruecksetzen?"

    Dim antwort As VbMsgBoxResult
    antwort = MsgBox(strMsg, vbYesNo + vbExclamation, "Crash-Recovery")

    If antwort = vbYes Then
        SyncZustandLoeschen
        LogInfo "Crash-Recovery: Benutzer hat Zustand zurueckgesetzt", "RECOVERY"
    End If

    ZeigeCrashRecovery = True
    Set dict = Nothing
End Function


' ===========================================================================
' BACKUP (v0.5.3)
' ===========================================================================
' Erstellt zeitgestempelte Kopien von Frontend und/oder Backend.
' Backup-Ziel: <FE-Verzeichnis>\Backups\
' Dateiname:   FE_yyyy-mm-dd_hh-nn-ss_<Original>.accdb
'              BE_yyyy-mm-dd_hh-nn-ss_<Original>.accdb
'
' Sicherheit:
'   - Groessenvergleich Quelle vs. Kopie (Toleranz 1 KB)
'   - Datei-Existenz-Pruefung nach Kopie
'   - Fehler pro Datei einzeln behandelt (eine fehlgeschlagene
'     Kopie verhindert nicht die andere)
' ===========================================================================

' Erstellt ein Backup von Frontend und/oder Backend.
' Gibt True zurueck wenn alle angeforderten Backups erfolgreich waren.
'
' Beispiele:
'   BackupErstellen              ' Beides (FE + BE)
'   BackupErstellen True, False  ' Nur Frontend
'   BackupErstellen False, True  ' Nur Backend
Public Function BackupErstellen(Optional ByVal blnFrontend As Boolean = True, _
                                 Optional ByVal blnBackend As Boolean = True) As Boolean
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Backup-Ordner sicherstellen
    Dim strBackupDir As String
    strBackupDir = CurrentProject.Path & "\Backups\"
    If Not fso.FolderExists(strBackupDir) Then
        fso.CreateFolder strBackupDir
    End If

    Dim strTimestamp As String
    strTimestamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")

    Dim lngAngefordert As Long
    Dim lngErfolgreich As Long
    Dim strDetails As String

    ' --- Frontend-Backup ---
    If blnFrontend Then
        lngAngefordert = lngAngefordert + 1
        Dim strFEQuelle As String
        strFEQuelle = CurrentProject.FullName

        If fso.FileExists(strFEQuelle) Then
            Dim strFEZiel As String
            strFEZiel = strBackupDir & "FE_" & strTimestamp & "_" & _
                        fso.GetFileName(strFEQuelle)

            If BackupDatei(fso, strFEQuelle, strFEZiel) Then
                lngErfolgreich = lngErfolgreich + 1
                strDetails = strDetails & "FE -> " & strFEZiel & vbCrLf
                LogInfo "Backup FE: " & strFEZiel, "BACKUP"
            Else
                strDetails = strDetails & "FE FEHLGESCHLAGEN" & vbCrLf
            End If
        Else
            LogWarn "Frontend-Datei nicht gefunden: " & strFEQuelle, "BACKUP"
        End If
    End If

    ' --- Backend-Backup ---
    If blnBackend Then
        Dim strBEQuelle As String
        strBEQuelle = CacheGetConfig(CFG_BACKEND_PFAD, "")

        If strBEQuelle <> "" Then
            lngAngefordert = lngAngefordert + 1

            If fso.FileExists(strBEQuelle) Then
                Dim strBEZiel As String
                strBEZiel = strBackupDir & "BE_" & strTimestamp & "_" & _
                            fso.GetFileName(strBEQuelle)

                If BackupDatei(fso, strBEQuelle, strBEZiel) Then
                    lngErfolgreich = lngErfolgreich + 1
                    strDetails = strDetails & "BE -> " & strBEZiel & vbCrLf
                    LogInfo "Backup BE: " & strBEZiel, "BACKUP"
                Else
                    strDetails = strDetails & "BE FEHLGESCHLAGEN" & vbCrLf
                End If
            Else
                LogWarn "Backend nicht erreichbar: " & strBEQuelle, "BACKUP"
                strDetails = strDetails & "BE nicht erreichbar" & vbCrLf
            End If
        End If
    End If

    ' Ergebnis
    Set fso = Nothing

    If lngAngefordert = 0 Then
        BackupErstellen = False
        Exit Function
    End If

    BackupErstellen = (lngErfolgreich = lngAngefordert)

    MsgBox "Backup " & IIf(BackupErstellen, "erfolgreich", "unvollstaendig") & ":" & _
           vbCrLf & vbCrLf & strDetails, _
           IIf(BackupErstellen, vbInformation, vbExclamation), _
           "OutlookSync - Backup"
    Exit Function

ErrHandler:
    HandleError "modBackend", "BackupErstellen"
    BackupErstellen = False
End Function


' Kopiert eine einzelne Datei und prueft Groesse + Existenz.
' Gibt True zurueck bei Erfolg.
Private Function BackupDatei(ByVal fso As Object, _
                              ByVal strQuelle As String, _
                              ByVal strZiel As String) As Boolean
    On Error GoTo ErrHandler

    ' Kopieren
    fso.CopyFile strQuelle, strZiel, True

    ' Existenz pruefen
    If Not fso.FileExists(strZiel) Then
        LogError "Backup-Datei nach Kopie nicht gefunden: " & strZiel, "BACKUP"
        BackupDatei = False
        Exit Function
    End If

    ' Groessenvergleich (Toleranz 1 KB fuer Jet Locking-Bytes)
    Dim dblQuellGroesse As Double
    Dim dblZielGroesse As Double
    dblQuellGroesse = CDbl(fso.GetFile(strQuelle).Size)
    dblZielGroesse = CDbl(fso.GetFile(strZiel).Size)

    If Abs(dblZielGroesse - dblQuellGroesse) > 1024 Then
        LogWarn "Backup Groessenabweichung: Quelle=" & dblQuellGroesse & _
                " Ziel=" & dblZielGroesse & " (" & strZiel & ")", "BACKUP"
    End If

    BackupDatei = True
    Exit Function

ErrHandler:
    LogError "Backup fehlgeschlagen: " & strQuelle & " -> " & strZiel & _
             " [" & Err.Number & "] " & Err.Description, "BACKUP"
    BackupDatei = False
End Function


' ===========================================================================
' DEV/LIVE-SWITCH (v0.5.3)
' ===========================================================================
' Ermoeglicht einfaches Umschalten zwischen:
'   - LIVE-Modus:  Backend auf Netzlaufwerk (Produktionsdaten)
'   - DEV-Modus:   Lokale Kopie des Backends (fuer Tests/Entwicklung)
'
' Der Switch:
'   1. WechsleZuDevModus  -> Kopiert Live-BE nach lokal, relinkt
'   2. WechsleZuLiveModus -> Relinkt zurueck auf Netzwerk-BE
'
' Der aktuelle Modus wird in tblConfig (CFG_DEV_MODUS) gespeichert.
' ===========================================================================

' Prueft ob aktuell im Dev-Modus gearbeitet wird
Public Function IstDevModus() As Boolean
    IstDevModus = (CacheGetConfig(CFG_DEV_MODUS, "0") = "1")
End Function


' Wechselt in den lokalen Dev-Modus:
'   1. Optional: Backup erstellen
'   2. Live-Backend nach lokal kopieren
'   3. Alle Backend-Tabellen auf lokale Kopie relinken
Public Sub WechsleZuDevModus()
    On Error GoTo ErrHandler

    ' Schon im Dev-Modus?
    If IstDevModus() Then
        MsgBox "Bereits im Dev-Modus!" & vbCrLf & vbCrLf & _
               "Lokales Backend: " & DevBackendPfad(), _
               vbInformation, "OutlookSync"
        Exit Sub
    End If

    ' Live-Pfad ermitteln
    Dim strLivePfad As String
    strLivePfad = CacheGetConfig(CFG_BACKEND_PFAD, "")

    If strLivePfad = "" Then
        MsgBox "Kein Backend konfiguriert." & vbCrLf & _
               "Im lokalen Modus gibt es keinen Dev-Switch.", _
               vbExclamation, "OutlookSync"
        Exit Sub
    End If

    ' Bestaeigung
    Dim strDevPfad As String
    strDevPfad = DevBackendPfad()

    If MsgBox("In den DEV-MODUS wechseln?" & vbCrLf & vbCrLf & _
              "Live-Backend: " & strLivePfad & vbCrLf & _
              "Lokale Kopie: " & strDevPfad & vbCrLf & vbCrLf & _
              "Das Live-Backend wird kopiert und die Tabellen " & _
              "auf die lokale Kopie umgelinkt." & vbCrLf & vbCrLf & _
              "Fortfahren?", _
              vbYesNo + vbQuestion, "OutlookSync - Dev-Modus") = vbNo Then
        Exit Sub
    End If

    ' Optional Backup vorher
    If MsgBox("Vorher ein Backup erstellen?", _
              vbYesNo + vbQuestion, "Backup") = vbYes Then
        If Not BackupErstellen(True, True) Then
            If MsgBox("Backup fehlgeschlagen. Trotzdem fortfahren?", _
                      vbYesNo + vbExclamation, "Warnung") = vbNo Then
                Exit Sub
            End If
        End If
    End If

    ' Live-Backend erreichbar?
    If Dir(strLivePfad) = "" Then
        MsgBox "Live-Backend nicht erreichbar:" & vbCrLf & strLivePfad, _
               vbCritical, "OutlookSync"
        Exit Sub
    End If

    ' Kopieren
    Debug.Print "DEV-SWITCH: Kopiere " & strLivePfad & " -> " & strDevPfad
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile strLivePfad, strDevPfad, True
    Set fso = Nothing

    If Dir(strDevPfad) = "" Then
        LogError "Dev-Kopie nicht erstellt: " & strDevPfad, "DEV-SWITCH"
        MsgBox "Kopie fehlgeschlagen!", vbCritical, "OutlookSync"
        Exit Sub
    End If

    LogInfo "Live-Backend nach lokal kopiert: " & strDevPfad, "DEV-SWITCH"

    ' Relinken auf lokale Kopie
    Dim lngOK As Long, lngFail As Long
    RelinkBackendTabellen strDevPfad, lngOK, lngFail

    If lngFail > 0 Then
        LogWarn "Dev-Relink: " & lngFail & " Tabelle(n) fehlgeschlagen", "DEV-SWITCH"
        MsgBox "Dev-Modus teilweise aktiviert." & vbCrLf & _
               lngFail & " Tabelle(n) konnten nicht umgelinkt werden.", _
               vbExclamation, "OutlookSync"
    Else
        ' Modus speichern
        CacheSetConfig CFG_DEV_MODUS, "1"
        LogInfo "Dev-Modus aktiviert (" & lngOK & " Tabellen relinkt)", "DEV-SWITCH"
        MsgBox "DEV-MODUS aktiviert!" & vbCrLf & vbCrLf & _
               lngOK & " Tabellen auf lokale Kopie umgelinkt." & vbCrLf & _
               "Backend: " & strDevPfad, _
               vbInformation, "OutlookSync"
    End If
    Exit Sub

ErrHandler:
    HandleError "modBackend", "WechsleZuDevModus"
End Sub


' Wechselt zurueck auf das Live-Backend (Netzlaufwerk).
Public Sub WechsleZuLiveModus()
    On Error GoTo ErrHandler

    ' Schon im Live-Modus?
    If Not IstDevModus() Then
        MsgBox "Bereits im Live-Modus!", vbInformation, "OutlookSync"
        Exit Sub
    End If

    Dim strLivePfad As String
    strLivePfad = CacheGetConfig(CFG_BACKEND_PFAD, "")

    If strLivePfad = "" Then
        ' Sollte nicht passieren, aber Sicherheit
        CacheSetConfig CFG_DEV_MODUS, "0"
        MsgBox "Kein Live-Backend konfiguriert. Dev-Flag zurueckgesetzt.", _
               vbExclamation, "OutlookSync"
        Exit Sub
    End If

    ' Bestaeigung
    If MsgBox("Zurueck zum LIVE-MODUS wechseln?" & vbCrLf & vbCrLf & _
              "ACHTUNG: Aenderungen an der lokalen Dev-Datenbank " & _
              "werden NICHT ins Live-Backend uebertragen!" & vbCrLf & vbCrLf & _
              "Live-Backend: " & strLivePfad & vbCrLf & vbCrLf & _
              "Fortfahren?", _
              vbYesNo + vbExclamation, "OutlookSync - Live-Modus") = vbNo Then
        Exit Sub
    End If

    ' Live erreichbar?
    If Dir(strLivePfad) = "" Then
        MsgBox "Live-Backend nicht erreichbar:" & vbCrLf & strLivePfad & vbCrLf & vbCrLf & _
               "Bitte Netzwerkverbindung pruefen.", _
               vbCritical, "OutlookSync"
        Exit Sub
    End If

    ' Relinken auf Live
    Dim lngOK As Long, lngFail As Long
    RelinkBackendTabellen strLivePfad, lngOK, lngFail

    If lngFail > 0 Then
        LogWarn "Live-Relink: " & lngFail & " Tabelle(n) fehlgeschlagen", "DEV-SWITCH"
        MsgBox "Umschaltung teilweise fehlgeschlagen." & vbCrLf & _
               lngFail & " Tabelle(n) konnten nicht umgelinkt werden.", _
               vbExclamation, "OutlookSync"
    Else
        CacheSetConfig CFG_DEV_MODUS, "0"
        LogInfo "Live-Modus aktiviert (" & lngOK & " Tabellen relinkt)", "DEV-SWITCH"
        MsgBox "LIVE-MODUS aktiviert!" & vbCrLf & vbCrLf & _
               lngOK & " Tabellen auf Live-Backend umgelinkt." & vbCrLf & _
               "Backend: " & strLivePfad, _
               vbInformation, "OutlookSync"
    End If
    Exit Sub

ErrHandler:
    HandleError "modBackend", "WechsleZuLiveModus"
End Sub


' Gibt den Pfad zur lokalen Dev-Kopie des Backends zurueck
Private Function DevBackendPfad() As String
    DevBackendPfad = CurrentProject.Path & "\" & BE_DEV_DATEINAME
End Function


' ===========================================================================
' RELINK MIT FORTSCHRITT + FEHLERSAMMLUNG (v0.5.3)
' ===========================================================================
' Ersetzt den einfachen RefreshLink() durch eine robuste Variante:
'   - Zaehlt OK/Fail separat (ByRef-Rueckgabe)
'   - Debug.Print Fortschritt pro Tabelle
'   - Sammelt Fehlertabellen fuer Zusammenfassung
'   - Einzelfehler brechen nicht den ganzen Lauf ab
' ===========================================================================

' Relinkt alle Backend-Tabellen auf einen neuen Pfad.
' lngOK/lngFail geben die Ergebnis-Zaehler zurueck.
Public Sub RelinkBackendTabellen(ByVal strZielPfad As String, _
                                  ByRef lngOK As Long, _
                                  ByRef lngFail As Long)
    On Error GoTo ErrHandler

    Dim arrTabellen As Variant
    arrTabellen = GetBackendTabellen()
    Dim lngGesamt As Long
    lngGesamt = UBound(arrTabellen) - LBound(arrTabellen) + 1

    lngOK = 0
    lngFail = 0

    Debug.Print String(60, "-")
    Debug.Print "RELINK -> " & strZielPfad
    Debug.Print String(60, "-")

    Dim db As DAO.Database
    Set db = CurrentDb

    Dim i As Long
    For i = LBound(arrTabellen) To UBound(arrTabellen)
        Dim strTabelle As String
        strTabelle = CStr(arrTabellen(i))

        Debug.Print "  [" & (i - LBound(arrTabellen) + 1) & "/" & lngGesamt & "] " & strTabelle & "...";

        ' Pruefe ob Tabelle existiert und verlinkt ist
        On Error Resume Next
        Dim td As DAO.TableDef
        Set td = db.TableDefs(strTabelle)

        If Err.Number <> 0 Then
            ' Tabelle nicht vorhanden -> neuen Link erstellen
            Err.Clear
            On Error GoTo ErrHandler

            Set td = db.CreateTableDef(strTabelle)
            td.Connect = ";DATABASE=" & strZielPfad
            td.SourceTableName = strTabelle
            db.TableDefs.Append td
            db.TableDefs.Refresh

            Debug.Print " NEU ERSTELLT"
            lngOK = lngOK + 1
            LogInfo "Relink: Neuer Link " & strTabelle & " -> " & strZielPfad, "BACKEND"
            GoTo WeiterNaechste
        End If

        ' Tabelle existiert -> Link aktualisieren
        Err.Clear
        td.Connect = ";DATABASE=" & strZielPfad
        td.RefreshLink

        If Err.Number <> 0 Then
            Dim lngErr As Long, strErr As String
            lngErr = Err.Number
            strErr = Err.Description
            Err.Clear
            On Error GoTo ErrHandler

            Debug.Print " FEHLER [" & lngErr & "] " & strErr
            lngFail = lngFail + 1
            LogWarn "Relink fehlgeschlagen: " & strTabelle & " [" & lngErr & "] " & strErr, "BACKEND"
            GoTo WeiterNaechste
        End If

        On Error GoTo ErrHandler
        Debug.Print " OK"
        lngOK = lngOK + 1

WeiterNaechste:
        Set td = Nothing
    Next i

    Set db = Nothing

    Debug.Print String(60, "-")
    Debug.Print "RELINK: " & lngOK & " OK, " & lngFail & " Fehler"
    Debug.Print String(60, "-")
    Exit Sub

ErrHandler:
    HandleError "modBackend", "RelinkBackendTabellen"
End Sub


' ===========================================================================
' STATUS-ERWEITERUNG (v0.5.3)
' ===========================================================================

' Erweiterte Statusausgabe: Backend-Modus + Tabellen-Link-Details
Public Sub BackendInfoReport()
    On Error Resume Next

    Debug.Print String(60, "=")
    Debug.Print "=== BACKEND INFO-REPORT ==="
    Debug.Print String(60, "=")

    ' Modus
    Debug.Print "  Modus:     " & IIf(IstDevModus(), "DEV (lokal)", "LIVE")
    Debug.Print "  Status:    " & BackendStatus()
    Debug.Print "  Offline:   " & IIf(g_blnBackendOffline, "JA (!)", "Nein")
    Debug.Print "  Watchdog:  " & IIf(m_blnWatchdogAktiv Or m_blnFormWatchdogAktiv, "Aktiv", "Inaktiv")

    ' Tabellen-Links
    Debug.Print ""
    Debug.Print "  --- Backend-Tabellen ---"
    Dim arrBE As Variant
    arrBE = GetBackendTabellen()
    Dim i As Long
    For i = LBound(arrBE) To UBound(arrBE)
        Dim strTbl As String
        strTbl = CStr(arrBE(i))
        Dim strInfo As String

        If Not TabelleExistiert(strTbl) Then
            strInfo = "FEHLT"
        ElseIf IstLinkedTable(strTbl) Then
            ' Link-Ziel anzeigen
            Dim td As DAO.TableDef
            Set td = CurrentDb.TableDefs(strTbl)
            If Len(td.Connect) > 10 Then
                strInfo = "LINK -> " & Mid(td.Connect, 11)
            Else
                strInfo = "LINK (unbekannt)"
            End If
        Else
            strInfo = "LOKAL"
        End If

        Debug.Print "  " & Left$(strTbl & String(28, " "), 28) & strInfo
    Next i

    ' Frontend-Tabellen
    Debug.Print ""
    Debug.Print "  --- Frontend-Tabellen ---"
    Dim arrFE As Variant
    arrFE = GetFrontendTabellen()
    For i = LBound(arrFE) To UBound(arrFE)
        strTbl = CStr(arrFE(i))
        If TabelleExistiert(strTbl) Then
            strInfo = "OK (lokal)"
        Else
            strInfo = "FEHLT"
        End If
        Debug.Print "  " & Left$(strTbl & String(28, " "), 28) & strInfo
    Next i

    ' Dev-Modus Info
    If IstDevModus() Then
        Debug.Print ""
        Debug.Print "  --- Dev-Modus ---"
        Debug.Print "  Lokales Backend: " & DevBackendPfad()
        Debug.Print "  Existiert:       " & IIf(Dir(DevBackendPfad()) <> "", "Ja", "Nein")
    End If

    ' Backup-Ordner
    Dim strBackupDir As String
    strBackupDir = CurrentProject.Path & "\Backups\"
    If Dir(strBackupDir, vbDirectory) <> "" Then
        Dim lngBackups As Long
        Dim strDatei As String
        strDatei = Dir(strBackupDir & "*.accdb")
        Do While strDatei <> ""
            lngBackups = lngBackups + 1
            strDatei = Dir()
        Loop
        Debug.Print ""
        Debug.Print "  Backups:   " & lngBackups & " Dateien in " & strBackupDir
    End If

    Debug.Print String(60, "=")
    On Error GoTo 0
End Sub


