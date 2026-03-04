Attribute VB_Name = "modBackend"
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
'   tblEmailStatus
'
' WICHTIG: Wenn kein Backend konfiguriert ist, arbeitet alles lokal.
'          Die Verknuepfung ist OPTIONAL und kann jederzeit hergestellt werden.
'
' Aufruf:
'   VerknuepfeBackend "\\Server\Share\OutlookSync_BE.accdb"
'   TrenneBackend
'   ? IstBackendVerfuegbar()
'   ? BackendStatus()
'
' Abhaengigkeiten: modSchema (TabelleExistiert, ErstelleAlleTabellen),
'                  modStringUtils (NormalisierePfad, ErstelleOrdner),
'                  modLogging
' ===========================================================================


' ---------------------------------------------------------------------------
' TABELLEN-KLASSIFIKATION
' ---------------------------------------------------------------------------

' Backend-Tabellen (werden auf Netzlaufwerk verknuepft)
Private Function GetBackendTabellen() As Variant
    GetBackendTabellen = Array("tblSyncLauf", "tblKontakte", "tblOutlookOrdner", _
                                "tblEmailThreads", "tblEmails", "tblEmailContent", _
                                "tblEmailEmpfaenger", "tblEmailAnhaenge", "tblEmailStatus")
End Function

' Frontend-Tabellen (bleiben immer lokal)
Private Function GetFrontendTabellen() As Variant
    GetFrontendTabellen = Array("tblConfig", "tblSyncProfil", "tblSyncProfilOrdner")
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

    ' 4. Pfad in Config speichern
    SchreibeConfig "BackendPfad", strBackendPfad

    Debug.Print "  Verknuepft: " & lngOK & " OK, " & lngFail & " Fehler"
    Debug.Print String(70, "=")

    LogInfo "Backend verknuepft: " & strBackendPfad & " (" & lngOK & " Tabellen)", "BACKEND"
    VerknuepfeBackend = (lngFail = 0)
    Exit Function

ErrHandler:
    LogVBAError "VerknuepfeBackend"
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

    ' Backend-Pfad aus Config entfernen
    SchreibeConfig "BackendPfad", ""

    ' Lokale Tabellen neu erstellen
    Debug.Print "  Erstelle lokale Tabellen..."
    ErstelleAlleTabellen

    Debug.Print "=== Backend getrennt ==="
    Debug.Print String(70, "=")

    LogInfo "Backend getrennt, lokale Tabellen wiederhergestellt", "BACKEND"
    Exit Sub

ErrHandler:
    LogVBAError "TrenneBackend"
End Sub


' ===========================================================================
' BACKEND-STATUS PRUEFEN
' ===========================================================================

' Prueft ob das konfigurierte Backend erreichbar ist
' Wenn kein Backend konfiguriert: TRUE (lokal = immer verfuegbar)
Public Function IstBackendVerfuegbar() As Boolean
    On Error Resume Next
    Dim strPfad As String
    strPfad = LeseConfig("BackendPfad", "")

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
    strPfad = LeseConfig("BackendPfad", "")

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


' Prueft ob aktuell ein Backend verknuepft ist
Public Function IstBackendVerknuepft() As Boolean
    IstBackendVerknuepft = (LeseConfig("BackendPfad", "") <> "")
End Function


' Gibt den konfigurierten Backend-Pfad zurueck (leer wenn lokal)
Public Function GetBackendPfad() As String
    GetBackendPfad = LeseConfig("BackendPfad", "")
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
' (Linked Table Links refreshen)
Public Function ReconnectBackend() As Boolean
    On Error GoTo ErrHandler

    Dim strPfad As String
    strPfad = LeseConfig("BackendPfad", "")

    If strPfad = "" Then
        ReconnectBackend = True  ' Lokal = immer OK
        Exit Function
    End If

    If Not PruefeDatei(strPfad) Then
        LogError "Backend nicht erreichbar fuer Reconnect: " & strPfad, "BACKEND"
        ReconnectBackend = False
        Exit Function
    End If

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
    LogVBAError "ReconnectBackend"
    ReconnectBackend = False
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
Private Function ErstelleBackendDB(ByVal strPfad As String) As Boolean
    On Error GoTo ErrHandler

    ' Verzeichnis erstellen
    Dim strDir As String
    strDir = Left(strPfad, InStrRev(strPfad, "\"))
    If strDir <> "" Then ErstelleOrdner strDir

    ' Neue Access-DB erstellen
    Dim dbNew As DAO.Database
    Set dbNew = DBEngine.CreateDatabase(strPfad, dbLangGeneral)

    ' --- tblSyncLauf ---
    dbNew.Execute _
        "CREATE TABLE tblSyncLauf (" & _
        "  SyncLaufID AUTOINCREMENT CONSTRAINT PK_SyncLauf PRIMARY KEY," & _
        "  StartZeit DATETIME, EndeZeit DATETIME," & _
        "  Status TEXT(20) DEFAULT 'Gestartet'," & _
        "  AnzahlGelesen LONG DEFAULT 0, AnzahlNeu LONG DEFAULT 0," & _
        "  AnzahlDuplikate LONG DEFAULT 0, AnzahlFehler LONG DEFAULT 0," & _
        "  OrdnerPfad TEXT(255), Projekt TEXT(100), Phase TEXT(100))"

    ' --- tblKontakte ---
    dbNew.Execute _
        "CREATE TABLE tblKontakte (" & _
        "  KontaktID AUTOINCREMENT CONSTRAINT PK_Kontakte PRIMARY KEY," & _
        "  Anzeigename TEXT(255), Email TEXT(255), EmailTyp TEXT(10)," & _
        "  Vorname TEXT(100), Nachname TEXT(100), Titel TEXT(50)," & _
        "  Namenszusatz TEXT(100), Institution TEXT(255)," & _
        "  Sortiername TEXT(255), KontaktTyp TEXT(20)," & _
        "  ErstelltAm DATETIME, AktualisiertAm DATETIME)"
    dbNew.Execute "CREATE INDEX idx_Kontakte_Email ON tblKontakte (Email)"

    ' --- tblOutlookOrdner ---
    dbNew.Execute _
        "CREATE TABLE tblOutlookOrdner (" & _
        "  OrdnerID AUTOINCREMENT CONSTRAINT PK_Ordner PRIMARY KEY," & _
        "  OrdnerName TEXT(255), OrdnerPfad TEXT(255)," & _
        "  ParentID LONG DEFAULT 0, PostfachName TEXT(255)," & _
        "  StoreID TEXT(255)," & _
        "  ElementAnzahl LONG DEFAULT 0, LetzterSync DATETIME)"
    dbNew.Execute "CREATE UNIQUE INDEX idx_Ordner_Pfad ON tblOutlookOrdner (OrdnerPfad)"

    ' --- tblEmailThreads ---
    dbNew.Execute _
        "CREATE TABLE tblEmailThreads (" & _
        "  ThreadID AUTOINCREMENT CONSTRAINT PK_Threads PRIMARY KEY," & _
        "  ThreadBetreff TEXT(255), ThreadIdentifier TEXT(255)," & _
        "  Antwortanzahl LONG DEFAULT 1, ErsterAbsender TEXT(255)," & _
        "  ErstesMailDatum DATETIME, LetztesMailDatum DATETIME," & _
        "  ErstelltAm DATETIME)"
    dbNew.Execute "CREATE UNIQUE INDEX idx_Thread_Ident ON tblEmailThreads (ThreadIdentifier)"

    ' --- tblEmails ---
    dbNew.Execute _
        "CREATE TABLE tblEmails (" & _
        "  EmailID AUTOINCREMENT CONSTRAINT PK_Emails PRIMARY KEY," & _
        "  OutlookEntryID TEXT(255), UniqueHash TEXT(64)," & _
        "  ThreadID LONG DEFAULT 0, OrdnerID LONG DEFAULT 0," & _
        "  KontaktID_Absender LONG DEFAULT 0, SyncLaufID LONG DEFAULT 0," & _
        "  Betreff TEXT(255), BetreffBereinigt TEXT(255)," & _
        "  AbsenderName TEXT(255), AbsenderEmail TEXT(255)," & _
        "  EmpfangenAm DATETIME, GesendetAm DATETIME," & _
        "  Groesse LONG DEFAULT 0, Wichtigkeit SHORT DEFAULT 1," & _
        "  Gelesen YESNO, HatAnhaenge YESNO, AnhangAnzahl SHORT DEFAULT 0," & _
        "  MessageClass TEXT(50), InternetMessageID TEXT(255)," & _
        "  MSGDateiPfad TEXT(255), Status TEXT(20) DEFAULT 'Neu'," & _
        "  ErstelltAm DATETIME)"
    dbNew.Execute "CREATE UNIQUE INDEX idx_Email_Hash ON tblEmails (UniqueHash)"
    dbNew.Execute "CREATE INDEX idx_Email_EntryID ON tblEmails (OutlookEntryID)"
    dbNew.Execute "CREATE INDEX idx_Email_ThreadID ON tblEmails (ThreadID)"
    dbNew.Execute "CREATE INDEX idx_Email_KontaktID ON tblEmails (KontaktID_Absender)"
    dbNew.Execute "CREATE INDEX idx_Email_OrdnerID ON tblEmails (OrdnerID)"
    dbNew.Execute "CREATE INDEX idx_Email_SyncLauf ON tblEmails (SyncLaufID)"
    dbNew.Execute "CREATE INDEX idx_Email_Datum ON tblEmails (EmpfangenAm)"

    ' --- tblEmailContent ---
    dbNew.Execute _
        "CREATE TABLE tblEmailContent (" & _
        "  ContentID AUTOINCREMENT CONSTRAINT PK_Content PRIMARY KEY," & _
        "  EmailID LONG NOT NULL, HTMLBody MEMO, PlainTextBody MEMO," & _
        "  HatHTML YESNO, GroesseHTML LONG DEFAULT 0, GroesseText LONG DEFAULT 0)"
    dbNew.Execute "CREATE UNIQUE INDEX idx_Content_EmailID ON tblEmailContent (EmailID)"

    ' --- tblEmailEmpfaenger ---
    dbNew.Execute _
        "CREATE TABLE tblEmailEmpfaenger (" & _
        "  EmpfaengerID AUTOINCREMENT CONSTRAINT PK_Empfaenger PRIMARY KEY," & _
        "  EmailID LONG NOT NULL, KontaktID LONG DEFAULT 0," & _
        "  Typ TEXT(5), Anzeigename TEXT(255), Email TEXT(255))"
    dbNew.Execute "CREATE INDEX idx_Empf_EmailID ON tblEmailEmpfaenger (EmailID)"

    ' --- tblEmailAnhaenge ---
    dbNew.Execute _
        "CREATE TABLE tblEmailAnhaenge (" & _
        "  AnhangID AUTOINCREMENT CONSTRAINT PK_Anhaenge PRIMARY KEY," & _
        "  EmailID LONG NOT NULL, Dateiname TEXT(255)," & _
        "  DateinameBereinigt TEXT(255), Erweiterung TEXT(20)," & _
        "  Groesse LONG DEFAULT 0, MimeType TEXT(100)," & _
        "  AnhangTyp SHORT DEFAULT 1, IstVersteckt YESNO," & _
        "  IstGespeichert YESNO, DateiPfad TEXT(255), ErstelltAm DATETIME)"
    dbNew.Execute "CREATE INDEX idx_Anh_EmailID ON tblEmailAnhaenge (EmailID)"

    ' --- tblEmailStatus ---
    dbNew.Execute _
        "CREATE TABLE tblEmailStatus (" & _
        "  StatusID AUTOINCREMENT CONSTRAINT PK_EmailStatus PRIMARY KEY," & _
        "  EmailID LONG NOT NULL, Status TEXT(50)," & _
        "  GeaendertVon TEXT(100), Bemerkung TEXT(255), GeaendertAm DATETIME)"
    dbNew.Execute "CREATE INDEX idx_Status_EmailID ON tblEmailStatus (EmailID)"

    dbNew.Close
    Set dbNew = Nothing

    Debug.Print "  [OK  ] Backend-DB erstellt: " & strPfad
    LogInfo "Backend-DB erstellt: " & strPfad, "BACKEND"
    ErstelleBackendDB = True
    Exit Function

ErrHandler:
    Debug.Print "  [FAIL] Backend-DB Erstellung: " & Err.Description
    LogVBAError "ErstelleBackendDB"
    ErstelleBackendDB = False
End Function


' Prueft ob eine Datei existiert und zugreifbar ist
Private Function PruefeDatei(ByVal strPfad As String) As Boolean
    On Error Resume Next
    PruefeDatei = (Dir(strPfad) <> "")
    If Err.Number <> 0 Then PruefeDatei = False: Err.Clear
    On Error GoTo 0
End Function
