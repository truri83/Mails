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
' Abhaengigkeiten: modSchema (Tabellenkonstanten, ErstelleAlleTabellen),
'                  modDDL (DDL_TabelleExistiert, DDL_IstVerknuepft),
'                  modStringUtils (ErstelleOrdner), modLogging
' ===========================================================================


' ---------------------------------------------------------------------------
' TABELLEN-KLASSIFIKATION
' ---------------------------------------------------------------------------

Private Function GetBackendTabellen() As Variant
    GetBackendTabellen = Array(TBL_SYNC_LAUF, TBL_KONTAKTE, TBL_ORDNER, TBL_THREADS, TBL_EMAILS, TBL_CONTENT, TBL_EMPFAENGER, TBL_ANHAENGE, TBL_EMAIL_STATUS)
End Function

Private Function GetFrontendTabellen() As Variant
    GetFrontendTabellen = Array(TBL_CONFIG, TBL_PROFIL, TBL_PROFIL_ORDNER)
End Function


' ===========================================================================
' BACKEND VERKNUEPFEN
' ===========================================================================

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

Public Sub TrenneBackend()
    On Error GoTo ErrHandler

    Dim arrTabellen As Variant
    Dim i As Long
    arrTabellen = GetBackendTabellen()

    Debug.Print String(70, "=")
    Debug.Print "=== BACKEND TRENNEN ==="

    For i = LBound(arrTabellen) To UBound(arrTabellen)
        EntferneVerknuepfung CStr(arrTabellen(i))
    Next i

    SchreibeConfig "BackendPfad", ""

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

Public Function IstBackendVerfuegbar() As Boolean
    On Error Resume Next
    Dim strPfad As String
    strPfad = LeseConfig("BackendPfad", "")

    If strPfad = "" Then
        IstBackendVerfuegbar = True
        Exit Function
    End If

    IstBackendVerfuegbar = PruefeDatei(strPfad)
    On Error GoTo 0
End Function

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

Public Function IstBackendVerknuepft() As Boolean
    IstBackendVerknuepft = (LeseConfig("BackendPfad", "") <> "")
End Function

Public Function GetBackendPfad() As String
    GetBackendPfad = LeseConfig("BackendPfad", "")
End Function

Public Function IstFETabelle(ByVal strTabelle As String) As Boolean
    Dim arr As Variant, v As Variant
    arr = GetFrontendTabellen()
    For Each v In arr
        If StrComp(CStr(v), strTabelle, vbTextCompare) = 0 Then
            IstFETabelle = True
            Exit Function
        End If
    Next v
    IstFETabelle = False
End Function

Public Function IstBETabelle(ByVal strTabelle As String) As Boolean
    Dim arr As Variant, v As Variant
    arr = GetBackendTabellen()
    For Each v In arr
        If StrComp(CStr(v), strTabelle, vbTextCompare) = 0 Then
            IstBETabelle = True
            Exit Function
        End If
    Next v
    IstBETabelle = False
End Function


' ===========================================================================
' RECONNECT (nach Netzwerk-Unterbrechung)
' ===========================================================================

Public Function ReconnectBackend() As Boolean
    On Error GoTo ErrHandler

    Dim strPfad As String
    strPfad = LeseConfig("BackendPfad", "")

    If strPfad = "" Then
        ReconnectBackend = True
        Exit Function
    End If

    If Not PruefeDatei(strPfad) Then
        LogError "Backend nicht erreichbar fuer Reconnect: " & strPfad, "BACKEND"
        ReconnectBackend = False
        Exit Function
    End If

    Dim arrTabellen As Variant, i As Long
    arrTabellen = GetBackendTabellen()

    For i = LBound(arrTabellen) To UBound(arrTabellen)
        RefreshLink CStr(arrTabellen(i)), strPfad
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

Private Function VerknuepfeTabelle(ByVal strTabelle As String, ByVal strBackendDB As String) As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    If DDL_TabelleExistiert(strTabelle) Then
        Dim td As DAO.TableDef
        Set td = db.TableDefs(strTabelle)

        If td.Connect = "" Then
            MigriereTabellenDaten strTabelle, strBackendDB
            db.Execute "DROP TABLE [" & strTabelle & "]"
            db.TableDefs.Refresh
            Debug.Print "  [MIGRIERT] " & strTabelle
        Else
            EntferneVerknuepfung strTabelle
        End If
    End If

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

Private Sub MigriereTabellenDaten(ByVal strTabelle As String, ByVal strBackendDB As String)
    On Error GoTo ErrHandler

    Dim lngCount As Long
    lngCount = DCount("*", strTabelle)
    If lngCount = 0 Then Exit Sub

    Dim dbBE As DAO.Database
    Set dbBE = DBEngine.OpenDatabase(strBackendDB)

    Dim lngBECount As Long
    On Error Resume Next
    lngBECount = dbBE.OpenRecordset("SELECT COUNT(*) FROM [" & strTabelle & "]").Fields(0)
    If Err.Number <> 0 Then lngBECount = 0: Err.Clear
    On Error GoTo ErrHandler

    If lngBECount = 0 Then
        CurrentDb.Execute "INSERT INTO [" & strTabelle & "] IN '" & strBackendDB & "' SELECT * FROM [" & strTabelle & "]"
        LogInfo "Migriert: " & lngCount & " Datensaetze von " & strTabelle, "BACKEND"
        Debug.Print "    -> " & lngCount & " Datensaetze migriert"
    Else
        LogWarn "Backend-Tabelle " & strTabelle & " hat bereits " & lngBECount & " Datensaetze, Migration uebersprungen", "BACKEND"
    End If

    dbBE.Close: Set dbBE = Nothing
    Exit Sub

ErrHandler:
    LogWarn "Datenmigration fehlgeschlagen fuer " & strTabelle & ": " & Err.Description, "BACKEND"
End Sub

Private Sub EntferneVerknuepfung(ByVal strTabelle As String)
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb

    If DDL_TabelleExistiert(strTabelle) Then
        db.TableDefs.Delete strTabelle
        db.TableDefs.Refresh
        Debug.Print "  [UNLINK] " & strTabelle
    End If

    Set db = Nothing
    On Error GoTo 0
End Sub

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
' BACKEND-DB ERSTELLEN
' ===========================================================================

' Erstellt eine neue Backend-Datenbank mit allen BE-Tabellen.
' Nutzt die SQL-Definitionen aus modSchema (kein doppeltes SQL).
Private Function ErstelleBackendDB(ByVal strPfad As String) As Boolean
    On Error GoTo ErrHandler

    ' Verzeichnis erstellen
    Dim strDir As String
    strDir = Left(strPfad, InStrRev(strPfad, "\"))
    If strDir <> "" Then ErstelleOrdner strDir

    ' Neue Access-DB erstellen
    Dim dbNew As DAO.Database
    Set dbNew = DBEngine.CreateDatabase(strPfad, dbLangGeneral)

    ' Backend-Tabellen in der neuen DB erstellen
    ' (Gleiche SQL-Strings wie in modSchema, aber gegen dbNew ausgefuehrt)
    Dim sql As String

    ' --- tblSyncLauf ---
    sql = "CREATE TABLE " & TBL_SYNC_LAUF & " ("
    sql = sql & "SyncLaufID AUTOINCREMENT CONSTRAINT PK_SyncLauf PRIMARY KEY, "
    sql = sql & "StartZeit DATETIME, EndeZeit DATETIME, Status TEXT(20), "
    sql = sql & "AnzahlGelesen LONG, AnzahlNeu LONG, AnzahlDuplikate LONG, AnzahlFehler LONG, "
    sql = sql & "OrdnerPfad TEXT(255), Projekt TEXT(100), Phase TEXT(100))"
    dbNew.Execute sql

    ' --- tblKontakte ---
    sql = "CREATE TABLE " & TBL_KONTAKTE & " ("
    sql = sql & "KontaktID AUTOINCREMENT CONSTRAINT PK_Kontakte PRIMARY KEY, "
    sql = sql & "Anzeigename TEXT(255), Email TEXT(255), EmailTyp TEXT(10), "
    sql = sql & "Vorname TEXT(100), Nachname TEXT(100), Titel TEXT(50), "
    sql = sql & "Namenszusatz TEXT(100), Institution TEXT(255), Sortiername TEXT(255), "
    sql = sql & "KontaktTyp TEXT(20), ErstelltAm DATETIME, AktualisiertAm DATETIME)"
    dbNew.Execute sql
    dbNew.Execute "CREATE INDEX idx_Kontakte_Email ON " & TBL_KONTAKTE & " (Email)"

    ' --- tblOutlookOrdner ---
    sql = "CREATE TABLE " & TBL_ORDNER & " ("
    sql = sql & "OrdnerID AUTOINCREMENT CONSTRAINT PK_Ordner PRIMARY KEY, "
    sql = sql & "OrdnerName TEXT(255), OrdnerPfad TEXT(255), ParentID LONG, "
    sql = sql & "PostfachName TEXT(255), StoreID TEXT(255), ElementAnzahl LONG, LetzterSync DATETIME)"
    dbNew.Execute sql
    dbNew.Execute "CREATE UNIQUE INDEX idx_Ordner_Pfad ON " & TBL_ORDNER & " (OrdnerPfad)"

    ' --- tblEmailThreads ---
    sql = "CREATE TABLE " & TBL_THREADS & " ("
    sql = sql & "ThreadID AUTOINCREMENT CONSTRAINT PK_Threads PRIMARY KEY, "
    sql = sql & "ThreadBetreff TEXT(255), ThreadIdentifier TEXT(255), Antwortanzahl LONG, "
    sql = sql & "ErsterAbsender TEXT(255), ErstesMailDatum DATETIME, LetztesMailDatum DATETIME, "
    sql = sql & "ErstelltAm DATETIME)"
    dbNew.Execute sql
    dbNew.Execute "CREATE UNIQUE INDEX idx_Thread_Ident ON " & TBL_THREADS & " (ThreadIdentifier)"

    ' --- tblEmails ---
    sql = "CREATE TABLE " & TBL_EMAILS & " ("
    sql = sql & "EmailID AUTOINCREMENT CONSTRAINT PK_Emails PRIMARY KEY, "
    sql = sql & "OutlookEntryID TEXT(255), UniqueHash TEXT(64), "
    sql = sql & "ThreadID LONG, OrdnerID LONG, KontaktID_Absender LONG, SyncLaufID LONG, "
    sql = sql & "Betreff TEXT(255), BetreffBereinigt TEXT(255), "
    sql = sql & "AbsenderName TEXT(255), AbsenderEmail TEXT(255), "
    sql = sql & "EmpfangenAm DATETIME, GesendetAm DATETIME, "
    sql = sql & "Groesse LONG, Wichtigkeit SHORT, Gelesen YESNO, HatAnhaenge YESNO, "
    sql = sql & "AnhangAnzahl SHORT, MessageClass TEXT(50), InternetMessageID TEXT(255), "
    sql = sql & "MSGDateiPfad TEXT(255), Status TEXT(20), ErstelltAm DATETIME)"
    dbNew.Execute sql
    dbNew.Execute "CREATE UNIQUE INDEX idx_Email_Hash ON " & TBL_EMAILS & " (UniqueHash)"
    dbNew.Execute "CREATE INDEX idx_Email_EntryID ON " & TBL_EMAILS & " (OutlookEntryID)"
    dbNew.Execute "CREATE INDEX idx_Email_ThreadID ON " & TBL_EMAILS & " (ThreadID)"
    dbNew.Execute "CREATE INDEX idx_Email_KontaktID ON " & TBL_EMAILS & " (KontaktID_Absender)"
    dbNew.Execute "CREATE INDEX idx_Email_OrdnerID ON " & TBL_EMAILS & " (OrdnerID)"
    dbNew.Execute "CREATE INDEX idx_Email_SyncLauf ON " & TBL_EMAILS & " (SyncLaufID)"
    dbNew.Execute "CREATE INDEX idx_Email_Datum ON " & TBL_EMAILS & " (EmpfangenAm)"

    ' --- tblEmailContent ---
    sql = "CREATE TABLE " & TBL_CONTENT & " ("
    sql = sql & "ContentID AUTOINCREMENT CONSTRAINT PK_Content PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, HTMLBody MEMO, PlainTextBody MEMO, "
    sql = sql & "HatHTML YESNO, GroesseHTML LONG, GroesseText LONG)"
    dbNew.Execute sql
    dbNew.Execute "CREATE UNIQUE INDEX idx_Content_EmailID ON " & TBL_CONTENT & " (EmailID)"

    ' --- tblEmailEmpfaenger ---
    sql = "CREATE TABLE " & TBL_EMPFAENGER & " ("
    sql = sql & "EmpfaengerID AUTOINCREMENT CONSTRAINT PK_Empfaenger PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, KontaktID LONG, Typ TEXT(5), "
    sql = sql & "Anzeigename TEXT(255), Email TEXT(255))"
    dbNew.Execute sql
    dbNew.Execute "CREATE INDEX idx_Empf_EmailID ON " & TBL_EMPFAENGER & " (EmailID)"

    ' --- tblEmailAnhaenge ---
    sql = "CREATE TABLE " & TBL_ANHAENGE & " ("
    sql = sql & "AnhangID AUTOINCREMENT CONSTRAINT PK_Anhaenge PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, Dateiname TEXT(255), DateinameBereinigt TEXT(255), "
    sql = sql & "Erweiterung TEXT(20), Groesse LONG, MimeType TEXT(100), AnhangTyp SHORT, "
    sql = sql & "IstVersteckt YESNO, IstGespeichert YESNO, DateiPfad TEXT(255), ErstelltAm DATETIME)"
    dbNew.Execute sql
    dbNew.Execute "CREATE INDEX idx_Anh_EmailID ON " & TBL_ANHAENGE & " (EmailID)"

    ' --- tblEmailStatus ---
    sql = "CREATE TABLE " & TBL_EMAIL_STATUS & " ("
    sql = sql & "StatusID AUTOINCREMENT CONSTRAINT PK_EmailStatus PRIMARY KEY, "
    sql = sql & "EmailID LONG NOT NULL, Status TEXT(50), GeaendertVon TEXT(100), "
    sql = sql & "Bemerkung TEXT(255), GeaendertAm DATETIME)"
    dbNew.Execute sql
    dbNew.Execute "CREATE INDEX idx_Status_EmailID ON " & TBL_EMAIL_STATUS & " (EmailID)"

    ' Defaults via DAO setzen (direkt in der Backend-DB)
    SetzeBEDefaults dbNew

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

' Setzt Defaults in einer frisch erstellten Backend-DB via DAO
Private Sub SetzeBEDefaults(dbBE As DAO.Database)
    On Error Resume Next

    SetzeFeldDefaultInDB dbBE, TBL_SYNC_LAUF, "Status", """Gestartet"""
    SetzeFeldDefaultInDB dbBE, TBL_SYNC_LAUF, "AnzahlGelesen", "0"
    SetzeFeldDefaultInDB dbBE, TBL_SYNC_LAUF, "AnzahlNeu", "0"
    SetzeFeldDefaultInDB dbBE, TBL_SYNC_LAUF, "AnzahlDuplikate", "0"
    SetzeFeldDefaultInDB dbBE, TBL_SYNC_LAUF, "AnzahlFehler", "0"

    SetzeFeldDefaultInDB dbBE, TBL_ORDNER, "ParentID", "0"
    SetzeFeldDefaultInDB dbBE, TBL_ORDNER, "ElementAnzahl", "0"

    SetzeFeldDefaultInDB dbBE, TBL_THREADS, "Antwortanzahl", "1"

    SetzeFeldDefaultInDB dbBE, TBL_EMAILS, "ThreadID", "0"
    SetzeFeldDefaultInDB dbBE, TBL_EMAILS, "OrdnerID", "0"
    SetzeFeldDefaultInDB dbBE, TBL_EMAILS, "KontaktID_Absender", "0"
    SetzeFeldDefaultInDB dbBE, TBL_EMAILS, "SyncLaufID", "0"
    SetzeFeldDefaultInDB dbBE, TBL_EMAILS, "Groesse", "0"
    SetzeFeldDefaultInDB dbBE, TBL_EMAILS, "Wichtigkeit", "1"
    SetzeFeldDefaultInDB dbBE, TBL_EMAILS, "AnhangAnzahl", "0"
    SetzeFeldDefaultInDB dbBE, TBL_EMAILS, "Status", """Neu"""

    SetzeFeldDefaultInDB dbBE, TBL_CONTENT, "GroesseHTML", "0"
    SetzeFeldDefaultInDB dbBE, TBL_CONTENT, "GroesseText", "0"

    SetzeFeldDefaultInDB dbBE, TBL_EMPFAENGER, "KontaktID", "0"

    SetzeFeldDefaultInDB dbBE, TBL_ANHAENGE, "Groesse", "0"
    SetzeFeldDefaultInDB dbBE, TBL_ANHAENGE, "AnhangTyp", "1"

    On Error GoTo 0
End Sub

' Setzt DefaultValue fuer ein Feld direkt in einer uebergebenen Datenbank
Private Sub SetzeFeldDefaultInDB(db As DAO.Database, ByVal strTabelle As String, ByVal strFeld As String, ByVal strDefault As String)
    On Error Resume Next
    db.TableDefs(strTabelle).Fields(strFeld).DefaultValue = strDefault
    Err.Clear
    On Error GoTo 0
End Sub


' ===========================================================================
' PRIVATE HELFER
' ===========================================================================

Private Function PruefeDatei(ByVal strPfad As String) As Boolean
    On Error Resume Next
    PruefeDatei = (Dir(strPfad) <> "")
    If Err.Number <> 0 Then PruefeDatei = False: Err.Clear
    On Error GoTo 0
End Function
