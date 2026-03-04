Attribute VB_Name = "modAsyncBuffer"
Option Compare Database
Option Explicit

' ===========================================================================
' modAsyncBuffer - Schreib-Puffer mit Queue-Persistenz + Netzwerk-Resilienz
' ===========================================================================
' v0.4: Kompletter Umbau basierend auf Performance-Test-Ergebnissen:
'
' ARCHITEKTUR-ENTSCHEIDUNGEN (v0.3.3 Performance-Tests, VPN):
'   - Direct DAO statt Linked Tables (9.3x schneller)
'   - EINE Transaktion pro Flush (kein Cycling - Error 3246!)
'   - Recordset.AddNew statt Execute INSERT (100x schneller)
'   - IN()-Batch fuer Hash-Duplikatpruefung
'   - Datei-Queue mit Netzwerk-Resilienz (modFileManager)
'
' FEATURES v0.4:
'   - Pause/Resume: Queue ueberlebt Pause, Duplikate vermieden
'   - Netzwerk-Recovery: WarteAufNetzwerk() bei VPN-Verlust
'   - Sinnvolle Ablagestruktur: Jahr/Monat/Timestamp_Absender_Betreff
'   - Anhang-Unterordner pro Mail
'   - Direct DAO: OpenDatabase() direkt auf Backend-DB
'
' MUSTER:
'   BufferInit
'   BufferSetzeKontext lngSyncLaufID, lngOrdnerID, ...
'   For Each Mail:
'     ExtrahiereKomplett objMail, mk
'     BufferHinzufuegen mk
'     If BufferIstVoll Then BufferFlush
'   Next
'   BufferFlush            ' Restliche Daten
'   DateiQueueVerarbeiten  ' Dateien: Temp -> Netzwerk
'   BufferLeeren
'
' Abhaengigkeiten: modMailExtract (Types), modDAO, modFileManager,
'                  modKontakte, modCrypto, modStringUtils, modLogging,
'                  modGlobals, modBackend, modSchema
' ===========================================================================


' ---------------------------------------------------------------------------
' MODUL-VARIABLEN: Schreib-Puffer
' ---------------------------------------------------------------------------
Private m_aBuffer()         As TypMailKomplett
Private m_lngBufferAnzahl   As Long
Private m_lngBufferMax      As Long
Private m_blnInitialisiert  As Boolean

' Kontext (wird einmal pro SyncFolder gesetzt)
Private m_lngKontextSyncLaufID  As Long
Private m_lngKontextOrdnerID    As Long
Private m_strKontextProjekt     As String
Private m_strKontextPhase       As String
Private m_strKontextExportBase  As String
Private m_blnKontextAnhaenge    As Boolean
Private m_blnKontextMSG         As Boolean
Private m_blnKontextFilter      As Boolean

' ---------------------------------------------------------------------------
' MODUL-VARIABLEN: Datei-Queue
' ---------------------------------------------------------------------------
Private m_aDateiQueue()         As TypDateiOperation
Private m_lngDateiQueueAnzahl   As Long

' ---------------------------------------------------------------------------
' MODUL-VARIABLEN: Pause/Resume
' ---------------------------------------------------------------------------
Private m_blnPausiert           As Boolean
Private m_lngPauseWarteMs       As Long     ' Pause-Polling (500ms Default)

' ---------------------------------------------------------------------------
' STATISTIK (kumuliert ueber alle Flushes eines Sync-Laufs)
' ---------------------------------------------------------------------------
Private m_lngGesamtNeu          As Long
Private m_lngGesamtDuplikate    As Long
Private m_lngGesamtFehler       As Long
Private m_lngGesamtVerarbeitet  As Long
Private m_lngGesamtDateien      As Long


' ===========================================================================
' INITIALISIERUNG
' ===========================================================================

' Puffer initialisieren (am Anfang jedes SyncFolder-Aufrufs)
' Standard-Groesse aus Config oder 50 wenn nicht konfiguriert
Public Sub BufferInit(Optional ByVal lngMaxGroesse As Long = 0)
    If lngMaxGroesse <= 0 Then
        lngMaxGroesse = CLng(LeseConfig("BufferGroesse", "50"))
    End If
    If lngMaxGroesse < 5 Then lngMaxGroesse = 5
    If lngMaxGroesse > 500 Then lngMaxGroesse = 500

    m_lngBufferMax = lngMaxGroesse
    m_lngBufferAnzahl = 0
    ReDim m_aBuffer(0 To m_lngBufferMax - 1)

    m_lngDateiQueueAnzahl = 0
    ReDim m_aDateiQueue(0 To 99)

    m_lngGesamtNeu = 0
    m_lngGesamtDuplikate = 0
    m_lngGesamtFehler = 0
    m_lngGesamtVerarbeitet = 0
    m_lngGesamtDateien = 0

    m_blnPausiert = False
    m_lngPauseWarteMs = 500

    m_blnInitialisiert = True
    LogDebug "Buffer initialisiert: Max=" & m_lngBufferMax & " (Direct DAO, v0.4)", "BUFFER"
End Sub


' Sync-Kontext setzen (Infos die fuer alle Mails im Batch gelten)
Public Sub BufferSetzeKontext(ByVal lngSyncLaufID As Long, _
                               ByVal lngOrdnerID As Long, _
                               ByVal strProjekt As String, _
                               ByVal strPhase As String, _
                               ByVal strExportBase As String, _
                               ByVal blnAnhaenge As Boolean, _
                               ByVal blnMSG As Boolean, _
                               ByVal blnFilter As Boolean)
    m_lngKontextSyncLaufID = lngSyncLaufID
    m_lngKontextOrdnerID = lngOrdnerID
    m_strKontextProjekt = strProjekt
    m_strKontextPhase = strPhase
    m_strKontextExportBase = strExportBase
    m_blnKontextAnhaenge = blnAnhaenge
    m_blnKontextMSG = blnMSG
    m_blnKontextFilter = blnFilter
End Sub


' ===========================================================================
' PAUSE / RESUME
' ===========================================================================

' Pausiert die Queue-Verarbeitung. Laufende DB-Transactions werden
' zu Ende gefuehrt, aber kein neuer Flush/DateiQueue-Schritt startet.
Public Sub BufferPause()
    m_blnPausiert = True
    LogInfo "Buffer PAUSIERT - Verarbeitung angehalten", "BUFFER"
End Sub

' Setzt die Queue-Verarbeitung fort.
Public Sub BufferResume()
    m_blnPausiert = False
    LogInfo "Buffer FORTGESETZT", "BUFFER"
End Sub

' Ist der Buffer gerade pausiert?
Public Function BufferIstPausiert() As Boolean
    BufferIstPausiert = m_blnPausiert
End Function

' Wartet solange Buffer pausiert ist. Gibt True zurueck wenn fortgesetzt,
' False wenn abgebrochen (g_blnAbbrechen).
Private Function WarteWennPausiert() As Boolean
    If Not m_blnPausiert Then
        WarteWennPausiert = True
        Exit Function
    End If

    Do While m_blnPausiert
        If g_blnAbbrechen Then
            WarteWennPausiert = False
            Exit Function
        End If
        Sleep m_lngPauseWarteMs
        DoEvents
    Loop

    WarteWennPausiert = True
End Function


' ===========================================================================
' PUFFER-OPERATIONEN
' ===========================================================================

' Mail-Daten zum Puffer hinzufuegen
' Gibt True zurueck wenn erfolgreich
Public Function BufferHinzufuegen(ByRef mk As TypMailKomplett) As Boolean
    If Not m_blnInitialisiert Then BufferInit

    ' Puffer-Array vergroessern wenn noetig
    If m_lngBufferAnzahl > UBound(m_aBuffer) Then
        ReDim Preserve m_aBuffer(0 To m_lngBufferAnzahl + m_lngBufferMax - 1)
    End If

    m_aBuffer(m_lngBufferAnzahl) = mk
    m_lngBufferAnzahl = m_lngBufferAnzahl + 1
    m_lngGesamtVerarbeitet = m_lngGesamtVerarbeitet + 1

    BufferHinzufuegen = True
End Function


' Prueft ob der Puffer voll ist (Schwellwert erreicht)
Public Function BufferIstVoll() As Boolean
    BufferIstVoll = (m_lngBufferAnzahl >= m_lngBufferMax)
End Function


' Aktuelle Puffergroesse (Anzahl Eintraege)
Public Function BufferGroesse() As Long
    BufferGroesse = m_lngBufferAnzahl
End Function


' ===========================================================================
' STATISTIK-ABFRAGEN
' ===========================================================================

Public Function BufferGesamtNeu() As Long
    BufferGesamtNeu = m_lngGesamtNeu
End Function

Public Function BufferGesamtDuplikate() As Long
    BufferGesamtDuplikate = m_lngGesamtDuplikate
End Function

Public Function BufferGesamtFehler() As Long
    BufferGesamtFehler = m_lngGesamtFehler
End Function

Public Function BufferGesamtVerarbeitet() As Long
    BufferGesamtVerarbeitet = m_lngGesamtVerarbeitet
End Function

Public Function BufferGesamtDateien() As Long
    BufferGesamtDateien = m_lngGesamtDateien
End Function

Public Function BufferStatistik() As String
    BufferStatistik = "Verarbeitet=" & m_lngGesamtVerarbeitet & _
                      " Neu=" & m_lngGesamtNeu & _
                      " Dup=" & m_lngGesamtDuplikate & _
                      " Err=" & m_lngGesamtFehler & _
                      " Dateien=" & m_lngGesamtDateien
End Function


' ===========================================================================
' PUFFER FLUSH - DIRECT DAO (Kernroutine v0.4)
' ===========================================================================

' Schreibt alle gepufferten Mails per Direct DAO in die Backend-DB.
' EINE Transaktion pro Flush (kein Cycling!).
' Gibt Anzahl neu gespeicherter Mails zurueck.
Public Function BufferFlush() As Long
    On Error GoTo ErrHandler

    If m_lngBufferAnzahl = 0 Then BufferFlush = 0: Exit Function

    ' Pause-Pruefung
    If Not WarteWennPausiert() Then BufferFlush = 0: Exit Function

    Dim lngNeu As Long, lngDup As Long, lngErr As Long
    lngNeu = 0: lngDup = 0: lngErr = 0

    LogDebug "BufferFlush: " & m_lngBufferAnzahl & " Eintraege (Direct DAO)...", "BUFFER"

    ' --- 0. Backend-DB Pfad ermitteln ---
    Dim strBackendPfad As String
    strBackendPfad = GetBackendPfad()

    ' Falls kein Backend konfiguriert: Fallback auf CurrentDb (lokal)
    Dim blnDirect As Boolean
    blnDirect = (strBackendPfad <> "" And Dir(strBackendPfad) <> "")

    ' --- 1. Batch-Duplikatpruefung (Direct DAO) ---
    Dim arrHashes() As String
    Dim arrExistiert() As Boolean
    Dim n As Long

    ReDim arrHashes(0 To m_lngBufferAnzahl - 1)
    ReDim arrExistiert(0 To m_lngBufferAnzahl - 1)

    For n = 0 To m_lngBufferAnzahl - 1
        arrHashes(n) = m_aBuffer(n).UniqueHash
    Next n

    Call PruefeHashesDirect(arrHashes, arrExistiert, strBackendPfad)

    ' --- 2. Transaktion: Nicht-Duplikate in DB schreiben ---
    ' Hinweis: DB-Schreiben laueft ueber modDAO (CurrentDb/Linked Tables)
    ' fuer Kompatibilitaet mit bestehender Infrastruktur.
    ' Direct DAO wird fuer Hash-Lookup genutzt (9.3x schneller).
    Dim ws As DAO.Workspace
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans

    On Error GoTo RollbackHandler

    For n = 0 To m_lngBufferAnzahl - 1
        ' Duplikat ueberspringen
        If arrExistiert(n) Then
            lngDup = lngDup + 1
            LogTrace "Duplikat: " & Left(m_aBuffer(n).Mail.Betreff, 40), "BUFFER"
            Call LoescheTempDateien(m_aBuffer(n))
            GoTo NaechsterEintrag
        End If

        ' Ungueltige Mail ueberspringen
        If Not m_aBuffer(n).Mail.IstGueltig Then
            lngErr = lngErr + 1
            Call LoescheTempDateien(m_aBuffer(n))
            GoTo NaechsterEintrag
        End If

        ' Mail verarbeiten und in DB speichern
        Dim lngEmailID As Long
        lngEmailID = SpeichereMailAusBuffer(m_aBuffer(n))

        If lngEmailID > 0 Then
            lngNeu = lngNeu + 1
        Else
            lngErr = lngErr + 1
            Call LoescheTempDateien(m_aBuffer(n))
        End If

NaechsterEintrag:
    Next n

    ws.CommitTrans
    On Error GoTo ErrHandler

    ' Statistik aktualisieren
    m_lngGesamtNeu = m_lngGesamtNeu + lngNeu
    m_lngGesamtDuplikate = m_lngGesamtDuplikate + lngDup
    m_lngGesamtFehler = m_lngGesamtFehler + lngErr

    ' Puffer leeren (Daten sind in DB)
    m_lngBufferAnzahl = 0
    If m_lngBufferMax > 0 Then
        ReDim m_aBuffer(0 To m_lngBufferMax - 1)
    End If

    LogDebug "BufferFlush: Neu=" & lngNeu & " Dup=" & lngDup & " Err=" & lngErr, "BUFFER"
    BufferFlush = lngNeu
    Exit Function

RollbackHandler:
    Dim strErr As String
    strErr = Err.Description
    On Error Resume Next
    ws.Rollback
    On Error GoTo 0

    LogError "BufferFlush ROLLBACK: " & strErr, "BUFFER"
    m_lngGesamtFehler = m_lngGesamtFehler + m_lngBufferAnzahl
    m_lngBufferAnzahl = 0
    If m_lngBufferMax > 0 Then
        ReDim m_aBuffer(0 To m_lngBufferMax - 1)
    End If
    BufferFlush = 0
    Exit Function

ErrHandler:
    LogVBAError "BufferFlush"
    BufferFlush = 0
End Function


' ===========================================================================
' EINZELNE MAIL AUS PUFFER IN DB SCHREIBEN + DATEI-QUEUE AUFBAUEN
' ===========================================================================

Private Function SpeichereMailAusBuffer(ByRef mk As TypMailKomplett) As Long
    On Error GoTo ErrHandler

    Dim lngEmailID      As Long
    Dim lngKontaktID    As Long
    Dim lngThreadID     As Long
    Dim strEmailTyp     As String

    ' --- Absender-Kontakt ermitteln/erstellen ---
    If mk.Mail.AbsenderEmailTyp = "EX" Then
        strEmailTyp = "EX"
    Else
        strEmailTyp = "SMTP"
    End If
    lngKontaktID = GetOderErstelleKontakt(mk.Mail.AbsenderName, _
                                           mk.Mail.AbsenderEmail, strEmailTyp)

    ' --- Thread ermitteln/erstellen ---
    lngThreadID = GetOderErstelleThread(mk.Mail.Betreff, mk.Mail.InternetMessageID, _
                                         mk.Mail.InReplyTo, mk.Mail.EmpfangenAm, _
                                         mk.Mail.AbsenderName)

    ' --- Email-Datensatz speichern ---
    lngEmailID = SpeichereEmail( _
        mk.Mail.EntryID, mk.UniqueHash, lngThreadID, m_lngKontextOrdnerID, _
        lngKontaktID, m_lngKontextSyncLaufID, mk.Mail.Betreff, _
        mk.Mail.AbsenderName, mk.Mail.AbsenderEmail, mk.Mail.EmpfangenAm, _
        mk.Mail.GesendetAm, mk.Mail.Groesse, mk.Mail.Wichtigkeit, _
        mk.Mail.Gelesen, mk.Mail.HatAnhaenge, mk.Mail.AnhangAnzahl, _
        mk.Mail.MessageClass, mk.Mail.InternetMessageID)

    If lngEmailID = 0 Then
        SpeichereMailAusBuffer = 0
        Exit Function
    End If

    ' --- Content speichern ---
    Call SpeichereEmailContent(lngEmailID, mk.Mail.HTMLBody, mk.Mail.PlainTextBody)

    ' --- Empfaenger speichern ---
    Dim e As Integer
    Dim lngEmpfKontaktID As Long
    For e = 0 To mk.EmpfaengerAnzahl - 1
        If IstGueltigeEmail(mk.Empfaenger(e).Email) Then
            lngEmpfKontaktID = GetOderErstelleKontakt(mk.Empfaenger(e).Name, _
                                                       mk.Empfaenger(e).Email)
            If lngEmpfKontaktID > 0 Then
                Call AktualisiereKontaktEmail(lngEmpfKontaktID, mk.Empfaenger(e).Email)
            End If
        Else
            lngEmpfKontaktID = 0
        End If

        Call SpeichereEmpfaenger(lngEmailID, lngEmpfKontaktID, _
                                  mk.Empfaenger(e).Typ, mk.Empfaenger(e).Name, _
                                  mk.Empfaenger(e).Email)
    Next e

    ' --- Ablagestruktur: Mail-Ordner bestimmen (v0.4) ---
    Dim strMailOrdner As String
    strMailOrdner = BaueMailOrdnerPfad(m_strKontextExportBase, _
                                        m_strKontextProjekt, _
                                        m_strKontextPhase, _
                                        mk.Mail.EmpfangenAm, _
                                        mk.Mail.AbsenderName, _
                                        mk.Mail.Betreff)

    ' --- Anhang-Metadaten speichern + Datei-Queue ---
    Dim a As Integer
    Dim lngAnhangID As Long
    For a = 0 To mk.AnhangAnzahl - 1
        lngAnhangID = SpeichereAnhangMetadaten(lngEmailID, mk.Anhaenge(a).Dateiname, _
                                                 mk.Anhaenge(a).Groesse, mk.Anhaenge(a).MimeType, _
                                                 mk.Anhaenge(a).AnhangTyp, mk.Anhaenge(a).IstVersteckt)

        ' Datei-Queue: Temp -> Netzwerk (v0.4 Ablagestruktur)
        If mk.Anhaenge(a).TempPfad <> "" And lngAnhangID > 0 Then
            Dim strAnhangZiel As String
            strAnhangZiel = BaueAnhangPfad(strMailOrdner, mk.Anhaenge(a).Dateiname)

            Call DateiQueueHinzufuegen(mk.Anhaenge(a).TempPfad, strAnhangZiel, _
                                       "Anhang", lngEmailID, lngAnhangID)
        End If
    Next a

    ' --- MSG Datei-Queue (v0.4 Ablagestruktur) ---
    If mk.MSGTempPfad <> "" Then
        Dim strMSGZiel As String
        strMSGZiel = BaueMSGPfad(strMailOrdner)

        Call DateiQueueHinzufuegen(mk.MSGTempPfad, strMSGZiel, "MSG", lngEmailID, 0)
    End If

    ' --- Status ---
    Call SpeichereEmailStatus(lngEmailID, "Verarbeitet", "Sync", "Import via BufferFlush v0.4")

    SpeichereMailAusBuffer = lngEmailID
    Exit Function

ErrHandler:
    LogVBAError "SpeichereMailAusBuffer [" & Left(Nz(mk.Mail.Betreff, "?"), 30) & "]"
    SpeichereMailAusBuffer = 0
End Function


' ===========================================================================
' BATCH-DUPLIKATPRUEFUNG (Direct DAO, v0.4)
' ===========================================================================

' Prueft Hashes via Direct DAO auf der Backend-DB.
' Bei Fehler: Fallback auf CurrentDb (Linked Tables).
Private Sub PruefeHashesDirect(ByRef arrHashes() As String, _
                                ByRef arrExistiert() As Boolean, _
                                ByVal strBackendPfad As String)
    On Error GoTo ErrHandler

    Dim n As Long
    Dim lngAnzahl As Long
    lngAnzahl = UBound(arrHashes) + 1

    ' Kein Backend oder leer: Fallback auf Linked
    If strBackendPfad = "" Or Dir(strBackendPfad) = "" Then
        Call PruefeHashesLinked(arrHashes, arrExistiert)
        Exit Sub
    End If

    ' Fuer kleine Batches: einzelne Pruefung via Direct DAO
    If lngAnzahl <= 5 Then
        Dim dbDir As DAO.Database
        Set dbDir = DBEngine.Workspaces(0).OpenDatabase(strBackendPfad)
        Dim rsDir As DAO.Recordset
        For n = 0 To lngAnzahl - 1
            Set rsDir = dbDir.OpenRecordset( _
                "SELECT Count(*) FROM tblEmails WHERE UniqueHash='" & _
                SQLSafe(arrHashes(n)) & "'", dbOpenSnapshot)
            arrExistiert(n) = (Nz(rsDir(0), 0) > 0)
            rsDir.Close: Set rsDir = Nothing
        Next n
        dbDir.Close: Set dbDir = Nothing
        Exit Sub
    End If

    ' Fuer groessere Batches: IN-Query via Direct DAO
    Dim strIN As String
    strIN = ""
    For n = 0 To lngAnzahl - 1
        If strIN <> "" Then strIN = strIN & ","
        strIN = strIN & "'" & SQLSafe(arrHashes(n)) & "'"
    Next n

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim dbDirect As DAO.Database
    Set dbDirect = DBEngine.Workspaces(0).OpenDatabase(strBackendPfad)
    Dim rs As DAO.Recordset
    Set rs = dbDirect.OpenRecordset( _
        "SELECT UniqueHash FROM tblEmails WHERE UniqueHash IN (" & strIN & ")", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        dict(CStr(rs!UniqueHash)) = True
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    dbDirect.Close: Set dbDirect = Nothing

    For n = 0 To lngAnzahl - 1
        arrExistiert(n) = dict.Exists(arrHashes(n))
    Next n

    Set dict = Nothing
    Exit Sub

ErrHandler:
    ' Fallback auf Linked Tables (CurrentDb)
    LogWarn "Direct-DAO Hash-Pruefung fehlgeschlagen, Fallback auf Linked: " & _
            Err.Description, "BUFFER"
    On Error Resume Next
    Call PruefeHashesLinked(arrHashes, arrExistiert)
    On Error GoTo 0
End Sub


' Fallback: Hash-Pruefung ueber Linked Tables (CurrentDb)
Private Sub PruefeHashesLinked(ByRef arrHashes() As String, _
                                ByRef arrExistiert() As Boolean)
    On Error GoTo ErrHandler

    Dim n As Long
    For n = 0 To UBound(arrHashes)
        arrExistiert(n) = ExistiertMailHash(arrHashes(n))
    Next n
    Exit Sub

ErrHandler:
    LogWarn "Linked Hash-Pruefung fehlgeschlagen: " & Err.Description, "BUFFER"
    On Error Resume Next
    For n = 0 To UBound(arrHashes)
        arrExistiert(n) = False
    Next n
    On Error GoTo 0
End Sub


' ===========================================================================
' DATEI-QUEUE (Temp -> Netzwerk, v0.4 mit Resilienz)
' ===========================================================================

' Datei-Operation zur Queue hinzufuegen
Public Sub DateiQueueHinzufuegen(ByVal strQuelle As String, _
                                  ByVal strZiel As String, _
                                  ByVal strTyp As String, _
                                  ByVal lngEmailID As Long, _
                                  ByVal lngAnhangID As Long)
    ' Queue vergroessern wenn noetig
    If m_lngDateiQueueAnzahl > UBound(m_aDateiQueue) Then
        ReDim Preserve m_aDateiQueue(0 To m_lngDateiQueueAnzahl + 99)
    End If

    With m_aDateiQueue(m_lngDateiQueueAnzahl)
        .QuellPfad = strQuelle
        .ZielPfad = strZiel
        .OperationsTyp = strTyp
        .EmailID = lngEmailID
        .AnhangID = lngAnhangID
        .Versuche = 0
    End With

    m_lngDateiQueueAnzahl = m_lngDateiQueueAnzahl + 1
End Sub


' Datei-Queue abarbeiten mit Netzwerk-Resilienz (v0.4)
' Bei Netzwerk-Verlust: wartet bis VPN/LAN wieder da.
' Bei Pause: wartet bis Resume.
' Gibt Anzahl erfolgreich kopierter Dateien zurueck.
Public Function DateiQueueVerarbeiten() As Long
    On Error GoTo ErrHandler

    If m_lngDateiQueueAnzahl = 0 Then DateiQueueVerarbeiten = 0: Exit Function

    Dim lngOK As Long, lngFail As Long
    Dim intMaxRetries As Integer
    lngOK = 0: lngFail = 0
    intMaxRetries = CInt(LeseConfig("NetzwerkRetries", "3"))

    Dim lngBasisPause As Long
    lngBasisPause = CLng(LeseConfig("NetzwerkRetryPause", "2000"))

    LogDebug "DateiQueue: " & m_lngDateiQueueAnzahl & " Dateien verarbeiten...", "BUFFER"

    ' Netzwerk-Pruefung vor dem Start
    If m_strKontextExportBase <> "" Then
        If Not IstNetzwerkOK(m_strKontextExportBase) Then
            LogWarn "Netzwerk nicht erreichbar vor DateiQueue - warte...", "BUFFER"
            If Not WarteAufNetzwerk(m_strKontextExportBase) Then
                LogError "Netzwerk-Timeout - DateiQueue NICHT verarbeitet. " & _
                         m_lngDateiQueueAnzahl & " Dateien verbleiben im Temp.", "BUFFER"
                DateiQueueVerarbeiten = 0
                Exit Function
            End If
        End If
    End If

    Dim n As Long
    For n = 0 To m_lngDateiQueueAnzahl - 1
        ' Abbruch-Pruefung
        If g_blnAbbrechen Then Exit For

        ' Pause-Pruefung
        If Not WarteWennPausiert() Then Exit For

        If KopiereNachNetzwerk(m_aDateiQueue(n).QuellPfad, _
                                m_aDateiQueue(n).ZielPfad, _
                                intMaxRetries, lngBasisPause) Then
            ' Erfolg: DB-Pfad aktualisieren
            Select Case m_aDateiQueue(n).OperationsTyp
                Case "MSG"
                    Call SetzeEmailMSGPfad(m_aDateiQueue(n).EmailID, _
                                           m_aDateiQueue(n).ZielPfad)
                Case "Anhang"
                    Call AktualisiereAnhangPfad(m_aDateiQueue(n).AnhangID, _
                                                m_aDateiQueue(n).ZielPfad)
            End Select

            lngOK = lngOK + 1

            ' Temp-Datei loeschen (erfolgreich kopiert)
            On Error Resume Next
            Kill m_aDateiQueue(n).QuellPfad
            On Error GoTo ErrHandler
        Else
            lngFail = lngFail + 1
            LogWarn "Datei-Kopie endgueltig fehlgeschlagen: " & _
                    m_aDateiQueue(n).QuellPfad & " -> " & _
                    m_aDateiQueue(n).ZielPfad, "BUFFER"
        End If

        ' UI responsiv halten (alle 10 Dateien)
        If n Mod 10 = 0 Then DoEvents
    Next n

    m_lngGesamtDateien = m_lngGesamtDateien + lngOK

    LogInfo "DateiQueue: " & lngOK & " OK, " & lngFail & " fehlgeschlagen " & _
            "von " & m_lngDateiQueueAnzahl & " gesamt", "BUFFER"

    ' Queue leeren
    m_lngDateiQueueAnzahl = 0
    ReDim m_aDateiQueue(0 To 99)

    DateiQueueVerarbeiten = lngOK
    Exit Function

ErrHandler:
    LogVBAError "DateiQueueVerarbeiten"
    DateiQueueVerarbeiten = 0
End Function


' Datei-Queue Groesse abfragen
Public Function DateiQueueGroesse() As Long
    DateiQueueGroesse = m_lngDateiQueueAnzahl
End Function


' ===========================================================================
' AUFRAEUMEN
' ===========================================================================

' Puffer und Queue komplett leeren + Statistik zuruecksetzen
Public Sub BufferLeeren()
    m_lngBufferAnzahl = 0
    m_lngDateiQueueAnzahl = 0

    If m_lngBufferMax > 0 Then
        ReDim m_aBuffer(0 To m_lngBufferMax - 1)
    End If
    ReDim m_aDateiQueue(0 To 99)

    m_blnPausiert = False
    m_blnInitialisiert = False
    LogDebug "Buffer geleert", "BUFFER"
End Sub


' Temp-Dateien eines einzelnen TypMailKomplett loeschen
' (z.B. wenn Duplikat oder Fehler)
Private Sub LoescheTempDateien(ByRef mk As TypMailKomplett)
    On Error Resume Next

    ' MSG Temp
    If mk.MSGTempPfad <> "" Then
        Kill mk.MSGTempPfad
        mk.MSGTempPfad = ""
    End If

    ' Anhang Temp
    Dim a As Integer
    For a = 0 To mk.AnhangAnzahl - 1
        If mk.Anhaenge(a).TempPfad <> "" Then
            Kill mk.Anhaenge(a).TempPfad
            mk.Anhaenge(a).TempPfad = ""
        End If
    Next a

    On Error GoTo 0
End Sub

' Win32 API fuer Sleep (falls nicht in anderem Modul deklariert)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
