Option Compare Database
Option Explicit

' ===========================================================================
' modSync - Synchronisations-Orchestrierung
' ===========================================================================
' Hauptmodul fuer die Outlook-zu-Access-Synchronisation.
' v0.4.3: COM-Resilienz + Throttling + Notbremse
'
' Architektur:
'   1. EXTRACT: Alle Daten aus RDO-Objekt lesen (modMailExtract)
'   2. RELEASE: COM-Objekt sofort freigeben (Outlook entlasten)
'   3. BUFFER:  In Speicher-Puffer sammeln (modAsyncBuffer)
'   4. FLUSH:   Batch in DB schreiben (Direct DAO, 1 Transaktion)
'   5. COPY:    Dateien Temp->Netzwerk (modFileManager, Retry+Recovery)
'
' COM-Resilienz (v0.4.3):
'   - Fehlerklassifikation pro Mail (TRANSIENT/FATAL/ITEM)
'   - Throttling: 10ms Pause alle 10 Mails (Outlook-Messagepump)
'   - Notbremse: 10 aufeinanderfolgende Fehler -> Abbruch
'   - FATAL-Erkennung: Ordner-Referenz verloren -> sauberes Ende
'   - Statistik: COM-Retries + Reconnects im Ergebnis
'
' Neu in v0.4:
'   - Direct DAO fuer Hash-Lookup (9.3x schneller)
'   - Netzwerk-Recovery (VPN-Verlust -> WarteAufNetzwerk)
'   - Pause/Resume (BufferPause/BufferResume)
'   - Ablagestruktur: Jahr/Monat/Timestamp_Absender_Betreff
'   - Anhang-Unterordner pro Mail
'
' Oeffentliche Routinen:
'   SyncPosteingang [Projekt, Phase, MaxMails, Subfolder]
'   SyncOrdner "Pfad" [, Projekt, Phase, Max, Subfolder]
'   SyncPostfach "Postfachname" [, Projekt, Phase, Max, Tiefe]
'   SyncOrdnerStruktur [Tiefe]
'   SyncMitProfil "ProfilName"
'   ErstelleSyncProfil "Name", "Projekt", "Phase"
'   ProfilOrdnerHinzufuegen ProfilID, "OrdnerPfad"
'
' Abhaengigkeiten: modMailExtract, modAsyncBuffer, modFileManager,
'                  modBackend, modOutlookConnect, modDAO, modKontakte,
'                  modStringUtils, modCrypto, modLogging, modGlobals
' ===========================================================================


' ---------------------------------------------------------------------------
' POSTEINGANG SYNCHRONISIEREN (Convenience-Wrapper)
' ---------------------------------------------------------------------------
Public Sub SyncPosteingang(Optional ByVal strProjekt As String = "Standard", _
                            Optional ByVal strPhase As String = "Standard", _
                            Optional ByVal lngMaxMails As Long = 0, _
                            Optional ByVal blnSubfolder As Boolean = False)

    If Not ConnectRDO() Then
        LogError "SyncPosteingang: Keine RDO-Verbindung", "SYNC"
        Exit Sub
    End If

    Dim objInbox As Object
    Set objInbox = g_objRDO.GetDefaultFolder(olFolderInbox)

    If objInbox Is Nothing Then
        LogError "SyncPosteingang: Posteingang nicht erreichbar", "SYNC"
        Exit Sub
    End If

    ' Maximale Mailanzahl aus Config lesen wenn nicht angegeben
    If lngMaxMails <= 0 Then
        lngMaxMails = CLng(CacheGetConfig(CFG_MAX_MAILS, "500"))
    End If

    Call SyncFolder(objInbox, strProjekt, strPhase, lngMaxMails, blnSubfolder)

    Set objInbox = Nothing
End Sub


' ---------------------------------------------------------------------------
' ORDNER PER PFAD SYNCHRONISIEREN
' ---------------------------------------------------------------------------
Public Sub SyncOrdner(ByVal strOrdnerPfad As String, _
                       Optional ByVal strProjekt As String = "Standard", _
                       Optional ByVal strPhase As String = "Standard", _
                       Optional ByVal lngMaxMails As Long = 0, _
                       Optional ByVal blnSubfolder As Boolean = False)

    Dim objFolder As Object
    Set objFolder = OeffneOrdner(strOrdnerPfad)

    If objFolder Is Nothing Then
        LogError "SyncOrdner: Ordner '" & strOrdnerPfad & "' nicht gefunden", "SYNC"
        Exit Sub
    End If

    If lngMaxMails <= 0 Then
        lngMaxMails = CLng(CacheGetConfig(CFG_MAX_MAILS, "500"))
    End If

    Call SyncFolder(objFolder, strProjekt, strPhase, lngMaxMails, blnSubfolder)

    Set objFolder = Nothing
End Sub


' ===========================================================================
' KERN-ROUTINE: Ordner-Objekt synchronisieren
' v0.4: Extract-Release-Process + Direct DAO + Netzwerk-Resilienz
' ===========================================================================
'
' Ablauf:
'   1. Netzwerk-Pruefung (WarteAufNetzwerk bei Bedarf)
'   2. Fuer jede Mail im Ordner:
'      a. ExtrahiereKomplett() -> alle Daten + Anhaenge in TypMailKomplett
'      b. COM-Objekt freigeben (Set objItem = Nothing)
'      c. BufferHinzufuegen() -> in Speicher-Puffer
'      d. Wenn Puffer voll: BufferFlush() -> Direct DAO Batch-INSERT
'   3. BufferFlush() -> Restliche Daten schreiben
'   4. DateiQueueVerarbeiten() -> Dateien Temp->Netzwerk (mit Recovery)
'   5. BereinigeTempDateien() -> Temp-Ordner aufraeumen

Public Sub SyncFolder(objFolder As Object, _
                       ByVal strProjekt As String, _
                       ByVal strPhase As String, _
                       ByVal lngMaxMails As Long, _
                       Optional ByVal blnSubfolder As Boolean = False)

    On Error GoTo ErrHandler

    Dim lngSyncLaufID   As Long
    Dim lngOrdnerID     As Long
    Dim lngGesamt       As Long
    Dim lngMailCount    As Long
    Dim lngFehler       As Long
    Dim objItem         As Object
    Dim i               As Long
    Dim mk              As TypMailKomplett
    Dim dblHeartbeatLast As Double
    Const HEARTBEAT_S As Double = 5#

    ' Performance-Profiling (zur Bottleneck-Analyse)
    Dim dblT0              As Double
    Dim dblFetchItemsS     As Double
    Dim dblMsgClassS       As Double
    Dim dblExtractS        As Double
    Dim dblBufferAddS      As Double
    Dim dblFlushS          As Double
    Dim dblFinalFlushS     As Double
    Dim dblDateiQueueS     As Double
    Dim dblBeendeSyncS     As Double
    Dim dblCleanupS        As Double
    Dim lngFlushCount      As Long

    ' --- Initialisierung ---
    InitGlobals
    g_blnAbbrechen = False
    g_dtSyncStart = Timer
    TimerStart "SyncFolder"

    If Nz(strProjekt, "") = "" Then strProjekt = "Standard"
    If Nz(strPhase, "") = "" Then strPhase = "Standard"

    Debug.Print "[SYNC] Zaehle Elemente in Ordner..."
    lngGesamt = objFolder.Items.Count
    Debug.Print "[SYNC] Elemente gezaehlt: " & lngGesamt
    lngMailCount = 0: lngFehler = 0
    dblHeartbeatLast = Timer

    Debug.Print String(70, "=")
    Debug.Print "=== SYNC START: " & objFolder.Name & " ==="
    Debug.Print "    Ordner     : " & objFolder.FolderPath
    Debug.Print "    Elemente   : " & lngGesamt
    Debug.Print "    Max. Mails : " & lngMaxMails
    Debug.Print "    Projekt    : " & strProjekt
    Debug.Print "    Phase      : " & strPhase
    Debug.Print "    Modus      : Extract-Buffer-Flush (v0.4, Direct DAO)"
    Debug.Print "    Backend    : " & BackendStatus()
    Debug.Print "    " & Now()
    Debug.Print String(70, "=")

    ' Effektive Speicherorte transparent ausgeben (Temp/Export/Backend/WorkerDB).
    DevPfadDiagnose "SyncFolder Start", False

    ' Backend-Verfuegbarkeit pruefen (mit Netzwerk-Recovery)
    If Not IstBackendVerfuegbar() Then
        LogWarn "SyncFolder: Backend nicht erreichbar - warte auf Netzwerk...", "SYNC"
        If Not WarteAufNetzwerk(GetBackendPfad()) Then
            LogError "SyncFolder: Backend nach Timeout nicht erreichbar - Abbruch", "SYNC"
            Exit Sub
        End If
        ' Nochmal pruefen nach Reconnect
        If Not IstBackendVerfuegbar() Then
            LogError "SyncFolder: Backend auch nach Reconnect nicht verfuegbar", "SYNC"
            Exit Sub
        End If
    End If

    Debug.Print "[SYNC] Starte Sync-Lauf..."
    ' Sync-Lauf in DB starten
    lngSyncLaufID = StarteSyncLauf(objFolder.FolderPath, strProjekt, strPhase)

    Debug.Print "[SYNC] Registriere Ordner in DB..."
    ' Ordner in DB registrieren
    lngOrdnerID = SpeichereOrdner(objFolder.Name, objFolder.FolderPath, 0, "", lngGesamt)

    ' Crash-Recovery: Sync-Zustand lokal markieren (ueberlebt Netzwerkverlust/Crash)
    SyncZustandMarkieren lngSyncLaufID, objFolder.FolderPath

    ' Config-Flags lesen (ueber Cache fuer Performance)
    Dim blnAnhaenge As Boolean
    Dim blnMSG      As Boolean
    Dim blnFilter   As Boolean
    blnAnhaenge = (CacheGetConfig(CFG_ANHAENGE, "1") = "1")
    blnMSG = (CacheGetConfig(CFG_MSG_EXPORT, "1") = "1")
    blnFilter = (CacheGetConfig(CFG_SIGNATUR_FILTER, "1") = "1")

    ' Export-Basispfad
    Dim strExportBase As String
    strExportBase = NormalisierePfad(CacheGetConfig(CFG_EXPORT_PFAD, _
                    Environ("USERPROFILE") & PATH_DEFAULT_FALLBACK))

    ' Projekt-Aufloesung: Name -> ProjektID (v0.6)
    Dim lngProjektID As Long: lngProjektID = 0
    If strProjekt <> "" Then
        lngProjektID = GetOderErstelleProjekt(strProjekt)
    End If

    Debug.Print "[SYNC] Initialisiere Buffer/Config..."
    ' --- BUFFER INITIALISIEREN ---
    BufferInit
    BufferSetzeKontext lngSyncLaufID, lngOrdnerID, strProjekt, strPhase, _
                       strExportBase, blnAnhaenge, blnMSG, blnFilter, lngProjektID
    ResetTempZaehler
    ResetCOMStatistik

    ' --- EXTRACT-RELEASE-BUFFER SCHLEIFE ---
    Dim lngConsecutiveErrors As Long: lngConsecutiveErrors = 0
    Const MAX_CONSECUTIVE_ERRORS As Long = 10  ' Notbremse: 10 Fehler hintereinander -> Abbruch

    For i = 1 To lngGesamt

        ' Dual-Access Worker: aktives Stop-Signal in globalen Sync-Abbruch ueberfuehren.
        If WorkerAktiverStopAngefordert() Then
            g_blnAbbrechen = True
            LogWarn "SyncFolder: Worker-Stop-Signal erkannt - Sync wird sauber beendet", "SYNC"
        End If

        ' Abbruch-Pruefung
        If g_blnAbbrechen Then
            LogWarn "Sync abgebrochen nach " & lngMailCount & " Mails", "SYNC"
            Exit For
        End If

        ' Throttling: alle 10 Mails kurze Atempause fuer Outlook
        If i Mod 10 = 0 Then
            DoEvents
            Sleep 10  ' 10ms Pause - Outlook-Messagepump kann arbeiten
        End If

        ' Heartbeat: zyklisch Laufstatus in Direktfenster ausgeben
        If SekundenDiff(dblHeartbeatLast, Timer) >= HEARTBEAT_S Then
            Debug.Print "  [HB] i=" & i & "/" & lngGesamt & _
                        " | ok=" & lngMailCount & _
                        " | err=" & lngFehler & _
                        " | retry=" & COMRetryZaehler() & _
                        " | rec=" & COMReconnectZaehler() & _
                        " | backendOff=" & IIf(g_blnBackendOffline, "1", "0")
            dblHeartbeatLast = Timer
        End If

        ' Crash-Recovery: Position alle 25 Mails lokal speichern
        If i Mod 25 = 0 Then
            SyncZustandAktualisieren i
        End If

        ' Netzwerk-Schutz: Periodisch pruefen ob Backend noch da
        If g_blnBackendOffline Then
            LogError "SyncFolder: Backend offline (Watchdog) - Sync pausiert", "SYNC"
            If Not PruefeBackendVorZugriff() Then
                LogError "SyncFolder: Backend-Reconnect fehlgeschlagen - Sync endet bei Mail " & i, "SYNC"
                SyncZustandAktualisieren i
                Exit For
            End If
        End If

        ' EXTRACT: Mail-Objekt holen (mit COM-Resilienz)
        dblT0 = Timer
        On Error Resume Next
        Set objItem = objFolder.Items(i)
        dblFetchItemsS = dblFetchItemsS + SekundenDiff(dblT0, Timer)
        If Err.Number <> 0 Or objItem Is Nothing Then
            Dim lngItemErr As Long: lngItemErr = Err.Number
            Err.Clear
            On Error GoTo ErrHandler

            ' Fehlerklassifikation: Outlook tot oder nur Item kaputt?
            Dim strItemKlasse As String
            strItemKlasse = KlassifiziereCOMFehler(lngItemErr)

            If strItemKlasse = "FATAL" Then
                ' Outlook ist weg -> Warten/Reconnect
                LogWarn "COM-Fatal bei Items(" & i & ") -> Warte auf Outlook", "SYNC"
                If WarteAufOutlook() Then
                    ' Ordner-Referenz ist nach Reconnect ungueltig!
                    LogError "Ordner-Referenz verloren nach Reconnect - Sync endet", "SYNC"
                    Exit For
                Else
                    LogError "Outlook nicht wiederherstellbar - Sync endet", "SYNC"
                    Exit For
                End If
            End If

            ' ITEM-Fehler: Mail ueberspring, Zaehler erhoehen
            lngFehler = lngFehler + 1
            lngConsecutiveErrors = lngConsecutiveErrors + 1

            ' Notbremse: Zu viele Fehler hintereinander?
            If lngConsecutiveErrors >= MAX_CONSECUTIVE_ERRORS Then
                LogError "Notbremse: " & MAX_CONSECUTIVE_ERRORS & _
                         " aufeinanderfolgende Fehler - Sync endet", "SYNC"
                Exit For
            End If

            GoTo NaechsteMail
        End If
        On Error GoTo ErrHandler

        ' Erfolg: Consecutive-Error-Zaehler zuruecksetzen
        lngConsecutiveErrors = 0

        ' Nur echte E-Mails (IPM.Note)
        dblT0 = Timer
        On Error Resume Next
        Dim strMsgClass As String
        strMsgClass = objItem.MessageClass
        dblMsgClassS = dblMsgClassS + SekundenDiff(dblT0, Timer)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo ErrHandler
            Set objItem = Nothing
            GoTo NaechsteMail
        End If
        On Error GoTo ErrHandler

        If Left(strMsgClass, 8) <> "IPM.Note" Then
            Set objItem = Nothing
            GoTo NaechsteMail
        End If

        lngMailCount = lngMailCount + 1

        ' Fortschritt (alle 25 Mails + erste)
        If lngMailCount Mod 25 = 0 Or lngMailCount = 1 Then
            Debug.Print "  [" & Format(lngMailCount, "000") & "/" & _
                        Format(lngGesamt, "000") & "] Extrahiere..."
        End If

        ' EXTRACT: Alle Daten aus COM-Objekt lesen
        dblT0 = Timer
        ExtrahiereKomplett objItem, mk, blnAnhaenge, blnMSG, blnFilter
        dblExtractS = dblExtractS + SekundenDiff(dblT0, Timer)

        ' RELEASE: COM-Objekt sofort freigeben (Outlook entlasten!)
        Set objItem = Nothing

        ' Gueltigkeitspruefung
        If Not mk.Mail.IstGueltig Then
            lngFehler = lngFehler + 1
            GoTo NaechsteMail
        End If

        ' BUFFER: In Schreib-Puffer aufnehmen
        dblT0 = Timer
        BufferHinzufuegen mk
        dblBufferAddS = dblBufferAddS + SekundenDiff(dblT0, Timer)

        ' FLUSH: Wenn Puffer voll, Batch in DB schreiben (Transaktion)
        If BufferIstVoll() Then
            dblT0 = Timer
            Debug.Print "  [FLUSH] Buffer voll bei Mail " & lngMailCount & " -> schreibe Batch..."
            BufferFlush
            dblFlushS = dblFlushS + SekundenDiff(dblT0, Timer)
            lngFlushCount = lngFlushCount + 1
            DoEvents  ' UI responsiv halten
        End If

        ' Maximum erreicht?
        If lngMailCount >= lngMaxMails Then
            LogInfo "Maximum erreicht (" & lngMaxMails & " Mails)", "SYNC"
            Exit For
        End If

NaechsteMail:
    Next i

    ' --- FINALER FLUSH (restliche Daten im Puffer) ---
    If WorkerAktiverStopAngefordert() Then g_blnAbbrechen = True

    dblT0 = Timer
    Debug.Print "[SYNC] Finaler BufferFlush..."
    BufferFlush
    dblFinalFlushS = SekundenDiff(dblT0, Timer)

    ' --- DATEI-QUEUE VERARBEITEN (Temp -> Netzwerk/Festplatte) ---
    Dim lngDateien As Long
    dblT0 = Timer
    Debug.Print "[SYNC] Verarbeite Datei-Queue..."
    lngDateien = DateiQueueVerarbeiten()
    dblDateiQueueS = SekundenDiff(dblT0, Timer)

    ' --- ERGEBNIS SAMMELN (VOR BufferLeeren!) ---
    Dim lngNeu As Long
    Dim lngDuplikate As Long
    Dim lngVerarbeitet As Long
    lngNeu = BufferGesamtNeu()
    lngDuplikate = BufferGesamtDuplikate()
    lngVerarbeitet = BufferGesamtVerarbeitet()
    lngFehler = lngFehler + BufferGesamtFehler()

    Dim strStatus As String
    If g_blnAbbrechen Then
        strStatus = "Abgebrochen"
    Else
        strStatus = "Abgeschlossen"
    End If

    dblT0 = Timer
    Debug.Print "[SYNC] Beende Sync-Lauf in DB..."
    Call BeendeSyncLauf(lngSyncLaufID, strStatus, lngVerarbeitet, lngNeu, lngDuplikate, lngFehler)
    dblBeendeSyncS = SekundenDiff(dblT0, Timer)

    ' --- AUFRAEUMEN ---
    dblT0 = Timer
    Debug.Print "[SYNC] Cleanup Temp/Buffer..."
    BereinigeTempDateien
    BufferLeeren

    ' Crash-Recovery: Sync sauber beendet -> Zustand loeschen
    SyncZustandLoeschen
    dblCleanupS = SekundenDiff(dblT0, Timer)

    Debug.Print String(70, "-")
    Debug.Print "=== SYNC ERGEBNIS (v0.4) ==="
    Debug.Print "  Verarbeitet : " & lngVerarbeitet
    Debug.Print "  Neu         : " & lngNeu
    Debug.Print "  Duplikate   : " & lngDuplikate
    Debug.Print "  Fehler      : " & lngFehler
    Debug.Print "  Dateien     : " & lngDateien & " (" & BufferGesamtDateien() & " kumuliert)"
    Debug.Print "  Dauer       : " & Format((Timer - g_dtSyncStart) / 60, "0.0") & " min"
    Debug.Print "  COM-Retries : " & COMRetryZaehler() & " | Reconnects: " & COMReconnectZaehler()
    Debug.Print "  --- Performance (Sekunden) ---"
    Debug.Print "  Fetch Items : " & Format(dblFetchItemsS, "0.000")
    Debug.Print "  MsgClass    : " & Format(dblMsgClassS, "0.000")
    Debug.Print "  Extraktion  : " & Format(dblExtractS, "0.000")
    Debug.Print "  BufferAdd   : " & Format(dblBufferAddS, "0.000")
    Debug.Print "  Flush(Loop) : " & Format(dblFlushS, "0.000") & " (" & lngFlushCount & "x)"
    Debug.Print "  Flush(Final): " & Format(dblFinalFlushS, "0.000")
    Debug.Print "  DateiQueue  : " & Format(dblDateiQueueS, "0.000")
    Debug.Print "  SyncEnde DB : " & Format(dblBeendeSyncS, "0.000")
    Debug.Print "  Cleanup     : " & Format(dblCleanupS, "0.000")
    If lngMailCount > 0 Then
        Debug.Print "  Extrakt/Mail: " & Format((dblExtractS / lngMailCount) * 1000, "0.0") & " ms"
        Debug.Print "  Fetch/Mail  : " & Format((dblFetchItemsS / lngMailCount) * 1000, "0.0") & " ms"
    End If
    Debug.Print "  Status      : " & strStatus
    TimerStop "SyncFolder"
    CacheStatus
    Debug.Print String(70, "=")

    ' --- Subfolder rekursiv verarbeiten ---
    If blnSubfolder And Not g_blnAbbrechen Then
        Call SyncSubfolder(objFolder, strProjekt, strPhase, lngMaxMails)
    End If

    Set objItem = Nothing
    Exit Sub

ErrHandler:
    HandleError "modSync", "SyncFolder"
    If lngSyncLaufID > 0 Then
        Call BeendeSyncLauf(lngSyncLaufID, "Fehler", 0, 0, 0, lngFehler + 1)
    End If
    SyncZustandLoeschen
    Set objItem = Nothing
    BufferLeeren
    BereinigeTempDateien
End Sub


' Liefert vergangene Sekunden robust auch ueber Mitternacht (Timer-Reset).
Private Function SekundenDiff(ByVal dblStart As Double, ByVal dblEnd As Double) As Double
    If dblEnd >= dblStart Then
        SekundenDiff = dblEnd - dblStart
    Else
        SekundenDiff = (86400# - dblStart) + dblEnd
    End If
End Function


' ===========================================================================
' ORDNERSTRUKTUR IN DB EINLESEN (mit StoreID v0.3)
' ===========================================================================

' Alle Postfaecher/Stores + Ordner rekursiv in tblOutlookOrdner speichern
Public Sub SyncOrdnerStruktur(Optional ByVal intMaxTiefe As Integer = 5)
    On Error GoTo ErrHandler

    If Not ConnectRDO() Then
        LogError "SyncOrdnerStruktur: Keine RDO-Verbindung", "SYNC"
        Exit Sub
    End If

    Debug.Print String(70, "=")
    Debug.Print "=== Ordnerstruktur einlesen (Tiefe=" & intMaxTiefe & ") ==="
    Debug.Print String(70, "=")

    Dim objStore As Object
    Dim objRoot As Object
    Dim lngRootID As Long
    Dim strStoreID As String

    For Each objStore In g_objRDO.Stores
        Set objRoot = objStore.RootFolder

        ' StoreID fuer eindeutige Identifikation
        On Error Resume Next
        strStoreID = objStore.StoreID
        If Err.Number <> 0 Then strStoreID = "": Err.Clear
        On Error GoTo ErrHandler

        Debug.Print "[STORE] " & objStore.DisplayName & _
                    IIf(strStoreID <> "", " (ID:" & Left(strStoreID, 20) & "...)", "")

        lngRootID = SpeichereOrdner(objRoot.Name, objRoot.FolderPath, 0, _
                                     objStore.DisplayName, 0, strStoreID)

        Call SyncOrdnerRekursiv(objRoot, lngRootID, objStore.DisplayName, _
                                 strStoreID, 1, intMaxTiefe)

        Set objRoot = Nothing
        Set objStore = Nothing
    Next objStore

    Debug.Print String(70, "=")
    Debug.Print "=== Ordnerstruktur eingelesen ==="
    Debug.Print String(70, "=")
    Set objRoot = Nothing
    Set objStore = Nothing
    Exit Sub

ErrHandler:
    HandleError "modSync", "SyncOrdnerStruktur"
    Set objRoot = Nothing
    Set objStore = Nothing
End Sub

' Rekursiver Ordner-Scan (v0.3: StoreID-Parameter)
Private Sub SyncOrdnerRekursiv(objParent As Object, _
                                ByVal lngParentID As Long, _
                                ByVal strPostfach As String, _
                                ByVal strStoreID As String, _
                                ByVal intLevel As Integer, _
                                ByVal intMaxDepth As Integer)
    If intLevel > intMaxDepth Then Exit Sub

    On Error Resume Next
    Dim objSub      As Object
    Dim lngElements As Long
    Dim lngID       As Long

    For Each objSub In objParent.Folders
        lngElements = objSub.Items.Count
        If Err.Number <> 0 Then lngElements = 0: Err.Clear

        lngID = SpeichereOrdner(objSub.Name, objSub.FolderPath, lngParentID, _
                                 strPostfach, lngElements, strStoreID)
        Debug.Print String(intLevel * 2, " ") & "|- " & objSub.Name & " (" & lngElements & ")"

        If intLevel < intMaxDepth Then
            Call SyncOrdnerRekursiv(objSub, lngID, strPostfach, strStoreID, _
                                    intLevel + 1, intMaxDepth)
        End If

        Set objSub = Nothing
    Next objSub
    Set objSub = Nothing
    On Error GoTo 0
End Sub


' ===========================================================================
' SUBFOLDER EINER MAIL-SYNC REKURSIV VERARBEITEN
' ===========================================================================

' Iteriert rekursiv ueber Unterordner und ruft SyncFolder fuer jeden auf.
' Stoppt bei g_blnAbbrechen = True.
Private Sub SyncSubfolder(objParent As Object, _
                           ByVal strProjekt As String, _
                           ByVal strPhase As String, _
                           ByVal lngMaxMails As Long, _
                           Optional ByVal intLevel As Integer = 1, _
                           Optional ByVal intMaxDepth As Integer = 5)
    If intLevel > intMaxDepth Then Exit Sub
    If g_blnAbbrechen Then Exit Sub

    On Error Resume Next
    Dim objSub As Object

    For Each objSub In objParent.Folders
        If g_blnAbbrechen Then Exit For
        If Err.Number <> 0 Then Err.Clear: GoTo NaechsterSub

        ' Nur Ordner mit E-Mails (nicht Kalender, Kontakte etc.)
        If objSub.DefaultItemType = olMail Then
            Debug.Print String(intLevel * 2, " ") & "|- SYNC: " & objSub.Name & _
                        " (" & objSub.Items.Count & " Elemente)"

            On Error GoTo 0
            Call SyncFolder(objSub, strProjekt, strPhase, lngMaxMails, False)
            On Error Resume Next

            ' Weitere Tiefe
            If intLevel < intMaxDepth Then
                Call SyncSubfolder(objSub, strProjekt, strPhase, lngMaxMails, _
                                    intLevel + 1, intMaxDepth)
            End If
        End If

NaechsterSub:
        Set objSub = Nothing
    Next objSub

    Set objSub = Nothing
    On Error GoTo 0
End Sub


' ===========================================================================
' GANZES POSTFACH SYNCHRONISIEREN
' ===========================================================================

' Synchronisiert alle Ordner eines Postfachs inkl. Unterordner.
' Aufruf: SyncPostfach "Torsten.Kugler@rps.bwl.de", "FLIWAS", "Test"
Public Sub SyncPostfach(ByVal strPostfachName As String, _
                         Optional ByVal strProjekt As String = "Standard", _
                         Optional ByVal strPhase As String = "Standard", _
                         Optional ByVal lngMaxMailsProOrdner As Long = 0, _
                         Optional ByVal intMaxTiefe As Integer = 5)
    On Error GoTo ErrHandler

    If Not ConnectRDO() Then
        LogError "SyncPostfach: Keine RDO-Verbindung", "SYNC"
        Exit Sub
    End If

    If lngMaxMailsProOrdner <= 0 Then
        lngMaxMailsProOrdner = CLng(CacheGetConfig(CFG_MAX_MAILS, "500"))
    End If

    ' Store finden
    Dim objStore As Object
    Dim blnGefunden As Boolean: blnGefunden = False

    For Each objStore In g_objRDO.Stores
        If InStr(1, objStore.DisplayName, strPostfachName, vbTextCompare) > 0 Then
            blnGefunden = True
            Exit For
        End If
    Next objStore

    If Not blnGefunden Then
        LogError "SyncPostfach: Store '" & strPostfachName & "' nicht gefunden", "SYNC"
        Exit Sub
    End If

    Debug.Print String(70, "=")
    Debug.Print "=== POSTFACH-SYNC: " & objStore.DisplayName & " ==="
    Debug.Print "    Tiefe     : " & intMaxTiefe
    Debug.Print "    Max/Ordner: " & lngMaxMailsProOrdner
    Debug.Print "    Backend   : " & BackendStatus()
    Debug.Print "    " & Now()
    Debug.Print String(70, "=")

    g_blnAbbrechen = False
    g_dtSyncStart = Timer

    ' Root-Ordner und rekursiv alle Unterordner synchronisieren
    Dim objRoot As Object
    Set objRoot = objStore.RootFolder

    Call SyncSubfolder(objRoot, strProjekt, strPhase, lngMaxMailsProOrdner, 1, intMaxTiefe)

    Debug.Print String(70, "=")
    Debug.Print "=== POSTFACH-SYNC ABGESCHLOSSEN ==="
    Debug.Print "    Dauer: " & Format((Timer - g_dtSyncStart) / 60, "0.0") & " min"
    Debug.Print String(70, "=")

    Set objRoot = Nothing
    Set objStore = Nothing
    Exit Sub

ErrHandler:
    HandleError "modSync", "SyncPostfach"
    Set objRoot = Nothing
    Set objStore = Nothing
End Sub


' ===========================================================================
' PROFIL-BASIERTER SYNC (v0.3)
' ===========================================================================

' Sync-Profil erstellen -> gibt ProfilID zurueck
' Ein Profil speichert Projekt/Phase/MaxMails/Tiefe fuer wiederholte Nutzung.
'
' Aufruf:
'   Dim id As Long
'   id = ErstelleSyncProfil("FLIWAS_Prod", "FLIWAS", "Produktion")
'   ProfilOrdnerHinzufuegen id, "Torsten.Kugler@rps.bwl.de\Posteingang\FLIWAS"
'   SyncMitProfil "FLIWAS_Prod"
Public Function ErstelleSyncProfil(ByVal strName As String, _
                                    ByVal strProjekt As String, _
                                    ByVal strPhase As String, _
                                    Optional ByVal lngMaxMails As Long = 500, _
                                    Optional ByVal intMaxTiefe As Integer = 5, _
                                    Optional ByVal strExportPfad As String = "", _
                                    Optional ByVal strBeschreibung As String = "") As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset

    Set db = CurrentDb

    ' Pruefen ob Profil bereits existiert
    Dim varID As Variant
    varID = DLookup("ProfilID", TBL_SYNC_PROFIL, "ProfilName='" & SQLSafe(strName) & "'")
    If Not IsNull(varID) Then
        LogWarn "Sync-Profil '" & strName & "' existiert bereits (ID=" & varID & ")", "SYNC"
        ErstelleSyncProfil = CLng(varID)
        Exit Function
    End If

    Set rs = db.OpenRecordset(TBL_SYNC_PROFIL, dbOpenDynaset)
    With rs
        .AddNew
        !ProfilName = Left(strName, 100)
        !Beschreibung = Left(Nz(strBeschreibung, ""), 255)
        !IstAktiv = True
        !Projekt = Left(strProjekt, 100)
        !Phase = Left(strPhase, 100)
        !MaxMailsProOrdner = lngMaxMails
        !MaxTiefe = intMaxTiefe
        !ExportPfad = Left(Nz(strExportPfad, ""), 255)
        !ErstelltAm = Now
        .Update
        .Bookmark = .LastModified
        ErstelleSyncProfil = !ProfilID
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    LogInfo "Sync-Profil erstellt: '" & strName & "' (ID=" & ErstelleSyncProfil & ")", "SYNC"
    Exit Function

ErrHandler:
    HandleError "modSync", "ErstelleSyncProfil"
    ErstelleSyncProfil = 0
End Function


' Ordner zu einem Sync-Profil hinzufuegen
Public Sub ProfilOrdnerHinzufuegen(ByVal lngProfilID As Long, _
                                    ByVal strOrdnerPfad As String, _
                                    Optional ByVal strPostfach As String = "")
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset(TBL_SYNC_PROFIL_ORDNER, dbOpenDynaset)

    With rs
        .AddNew
        !ProfilID = lngProfilID
        !OrdnerPfad = Left(strOrdnerPfad, 255)
        !PostfachName = Left(Nz(strPostfach, ""), 255)
        !IstAktiv = True
        .Update
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    LogDebug "Profil-Ordner hinzugefuegt: ProfilID=" & lngProfilID & _
             " Pfad=" & strOrdnerPfad, "SYNC"
    Exit Sub

ErrHandler:
    HandleError "modSync", "ProfilOrdnerHinzufuegen"
End Sub


' Synchronisation anhand eines gespeicherten Profils ausfuehren
' Liest Profil-Einstellungen + Ordnerliste und synchronisiert jeden Ordner.
Public Sub SyncMitProfil(ByVal strProfilName As String)
    On Error GoTo ErrHandler

    If Not ConnectRDO() Then
        LogError "SyncMitProfil: Keine RDO-Verbindung", "SYNC"
        Exit Sub
    End If

    ' Profil laden
    Dim db As DAO.Database, rsProfil As DAO.Recordset
    Set db = CurrentDb
    Set rsProfil = db.OpenRecordset( _
        "SELECT * FROM [" & TBL_SYNC_PROFIL & "] WHERE ProfilName='" & SQLSafe(strProfilName) & "'" & _
        " AND IstAktiv=True", dbOpenSnapshot)

    If rsProfil.EOF Then
        LogError "Sync-Profil '" & strProfilName & "' nicht gefunden oder deaktiviert", "SYNC"
        rsProfil.Close: Set rsProfil = Nothing: Set db = Nothing
        Exit Sub
    End If

    Dim lngProfilID As Long
    Dim strProjekt As String
    Dim strPhase As String
    Dim lngMaxMails As Long
    Dim intMaxTiefe As Integer

    lngProfilID = rsProfil!ProfilID
    strProjekt = Nz(rsProfil!Projekt, "Standard")
    strPhase = Nz(rsProfil!Phase, "Standard")
    lngMaxMails = Nz(rsProfil!MaxMailsProOrdner, 500)
    intMaxTiefe = Nz(rsProfil!MaxTiefe, 5)
    rsProfil.Close: Set rsProfil = Nothing

    ' Export-Pfad (Profil-spezifisch oder Standard)
    Dim strExport As String
    strExport = Nz(DLookup("ExportPfad", TBL_SYNC_PROFIL, "ProfilID=" & lngProfilID), "")
    If strExport <> "" Then
        CacheSetConfig CFG_EXPORT_PFAD, strExport
    End If

    Debug.Print String(70, "=")
    Debug.Print "=== PROFIL-SYNC: " & strProfilName & " ==="
    Debug.Print "    Projekt    : " & strProjekt
    Debug.Print "    Phase      : " & strPhase
    Debug.Print "    Max/Ordner : " & lngMaxMails
    Debug.Print "    " & Now()
    Debug.Print String(70, "=")

    g_blnAbbrechen = False
    g_dtSyncStart = Timer

    ' Ordner des Profils laden und synchronisieren
    Dim rsOrdner As DAO.Recordset
    Set rsOrdner = db.OpenRecordset( _
        "SELECT OrdnerPfad, PostfachName FROM [" & TBL_SYNC_PROFIL_ORDNER & "] " & _
        "WHERE ProfilID=" & lngProfilID & " AND IstAktiv=True", dbOpenSnapshot)

    Dim lngOrdnerCount As Long: lngOrdnerCount = 0

    Do While Not rsOrdner.EOF
        If g_blnAbbrechen Then Exit Do

        Dim strPfad As String
        strPfad = Nz(rsOrdner!OrdnerPfad, "")

        If strPfad <> "" Then
            Debug.Print "  >>> Ordner: " & strPfad

            Dim objFolder As Object
            Set objFolder = OeffneOrdner(strPfad)

            If Not objFolder Is Nothing Then
                Call SyncFolder(objFolder, strProjekt, strPhase, lngMaxMails, True)
                lngOrdnerCount = lngOrdnerCount + 1
                Set objFolder = Nothing
            Else
                LogWarn "Profil-Ordner nicht gefunden: " & strPfad, "SYNC"
            End If
        End If

        rsOrdner.MoveNext
    Loop

    rsOrdner.Close: Set rsOrdner = Nothing: Set db = Nothing

    Debug.Print String(70, "=")
    Debug.Print "=== PROFIL-SYNC ABGESCHLOSSEN ==="
    Debug.Print "    Ordner    : " & lngOrdnerCount
    Debug.Print "    Dauer     : " & Format((Timer - g_dtSyncStart) / 60, "0.0") & " min"
    Debug.Print String(70, "=")
    Exit Sub

ErrHandler:
    HandleError "modSync", "SyncMitProfil"
End Sub


