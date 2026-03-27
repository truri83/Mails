Attribute VB_Name = "modSync"
Option Compare Database
Option Explicit

' ===========================================================================
' modSync - Synchronisations-Orchestrierung
' ===========================================================================
' Hauptmodul fuer die Outlook-zu-Access-Synchronisation.
' v0.4: Extract-Release-Process mit Direct DAO + Netzwerk-Resilienz.
'
' Architektur:
'   1. EXTRACT: Alle Daten aus RDO-Objekt lesen (modMailExtract)
'   2. RELEASE: COM-Objekt sofort freigeben (Outlook entlasten)
'   3. BUFFER:  In Speicher-Puffer sammeln (modAsyncBuffer)
'   4. FLUSH:   Batch in DB schreiben (Direct DAO, 1 Transaktion)
'   5. COPY:    Dateien Temp->Netzwerk (modFileManager, Retry+Recovery)
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
        lngMaxMails = CLng(LeseConfig("MaxMailsProSync", "500"))
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
        lngMaxMails = CLng(LeseConfig("MaxMailsProSync", "500"))
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

    ' --- Initialisierung ---
    InitGlobals
    g_blnAbbrechen = False
    g_dtSyncStart = Timer

    If Nz(strProjekt, "") = "" Then strProjekt = "Standard"
    If Nz(strPhase, "") = "" Then strPhase = "Standard"

    lngGesamt = objFolder.Items.Count
    lngMailCount = 0: lngFehler = 0

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

    ' Sync-Lauf in DB starten
    lngSyncLaufID = StarteSyncLauf(objFolder.FolderPath, strProjekt, strPhase)

    ' Ordner in DB registrieren
    lngOrdnerID = SpeichereOrdner(objFolder.Name, objFolder.FolderPath, 0, "", lngGesamt)

    ' Config-Flags lesen
    Dim blnAnhaenge As Boolean
    Dim blnMSG      As Boolean
    Dim blnFilter   As Boolean
    blnAnhaenge = (LeseConfig("AnhaengeExtrahieren", "1") = "1")
    blnMSG = (LeseConfig("MSGExportieren", "1") = "1")
    blnFilter = (LeseConfig("SignaturBilderFiltern", "1") = "1")

    ' Export-Basispfad
    Dim strExportBase As String
    strExportBase = NormalisierePfad(LeseConfig("ExportBasisPfad", _
                    Environ("USERPROFILE") & "\OutlookSync\"))

    ' --- BUFFER INITIALISIEREN ---
    BufferInit
    BufferSetzeKontext lngSyncLaufID, lngOrdnerID, strProjekt, strPhase, _
                       strExportBase, blnAnhaenge, blnMSG, blnFilter
    ResetTempZaehler

    ' --- EXTRACT-RELEASE-BUFFER SCHLEIFE ---
    For i = 1 To lngGesamt

        ' Abbruch-Pruefung
        If g_blnAbbrechen Then
            LogWarn "Sync abgebrochen nach " & lngMailCount & " Mails", "SYNC"
            Exit For
        End If

        ' EXTRACT: Mail-Objekt holen
        On Error Resume Next
        Set objItem = objFolder.Items(i)
        If Err.Number <> 0 Or objItem Is Nothing Then
            Err.Clear
            lngFehler = lngFehler + 1
            On Error GoTo ErrHandler
            GoTo NaechsteMail
        End If
        On Error GoTo ErrHandler

        ' Nur echte E-Mails (IPM.Note)
        If Left(objItem.MessageClass, 8) <> "IPM.Note" Then GoTo NaechsteMail

        lngMailCount = lngMailCount + 1

        ' Fortschritt (alle 25 Mails + erste)
        If lngMailCount Mod 25 = 0 Or lngMailCount = 1 Then
            Debug.Print "  [" & Format(lngMailCount, "000") & "/" & _
                        Format(lngGesamt, "000") & "] Extrahiere..."
        End If

        ' EXTRACT: Alle Daten aus COM-Objekt lesen
        ExtrahiereKomplett objItem, mk, blnAnhaenge, blnMSG, blnFilter

        ' RELEASE: COM-Objekt sofort freigeben (Outlook entlasten!)
        Set objItem = Nothing

        ' Gueltigkeitspruefung
        If Not mk.Mail.IstGueltig Then
            lngFehler = lngFehler + 1
            GoTo NaechsteMail
        End If

        ' BUFFER: In Schreib-Puffer aufnehmen
        BufferHinzufuegen mk

        ' FLUSH: Wenn Puffer voll, Batch in DB schreiben (Transaktion)
        If BufferIstVoll() Then
            BufferFlush
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
    BufferFlush

    ' --- DATEI-QUEUE VERARBEITEN (Temp -> Netzwerk/Festplatte) ---
    Dim lngDateien As Long
    lngDateien = DateiQueueVerarbeiten()

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

    Call BeendeSyncLauf(lngSyncLaufID, strStatus, lngVerarbeitet, lngNeu, lngDuplikate, lngFehler)

    ' --- AUFRAEUMEN ---
    BereinigeTempDateien
    BufferLeeren

    Debug.Print String(70, "-")
    Debug.Print "=== SYNC ERGEBNIS (v0.4) ==="
    Debug.Print "  Verarbeitet : " & lngVerarbeitet
    Debug.Print "  Neu         : " & lngNeu
    Debug.Print "  Duplikate   : " & lngDuplikate
    Debug.Print "  Fehler      : " & lngFehler
    Debug.Print "  Dateien     : " & lngDateien & " (" & BufferGesamtDateien() & " kumuliert)"
    Debug.Print "  Dauer       : " & Format((Timer - g_dtSyncStart) / 60, "0.0") & " min"
    Debug.Print "  Status      : " & strStatus
    Debug.Print String(70, "=")

    ' --- Subfolder rekursiv verarbeiten ---
    If blnSubfolder And Not g_blnAbbrechen Then
        Call SyncSubfolder(objFolder, strProjekt, strPhase, lngMaxMails)
    End If

    Set objItem = Nothing
    Exit Sub

ErrHandler:
    LogVBAError "SyncFolder"
    If lngSyncLaufID > 0 Then
        Call BeendeSyncLauf(lngSyncLaufID, "Fehler", 0, 0, 0, lngFehler + 1)
    End If
    BufferLeeren
    BereinigeTempDateien
End Sub


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
    Next objStore

    Debug.Print String(70, "=")
    Debug.Print "=== Ordnerstruktur eingelesen ==="
    Debug.Print String(70, "=")
    Exit Sub

ErrHandler:
    LogVBAError "SyncOrdnerStruktur"
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
    Next objSub
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
    Next objSub

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
        lngMaxMailsProOrdner = CLng(LeseConfig("MaxMailsProSync", "500"))
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
    Exit Sub

ErrHandler:
    LogVBAError "SyncPostfach"
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
    varID = DLookup("ProfilID", "tblSyncProfil", "ProfilName='" & SQLSafe(strName) & "'")
    If Not IsNull(varID) Then
        LogWarn "Sync-Profil '" & strName & "' existiert bereits (ID=" & varID & ")", "SYNC"
        ErstelleSyncProfil = CLng(varID)
        Exit Function
    End If

    Set rs = db.OpenRecordset("tblSyncProfil", dbOpenDynaset)
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
    LogVBAError "ErstelleSyncProfil"
    ErstelleSyncProfil = 0
End Function


' Ordner zu einem Sync-Profil hinzufuegen
Public Sub ProfilOrdnerHinzufuegen(ByVal lngProfilID As Long, _
                                    ByVal strOrdnerPfad As String, _
                                    Optional ByVal strPostfach As String = "")
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblSyncProfilOrdner", dbOpenDynaset)

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
    LogVBAError "ProfilOrdnerHinzufuegen"
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
        "SELECT * FROM tblSyncProfil WHERE ProfilName='" & SQLSafe(strProfilName) & "'" & _
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
    strExport = Nz(DLookup("ExportPfad", "tblSyncProfil", "ProfilID=" & lngProfilID), "")
    If strExport <> "" Then
        SchreibeConfig "ExportBasisPfad", strExport
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
        "SELECT OrdnerPfad, PostfachName FROM tblSyncProfilOrdner " & _
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
    LogVBAError "SyncMitProfil"
End Sub
