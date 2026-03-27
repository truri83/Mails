Option Compare Database
Option Explicit

' ===========================================================================
' modOrdner - Outlook-Ordner Synchronisation
' ===========================================================================
' v0.4.2: Scannt Outlook-Stores/Ordner und speichert sie in tblOutlookOrdner.
'
' Funktionen:
'   ScanneAlleOrdner()         - Hauptfunktion: Alle Stores durchlaufen
'   ScanneOrdnerRekursiv()     - Einzelnen Ordner + Unterordner scannen
'   IstStoreNutzbar()          - Store-Tauglichkeitspruefung
'   LeereOrdnerTabelle()       - tblOutlookOrdner zuruecksetzen
'
' Abhaengigkeiten:
'   modOutlookConnect  (ConnectOutlook, ConnectRDO, g_objOutlook)
'   modDAO             (SpeichereOrdner)
'   modGlobals         (g_objOutlook, g_blnAbbrechen)
'   modLogging         (LogInfo, LogWarn, LogError, LogVBAError)
' ===========================================================================


' ---------------------------------------------------------------------------
' Statistik fuer den aktuellen Scan-Lauf
' ---------------------------------------------------------------------------
Private m_lngOrdnerGesamt   As Long
Private m_lngOrdnerNeu      As Long
Private m_lngStoresGescannt As Long
Private m_lngStoresSkipped  As Long


' ===========================================================================
' HAUPTFUNKTION: Alle Outlook-Stores scannen
' ===========================================================================

' Liest alle Outlook-Postfaecher (Stores) und deren Ordner rekursiv ein.
' Speichert/aktualisiert jeden Ordner ueber modDAO.SpeichereOrdner().
'
' Parameter:
'   blnTabellLeeren  - True: tblOutlookOrdner vorher leeren (Default True)
'
' Rueckgabe: Anzahl gespeicherter Ordner (oder -1 bei Fehler)
Public Function ScanneAlleOrdner(Optional ByVal blnTabelleLeeren As Boolean = True) As Long
    On Error GoTo ErrHandler

    ' Statistik initialisieren
    m_lngOrdnerGesamt = 0
    m_lngOrdnerNeu = 0
    m_lngStoresGescannt = 0
    m_lngStoresSkipped = 0

    LogInfo "=== Outlook-Ordner-Scan gestartet ===", "ORDNER"

    ' Outlook sicherstellen
    If Not ConnectOutlook() Then
        LogError "Outlook nicht verfuegbar - Scan abgebrochen", "ORDNER"
        ScanneAlleOrdner = -1
        Exit Function
    End If

    ' Optional: Tabelle leeren
    If blnTabelleLeeren Then
        LeereOrdnerTabelle
    End If

    ' Namespace abrufen
    Dim objNamespace As Object
    Set objNamespace = g_objOutlook.GetNamespace("MAPI")

    Dim lngAnzahlStores As Long
    lngAnzahlStores = objNamespace.Stores.Count

    LogInfo "Erkannte Stores: " & lngAnzahlStores, "ORDNER"

    ' Alle Stores durchlaufen
    Dim objStore As Object
    Dim objRoot As Object

    For Each objStore In objNamespace.Stores
        ' Abbruch durch Benutzer?
        If g_blnAbbrechen Then
            LogWarn "Benutzerabbruch - Scan wird beendet", "ORDNER"
            Exit For
        End If
        DoEvents

        ' Store pruefen
        If Not IstStoreNutzbar(objStore) Then
            m_lngStoresSkipped = m_lngStoresSkipped + 1
            GoTo NaechsterStore
        End If

        ' RootFolder abrufen
        On Error Resume Next
        Set objRoot = objStore.GetRootFolder
        If Err.Number <> 0 Or objRoot Is Nothing Then
            LogWarn "Kein RootFolder fuer Store '" & objStore.DisplayName & "'", "ORDNER"
            Err.Clear
            m_lngStoresSkipped = m_lngStoresSkipped + 1
            On Error GoTo ErrHandler
            GoTo NaechsterStore
        End If
        On Error GoTo ErrHandler

        ' Rekursiv scannen
        LogDebug "Scanne Store: " & objStore.DisplayName, "ORDNER"
        ScanneOrdnerRekursiv objRoot, 0, objStore.DisplayName, 0

        m_lngStoresGescannt = m_lngStoresGescannt + 1

NaechsterStore:
        Set objRoot = Nothing
        Set objStore = Nothing
    Next objStore

    ' Statistik ausgeben
    LogInfo "=== Ordner-Scan abgeschlossen ===" & vbCrLf & _
            "  Stores gescannt: " & m_lngStoresGescannt & vbCrLf & _
            "  Stores uebersprungen: " & m_lngStoresSkipped & vbCrLf & _
            "  Ordner gesamt: " & m_lngOrdnerGesamt, "ORDNER"

    ScanneAlleOrdner = m_lngOrdnerGesamt

    Set objRoot = Nothing
    Set objStore = Nothing
    Set objNamespace = Nothing
    Exit Function

ErrHandler:
    HandleError "modOrdner", "ScanneAlleOrdner"
    Set objRoot = Nothing
    Set objStore = Nothing
    Set objNamespace = Nothing
    ScanneAlleOrdner = -1
End Function


' ===========================================================================
' REKURSIVER ORDNER-SCAN
' ===========================================================================

' Scannt einen Outlook-Ordner und alle Unterordner.
' Jeder Ordner wird ueber modDAO.SpeichereOrdner() gespeichert.
'
' Parameter:
'   objFolder     - Outlook-Ordner (RootFolder oder Subfolder)
'   lngParentID   - OrdnerID des Eltern-Ordners (0 fuer Root)
'   strPostfach   - Name des Postfachs/Store
'   iTiefe        - Aktuelle Rekursionstiefe (fuer Logging)
Public Sub ScanneOrdnerRekursiv(objFolder As Object, _
                                 ByVal lngParentID As Long, _
                                 ByVal strPostfach As String, _
                                 ByVal iTiefe As Long)
    On Error GoTo ErrHandler

    If objFolder Is Nothing Then Exit Sub

    ' Abbruch?
    If g_blnAbbrechen Then Exit Sub

    ' Ordnerdaten lesen (robust)
    Dim strName As String, strPfad As String
    Dim lngElemente As Long
    On Error Resume Next
    strName = objFolder.Name
    strPfad = objFolder.FolderPath
    lngElemente = objFolder.Items.Count
    If Err.Number <> 0 Then
        LogWarn "Ordner nicht lesbar (Tiefe " & iTiefe & "): " & Err.Description, "ORDNER"
        Err.Clear
        On Error GoTo ErrHandler
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' Ignorierte Ordner ausfiltern
    If IstIgnorierterOrdner(strName) Then
        LogTrace "Uebersprungen: " & strPfad, "ORDNER"
        Exit Sub
    End If

    ' StoreID (robust, manche Stores geben keine zurueck)
    Dim strStoreID As String
    On Error Resume Next
    strStoreID = objFolder.StoreID
    If Err.Number <> 0 Then strStoreID = "": Err.Clear
    On Error GoTo ErrHandler

    ' In DB speichern/aktualisieren (via modDAO)
    Dim lngOrdnerID As Long
    lngOrdnerID = SpeichereOrdner(strName, strPfad, lngParentID, _
                                   strPostfach, lngElemente, strStoreID)

    If lngOrdnerID > 0 Then
        m_lngOrdnerGesamt = m_lngOrdnerGesamt + 1
    End If

    ' Unterordner rekursiv verarbeiten
    Dim objSub As Object
    For Each objSub In objFolder.Folders
        DoEvents
        If g_blnAbbrechen Then Exit For
        ScanneOrdnerRekursiv objSub, lngOrdnerID, strPostfach, iTiefe + 1
        Set objSub = Nothing
    Next objSub

    Set objSub = Nothing
    Exit Sub

ErrHandler:
    HandleError "modOrdner", "ScanneOrdnerRekursiv", strName
    Set objSub = Nothing
End Sub


' ===========================================================================
' STORE-TAUGLICHKEITSPRUEFUNG
' ===========================================================================

' Prueft ob ein Outlook-Store gescannt werden sollte.
' Schliesst PST-Dateien und nicht erreichbare Stores aus.
'
' Rueckgabe: True wenn Store gescannt werden soll
Public Function IstStoreNutzbar(objStore As Object) As Boolean
    On Error Resume Next

    If objStore Is Nothing Then
        IstStoreNutzbar = False
        Exit Function
    End If

    ' PST-Dateien ausschliessen
    Dim strFilePath As String
    strFilePath = objStore.FilePath
    If Err.Number = 0 And Len(strFilePath) > 0 Then
        If UCase(Right(strFilePath, 4)) = ".PST" Then
            LogInfo "Store uebersprungen (PST): " & objStore.DisplayName, "ORDNER"
            IstStoreNutzbar = False
            Exit Function
        End If
    End If
    Err.Clear

    ' RootFolder-Zugriff testen
    Dim objTest As Object
    Set objTest = objStore.GetRootFolder
    If Err.Number <> 0 Or objTest Is Nothing Then
        LogWarn "Store nicht erreichbar: " & objStore.DisplayName & _
                " (" & Err.Description & ")", "ORDNER"
        Err.Clear
        IstStoreNutzbar = False
        Exit Function
    End If

    Set objTest = Nothing
    IstStoreNutzbar = True
End Function


' ===========================================================================
' HILFSFUNKTIONEN (privat)
' ===========================================================================

' Tabelle tblOutlookOrdner leeren
Public Sub LeereOrdnerTabelle()
    On Error GoTo ErrHandler
    CurrentDb.Execute "DELETE FROM [" & TBL_OUTLOOK_ORDNER & "]", dbFailOnError
    LogInfo TBL_OUTLOOK_ORDNER & " geleert", "ORDNER"
    Exit Sub

ErrHandler:
    HandleError "modOrdner", "LeereOrdnerTabelle"
End Sub


' Prueft ob ein Ordnername ignoriert werden soll
Private Function IstIgnorierterOrdner(ByVal strName As String) As Boolean
    Dim strLower As String
    strLower = LCase(Trim(Nz(strName, "")))

    ' Systemordner / nicht relevante Ordner
    If strLower Like "*oeffentliche ordner*" _
       Or strLower Like "*public folders*" _
       Or strLower Like "*favoriten*" _
       Or strLower Like "*favorites*" _
       Or strLower Like "*conversation history*" _
       Or strLower Like "*sync issues*" _
       Or strLower Like "*conflicts*" _
       Or strLower Like "*local failures*" _
       Or strLower Like "*server failures*" Then
        IstIgnorierterOrdner = True
    Else
        IstIgnorierterOrdner = False
    End If
End Function


