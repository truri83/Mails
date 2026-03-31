Option Compare Database
Option Explicit

' ===========================================================================
' modOutlookConnect - Outlook/Redemption Verbindungsverwaltung + Resilienz
' ===========================================================================
' v0.5.1: RedemptionLoader (DLL ohne COM-Registrierung)
' v0.4.3: COM-Robustheit + Selbstheilung fuer Massen-Mail-Sync
'
' VERBINDUNG:
'   ConnectOutlook()        -> Outlook.Application initialisieren
'   ConnectRDO()            -> Redemption.RDOSession + Logon
'   DisconnectAll()         -> Alle Verbindungen trennen
'
' DATEN:
'   ErstelleSafeMail()      -> SafeMailItem aus OOM-MailItem
'   GetSMTPFromEntry()      -> SMTP aus AddressEntry/Sender
'   GetSMTPFromRecipient()  -> SMTP aus Recipient-Objekt
'   GetAbsenderSMTP()       -> Absender-SMTP (RDO-Mail)
'   OeffneOrdner()          -> Outlook-Ordner per Pfad oeffnen
'
' RESILIENZ (v0.4.3):
'   IstOutlookAktiv()       -> Heartbeat-Check (OOM lebt?)
'   IstRDOAktiv()           -> Heartbeat-Check (RDO lebt?)
'   ReconnectOutlook()      -> Verbindung reparieren (mit Retry)
'   ReconnectRDO()          -> RDO-Session reparieren
'   KlassifiziereCOMFehler() -> TRANSIENT/FATAL/ITEM Klassifikation
'   WarteAufOutlook()       -> Wartet bis Outlook wieder reagiert
'   SichererCOMZugriff()    -> Property-Read mit auto. Retry+Reconnect
' ===========================================================================


' ---------------------------------------------------------------------------
' COM-FEHLER-KONSTANTEN (fuer Klassifikation)
' ---------------------------------------------------------------------------
' TRANSIENT (Outlook busy, Retry sinnvoll)
Private Const RPC_E_CALL_REJECTED        As Long = &H80010001
Private Const RPC_E_SERVERCALL_RETRYLATER As Long = &H8001010A

' FATAL (Outlook tot/getrennt, Reconnect noetig)
Private Const RPC_E_DISCONNECTED          As Long = &H80010108
Private Const RPC_E_SERVER_DIED           As Long = &H80010007
Private Const RPC_E_SERVER_DIED_DNE       As Long = &H80010012
Private Const CO_E_OBJNOTCONNECTED        As Long = &H800401FD
Private Const RPC_E_SERVER_UNAVAILABLE    As Long = -2147023174  ' RPC server unavailable

' ITEM (Einzelnes Element defekt, ueberspringen)
Private Const MAPI_E_NOT_FOUND            As Long = -2147221233
Private Const E_INVALIDARG                As Long = -2147024809
Private Const E_FAIL                      As Long = -2147467259  ' oft bei beschaedigten Mails
Private Const DISP_E_EXCEPTION            As Long = -2147352567  ' Automation Error (PropertyAccessor)

' ---------------------------------------------------------------------------
' RESILIENZ-KONFIGURATION
' ---------------------------------------------------------------------------
Private Const MAX_COM_RETRIES     As Integer = 5     ' Max Wiederholungen pro Aufruf
Private Const COM_RETRY_DELAY_MS  As Long = 500      ' Basis-Wartezeit zwischen Retries
Private Const OUTLOOK_WAIT_MAX_S  As Long = 30       ' Max Wartezeit auf Outlook
Private Const UI_YIELD_EVERY      As Long = 20       ' Alle N Schritte UI freigeben
Private Const UI_YIELD_SLEEP_MS   As Long = 1        ' Mini-Pause fuer Messagepump

' ---------------------------------------------------------------------------
' MODUL-VARIABLEN
' ---------------------------------------------------------------------------
Private m_lngCOMRetries     As Long   ' Gesamtzahl Retries (Statistik)
Private m_lngCOMReconnects  As Long   ' Gesamtzahl Reconnects (Statistik)
Private m_dictFolderCache   As Object ' Key: bereinigter Pfad -> "EntryID|StoreID"


' ---------------------------------------------------------------------------
' OUTLOOK.APPLICATION INITIALISIEREN (OOM)
' ---------------------------------------------------------------------------
Public Function ConnectOutlook() As Boolean
    On Error GoTo ErrHandler
    Dim strVer As String

    If Not g_objOutlook Is Nothing Then
        ' Pruefen ob Verbindung noch aktiv
        On Error Resume Next
        strVer = g_objOutlook.Version
        If Err.Number = 0 Then
            ConnectOutlook = True
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        Set g_objOutlook = Nothing
        On Error GoTo ErrHandler
    End If

    Set g_objOutlook = CreateObject("Outlook.Application")
    If g_objOutlook Is Nothing Then
        LogError "Outlook.Application konnte nicht erstellt werden", "CONNECT"
        ConnectOutlook = False
    Else
        LogInfo "Outlook " & g_objOutlook.Version & " verbunden", "CONNECT"
        ConnectOutlook = True
    End If
    Exit Function

ErrHandler:
    HandleError "modOutlookConnect", "ConnectOutlook"
    ConnectOutlook = False
End Function


' ---------------------------------------------------------------------------
' REDEMPTION RDOSESSION INITIALISIEREN
' ---------------------------------------------------------------------------
Public Function ConnectRDO() As Boolean
    On Error GoTo ErrHandler
    Dim strVer As String

    If Not g_objRDO Is Nothing Then
        ' Pruefen ob Session noch gueltig
        On Error Resume Next
        strVer = g_objRDO.Version
        If Err.Number = 0 Then
            ConnectRDO = True
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        g_objRDO.Logoff
        Set g_objRDO = Nothing
        On Error GoTo ErrHandler
    End If

    Set g_objRDO = ErstelleRedemptionObjekt("RDOSession")
    If g_objRDO Is Nothing Then
        LogError "Redemption.RDOSession nicht verfuegbar. " & _
                 "DLL pruefen oder regsvr32 """ & RDO_DLL_64 & """", "CONNECT"
        ConnectRDO = False
        Exit Function
    End If

    DoEvents
    g_objRDO.Logon
    DoEvents
    EnsureFolderCache
    LogInfo "Redemption " & g_objRDO.Version & " verbunden", "CONNECT"
    ConnectRDO = True
    Exit Function

ErrHandler:
    HandleError "modOutlookConnect", "ConnectRDO"
    ConnectRDO = False
End Function


' ---------------------------------------------------------------------------
' ALLE VERBINDUNGEN TRENNEN
' ---------------------------------------------------------------------------
Public Sub DisconnectAll()
    On Error Resume Next

    If Not g_objRDO Is Nothing Then
        g_objRDO.Logoff
        Set g_objRDO = Nothing
        LogDebug "RDO Logoff", "CONNECT"
    End If

    Set g_objOutlook = Nothing
    Set m_dictFolderCache = Nothing
    LogDebug "Outlook freigegeben", "CONNECT"

    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' REDEMPTION SAFEMAILITEM ERSTELLEN (OOM-Wrapper)
' ---------------------------------------------------------------------------
Public Function ErstelleSafeMail(objOOMMail As Object) As Object
    On Error GoTo ErrHandler

    If objOOMMail Is Nothing Then
        Set ErstelleSafeMail = Nothing
        Exit Function
    End If

    Dim objSafe As Object
    Set objSafe = ErstelleRedemptionObjekt("SafeMailItem")
    objSafe.Item = objOOMMail
    Set ErstelleSafeMail = objSafe
    Exit Function

ErrHandler:
    HandleError "modOutlookConnect", "ErstelleSafeMail"
    Set ErstelleSafeMail = Nothing
End Function


' ---------------------------------------------------------------------------
' SMTP-ADRESSE AUS ADDRESSENTRY / SENDER-OBJEKT HOLEN
' Loest interne Exchange-Adressen (/O=BWL/...) korrekt auf
' ---------------------------------------------------------------------------
Public Function GetSMTPFromEntry(objEntry As Object) As String
    On Error Resume Next
    Dim strResult As String
    Dim objExUser As Object

    If objEntry Is Nothing Then
        GetSMTPFromEntry = ""
        Exit Function
    End If

    ' Versuch 1: Direkt SMTPAddress
    strResult = objEntry.SMTPAddress
    If Err.Number <> 0 Then Err.Clear: strResult = ""

    ' Versuch 2: GetExchangeUser (fuer Exchange-Eintraege)
    If strResult = "" Or Left(strResult, 3) = "/O=" Then
        Set objExUser = objEntry.GetExchangeUser
        If Not objExUser Is Nothing Then
            strResult = objExUser.PrimarySmtpAddress
            If Err.Number <> 0 Then Err.Clear: strResult = ""
        End If
    End If

    ' Versuch 3: Fallback auf .Address
    If strResult = "" Or Left(strResult, 3) = "/O=" Then
        strResult = objEntry.Address
        If Err.Number <> 0 Then Err.Clear: strResult = ""
    End If

    Set objExUser = Nothing
    On Error GoTo 0
    GetSMTPFromEntry = strResult
End Function


' ---------------------------------------------------------------------------
' SMTP-ADRESSE AUS REDEMPTION RECIPIENT HOLEN
' ---------------------------------------------------------------------------
Public Function GetSMTPFromRecipient(objRecipient As Object) As String
    On Error Resume Next
    Dim strResult As String
    Dim objAE As Object

    If objRecipient Is Nothing Then
        GetSMTPFromRecipient = ""
        Exit Function
    End If

    Set objAE = objRecipient.AddressEntry
    If Not objAE Is Nothing Then
        strResult = objAE.SMTPAddress
        If Err.Number <> 0 Or strResult = "" Then
            Err.Clear
            strResult = objAE.Address
        End If
    End If

    If strResult = "" Then strResult = objRecipient.Address
    If Err.Number <> 0 Then Err.Clear

    Set objAE = Nothing
    On Error GoTo 0
    GetSMTPFromRecipient = Nz(strResult, "")
End Function


' ---------------------------------------------------------------------------
' ABSENDER-SMTP EINER RDO-MAIL ERMITTELN
' Behandelt sowohl SMTP- als auch EX-Absender korrekt
' ---------------------------------------------------------------------------
Public Function GetAbsenderSMTP(objRDOMail As Object) As String
    On Error Resume Next
    Dim strResult As String

    ' Pruefen ob Exchange-Adresse
    If objRDOMail.SenderEmailType = "EX" Then
        ' Ueber Sender-Objekt aufloesen
        strResult = GetSMTPFromEntry(objRDOMail.Sender)
    Else
        strResult = objRDOMail.SenderEmailAddress
    End If

    If Err.Number <> 0 Then Err.Clear: strResult = ""

    ' Fallback: MAPI Property PR_SMTP_ADDRESS
    If strResult = "" Or Left(strResult, 3) = "/O=" Then
        strResult = objRDOMail.Fields(PR_SMTP_ADDRESS)
        If Err.Number <> 0 Then Err.Clear: strResult = ""
    End If

    On Error GoTo 0
    If strResult = "" Then strResult = DEFAULT_EMAIL
    GetAbsenderSMTP = strResult
End Function


' ---------------------------------------------------------------------------
' OUTLOOK-ORDNER ANHAND PFAD OEFFNEN (via RDO)
' Format: "Postfachname\Ordner\Unterordner"
' ---------------------------------------------------------------------------
Public Function OeffneOrdner(ByVal strPfad As String) As Object
    On Error GoTo ErrHandler

    If Not ConnectRDO() Then
        Set OeffneOrdner = Nothing
        Exit Function
    End If

    ' Pfad bereinigen (zentralisiert in modStringUtils)
    strPfad = BereinigeOutlookPfad(strPfad)

    ' 1) Fast path: Cache -> GetFolderFromID
    Set OeffneOrdner = HoleOrdnerAusCache(strPfad)
    If Not OeffneOrdner Is Nothing Then Exit Function

    ' 2) Fast path: RDOSession.GetFolderFromPath (spart Store-Scan)
    Dim objByPath As Object
    On Error Resume Next
    Set objByPath = g_objRDO.GetFolderFromPath(strPfad)
    If objByPath Is Nothing Then
        Set objByPath = g_objRDO.GetFolderFromPath("\\" & strPfad)
    End If
    On Error GoTo ErrHandler
    If Not objByPath Is Nothing Then
        CacheOrdnerPfad strPfad, objByPath
        Set OeffneOrdner = objByPath
        Set objByPath = Nothing
        Exit Function
    End If

    Dim arrTeile() As String
    arrTeile = Split(strPfad, "\")

    If UBound(arrTeile) < 0 Then
        Set OeffneOrdner = Nothing
        Exit Function
    End If

    ' Store anhand des ersten Pfadteils finden
    Dim objStore As Object
    Dim objFolder As Object
    Dim i As Long
    Dim strStoreWanted As String
    Dim lngScan As Long
    strStoreWanted = arrTeile(0)

    For Each objStore In g_objRDO.Stores
        lngScan = lngScan + 1
        If (lngScan Mod UI_YIELD_EVERY) = 0 Then
            DoEvents
            Sleep UI_YIELD_SLEEP_MS
        End If
        If StrComp(objStore.DisplayName, strStoreWanted, vbTextCompare) = 0 Then
            Set objFolder = objStore.RootFolder
            Exit For
        End If
    Next objStore

    ' Fallback: unscharfer Store-Match (haeufig bei SMTP/Anzeigename-Abweichungen)
    If objFolder Is Nothing Then
        For Each objStore In g_objRDO.Stores
            lngScan = lngScan + 1
            If (lngScan Mod UI_YIELD_EVERY) = 0 Then
                DoEvents
                Sleep UI_YIELD_SLEEP_MS
            End If
            If InStr(1, objStore.DisplayName, strStoreWanted, vbTextCompare) > 0 Or _
               InStr(1, strStoreWanted, objStore.DisplayName, vbTextCompare) > 0 Then
                Set objFolder = objStore.RootFolder
                Exit For
            End If
        Next objStore
    End If

    ' Letzter Fallback: nur ein Teilpfad (ohne Store) wurde uebergeben
    If objFolder Is Nothing And UBound(arrTeile) >= 1 Then
        For Each objStore In g_objRDO.Stores
            lngScan = lngScan + 1
            If (lngScan Mod UI_YIELD_EVERY) = 0 Then
                DoEvents
                Sleep UI_YIELD_SLEEP_MS
            End If
            On Error Resume Next
            Set objFolder = objStore.RootFolder.Folders(arrTeile(1))
            On Error GoTo ErrHandler
            If Not objFolder Is Nothing Then Exit For
        Next objStore
        ' Pfadteile um eins nach links schieben, da Store fehlte
        If Not objFolder Is Nothing Then
            For i = 2 To UBound(arrTeile)
                If (i Mod 5) = 0 Then DoEvents
                If Trim(arrTeile(i)) <> "" Then
                    On Error Resume Next
                    Set objFolder = objFolder.Folders(arrTeile(i))
                    On Error GoTo ErrHandler
                    If objFolder Is Nothing Then Exit For
                End If
            Next i
            If Not objFolder Is Nothing Then
                CacheOrdnerPfad strPfad, objFolder
                Set OeffneOrdner = objFolder
                Set objStore = Nothing
                Exit Function
            End If
        End If
    End If

    If objFolder Is Nothing Then
        LogWarn "Store/Startordner fuer '" & strStoreWanted & "' nicht gefunden", "CONNECT"
        Set OeffneOrdner = Nothing
        Set objStore = Nothing
        Exit Function
    End If

    ' Durch Unterordner navigieren
    For i = 1 To UBound(arrTeile)
        If (i Mod 5) = 0 Then DoEvents
        If Trim(arrTeile(i)) <> "" Then
            On Error Resume Next
            Set objFolder = objFolder.Folders(arrTeile(i))
            On Error GoTo ErrHandler
            If objFolder Is Nothing Then
                LogWarn "Ordner '" & arrTeile(i) & "' nicht gefunden in: " & strPfad, "CONNECT"
                Set OeffneOrdner = Nothing
                Set objStore = Nothing
                Exit Function
            End If
        End If
    Next i

    CacheOrdnerPfad strPfad, objFolder
    Set OeffneOrdner = objFolder
    Set objStore = Nothing
    Exit Function

ErrHandler:
    HandleError "modOutlookConnect", "OeffneOrdner", strPfad
    Set OeffneOrdner = Nothing
    Set objStore = Nothing
End Function


' ---------------------------------------------------------------------------
' FOLDER-CACHE (Pfad -> EntryID/StoreID)
' ---------------------------------------------------------------------------
Private Sub EnsureFolderCache()
    If m_dictFolderCache Is Nothing Then
        Set m_dictFolderCache = CreateObject("Scripting.Dictionary")
        m_dictFolderCache.CompareMode = vbTextCompare
    End If
End Sub

Private Sub CacheOrdnerPfad(ByVal strPfad As String, objFolder As Object)
    On Error Resume Next
    EnsureFolderCache
    If objFolder Is Nothing Then Exit Sub

    Dim strEntryID As String
    Dim strStoreID As String
    strEntryID = Nz(objFolder.EntryID, "")
    strStoreID = Nz(objFolder.StoreID, "")
    If strEntryID = "" Then Exit Sub

    m_dictFolderCache(BereinigeOutlookPfad(strPfad)) = strEntryID & "|" & strStoreID
End Sub

Private Function HoleOrdnerAusCache(ByVal strPfad As String) As Object
    On Error GoTo ErrHandler
    EnsureFolderCache

    Dim strKey As String
    strKey = BereinigeOutlookPfad(strPfad)
    If Not m_dictFolderCache.Exists(strKey) Then
        Set HoleOrdnerAusCache = Nothing
        Exit Function
    End If

    Dim strPacked As String
    strPacked = CStr(m_dictFolderCache(strKey))
    If strPacked = "" Then
        Set HoleOrdnerAusCache = Nothing
        Exit Function
    End If

    Dim p As Long
    Dim strEntryID As String
    Dim strStoreID As String
    p = InStr(1, strPacked, "|", vbBinaryCompare)
    If p > 0 Then
        strEntryID = Left$(strPacked, p - 1)
        strStoreID = Mid$(strPacked, p + 1)
    Else
        strEntryID = strPacked
        strStoreID = ""
    End If

    If strEntryID = "" Then
        Set HoleOrdnerAusCache = Nothing
        Exit Function
    End If

    On Error Resume Next
    If strStoreID <> "" Then
        Set HoleOrdnerAusCache = g_objRDO.GetFolderFromID(strEntryID, strStoreID)
    Else
        Set HoleOrdnerAusCache = g_objRDO.GetFolderFromID(strEntryID)
    End If
    On Error GoTo ErrHandler

    If HoleOrdnerAusCache Is Nothing Then
        ' Cache-Eintrag ungueltig -> entfernen
        On Error Resume Next
        m_dictFolderCache.Remove strKey
        On Error GoTo ErrHandler
    End If
    Exit Function

ErrHandler:
    Set HoleOrdnerAusCache = Nothing
End Function


' ===========================================================================
' COM-RESILIENZ (v0.4.3)
' ===========================================================================


' ---------------------------------------------------------------------------
' HEARTBEAT: Ist Outlook noch am Leben?
' ---------------------------------------------------------------------------
' Prueft durch einen harmlosen Eigenschaftszugriff ob die COM-Verbindung
' zu Outlook noch funktioniert.
'
' Rueckgabe: True wenn Outlook antwortet, False wenn tot/nicht erreichbar
Public Function IstOutlookAktiv() As Boolean
    On Error Resume Next

    If g_objOutlook Is Nothing Then
        IstOutlookAktiv = False
        Exit Function
    End If

    ' Harmlosen Zugriff versuchen
    Dim strTest As String
    strTest = g_objOutlook.Version

    If Err.Number = 0 And Len(strTest) > 0 Then
        IstOutlookAktiv = True
    Else
        IstOutlookAktiv = False
        Err.Clear
    End If

    On Error GoTo 0
End Function


' ---------------------------------------------------------------------------
' HEARTBEAT: Ist RDO-Session noch gueltig?
' ---------------------------------------------------------------------------
Public Function IstRDOAktiv() As Boolean
    On Error Resume Next

    If g_objRDO Is Nothing Then
        IstRDOAktiv = False
        Exit Function
    End If

    Dim strTest As String
    strTest = g_objRDO.Version

    If Err.Number = 0 And Len(strTest) > 0 Then
        IstRDOAktiv = True
    Else
        IstRDOAktiv = False
        Err.Clear
    End If

    On Error GoTo 0
End Function


' ---------------------------------------------------------------------------
' RECONNECT: Outlook-Verbindung reparieren (mit Retry)
' ---------------------------------------------------------------------------
' Versucht MAX_RETRIES mal, eine neue Outlook-Instanz zu erstellen.
' Wartet mit progressivem Backoff zwischen den Versuchen.
'
' Rueckgabe: True wenn Verbindung wiederhergestellt
Public Function ReconnectOutlook() As Boolean
    On Error Resume Next

    LogWarn "=== Outlook-Reconnect gestartet ===", "CONNECT"

    ' Alte Instanz freigeben
    Set g_objOutlook = Nothing
    Err.Clear

    Dim i As Integer
    Dim strVer As String
    For i = 1 To MAX_RETRIES
        Sleep COM_RETRY_DELAY_MS * i   ' Progressiver Backoff
        DoEvents

        Set g_objOutlook = CreateObject("Outlook.Application")
        If Err.Number = 0 And Not g_objOutlook Is Nothing Then
            ' Verifizieren dass es wirklich funktioniert
            strVer = g_objOutlook.Version
            If Err.Number = 0 Then
                m_lngCOMReconnects = m_lngCOMReconnects + 1
                LogInfo "Outlook-Reconnect erfolgreich (Versuch " & i & _
                        ", Gesamt: " & m_lngCOMReconnects & ")", "CONNECT"
                On Error GoTo 0
                ReconnectOutlook = True
                Exit Function
            End If
        End If
        Err.Clear
        Set g_objOutlook = Nothing

        LogWarn "Reconnect-Versuch " & i & "/" & MAX_RETRIES & " fehlgeschlagen", "CONNECT"
    Next i

    On Error GoTo 0
    LogError "Outlook-Reconnect gescheitert nach " & MAX_RETRIES & " Versuchen", "CONNECT"
    ReconnectOutlook = False
End Function


' ---------------------------------------------------------------------------
' RECONNECT: RDO-Session reparieren
' ---------------------------------------------------------------------------
Public Function ReconnectRDO() As Boolean
    On Error Resume Next

    LogWarn "=== RDO-Reconnect gestartet ===", "CONNECT"

    ' Alte Session freigeben
    If Not g_objRDO Is Nothing Then
        g_objRDO.Logoff
        Set g_objRDO = Nothing
    End If
    Set m_dictFolderCache = Nothing
    Err.Clear

    Dim i As Integer
    For i = 1 To MAX_RETRIES
        Sleep COM_RETRY_DELAY_MS * i
        DoEvents

        Set g_objRDO = ErstelleRedemptionObjekt("RDOSession")
        If Err.Number = 0 And Not g_objRDO Is Nothing Then
            g_objRDO.Logon
            If Err.Number = 0 Then
                m_lngCOMReconnects = m_lngCOMReconnects + 1
                LogInfo "RDO-Reconnect erfolgreich (Versuch " & i & ")", "CONNECT"
                On Error GoTo 0
                ReconnectRDO = True
                Exit Function
            End If
        End If
        Err.Clear
        Set g_objRDO = Nothing

        LogWarn "RDO-Reconnect-Versuch " & i & "/" & MAX_RETRIES & " fehlgeschlagen", "CONNECT"
    Next i

    On Error GoTo 0
    LogError "RDO-Reconnect gescheitert", "CONNECT"
    ReconnectRDO = False
End Function


' ---------------------------------------------------------------------------
' FEHLER-KLASSIFIKATION: Transient vs. Fatal vs. Item-corrupt
' ---------------------------------------------------------------------------
' Klassifiziert einen COM-Fehler:
'   "TRANSIENT"  -> Retry lohnt sich (Outlook busy, kurzer RPC-Timeout)
'   "FATAL"      -> Outlook tot/getrennt, Reconnect noetig
'   "ITEM"       -> Einzelnes Element defekt, ueberspringen
'   "UNKNOWN"    -> Nicht klassifizierbar
Public Function KlassifiziereCOMFehler(ByVal lngErrNum As Long) As String
    Select Case lngErrNum
        ' --- TRANSIENT: Outlook ist busy, Retry sinnvoll ---
        Case RPC_E_CALL_REJECTED, _
             RPC_E_SERVERCALL_RETRYLATER
            KlassifiziereCOMFehler = "TRANSIENT"

        ' --- FATAL: Outlook-Prozess tot, Reconnect noetig ---
        Case RPC_E_DISCONNECTED, _
             RPC_E_SERVER_DIED, _
             RPC_E_SERVER_DIED_DNE, _
             CO_E_OBJNOTCONNECTED, _
             RPC_E_SERVER_UNAVAILABLE
            KlassifiziereCOMFehler = "FATAL"

        ' --- ITEM: MAPI-Fehler, einzelnes Element defekt ---
        Case MAPI_E_NOT_FOUND, _
             E_INVALIDARG, _
             E_FAIL, _
             DISP_E_EXCEPTION
            KlassifiziereCOMFehler = "ITEM"

        Case Else
            ' Heuristik: Negative Zahlen < -2Mrd sind COM-HRESULTs
            If lngErrNum < -2000000000 Then
                KlassifiziereCOMFehler = "TRANSIENT"  ' Im Zweifel retry
            Else
                KlassifiziereCOMFehler = "UNKNOWN"
            End If
    End Select
End Function


' ---------------------------------------------------------------------------
' WARTE AUF OUTLOOK: Blockiert bis Outlook wieder reagiert
' ---------------------------------------------------------------------------
' Wartet maximal OUTLOOK_WAIT_MAX_S Sekunden darauf, dass Outlook
' wieder auf COM-Aufrufe antwortet. Bei Tod: Reconnect-Versuch.
'
' Rueckgabe: True wenn Outlook (wieder) erreichbar
Public Function WarteAufOutlook() As Boolean
    Dim dtStart As Double
    dtStart = Timer

    LogWarn "Warte auf Outlook...", "CONNECT"

    Do While (Timer - dtStart) < OUTLOOK_WAIT_MAX_S
        DoEvents
        Sleep 200

        If IstOutlookAktiv() Then
            LogInfo "Outlook antwortet wieder (" & _
                    Format(Timer - dtStart, "0.0") & "s)", "CONNECT"
            WarteAufOutlook = True
            Exit Function
        End If
    Loop

    ' Timeout: Reconnect versuchen
    LogWarn "Outlook-Timeout nach " & OUTLOOK_WAIT_MAX_S & _
            "s - Reconnect...", "CONNECT"

    If ReconnectOutlook() Then
        ' Auch RDO erneuern (haengt von Outlook ab)
        ReconnectRDO
        WarteAufOutlook = True
    Else
        LogError "Outlook nicht wiederherstellbar", "CONNECT"
        WarteAufOutlook = False
    End If
End Function


' ---------------------------------------------------------------------------
' SICHERER COM-ZUGRIFF: Property-Read mit automatischem Retry
' ---------------------------------------------------------------------------
' Liest eine COM-Eigenschaft mit Retry bei transientem Fehler.
' Verwendet die Fehlerklassifikation:
'   TRANSIENT -> Sleep + Retry (bis MAX_COM_RETRIES)
'   FATAL     -> WarteAufOutlook() + Reconnect
'   ITEM      -> Sofort Default (Element defekt)
'
' Beispiele:
'   strBetreff = SichererCOMZugriff(objMail, "Subject", "")
'   dtDate = CDate(SichererCOMZugriff(objMail, "ReceivedTime", Now))
'
' Parameter:
'   objCOM       - COM-Objekt (Outlook.MailItem, RDOMail etc.)
'   strProperty  - Name der Eigenschaft
'   varDefault   - Fallback-Wert bei Fehler
Public Function SichererCOMZugriff(objCOM As Object, _
                                    ByVal strProperty As String, _
                                    Optional ByVal varDefault As Variant = "") As Variant
    On Error Resume Next

    If objCOM Is Nothing Then
        SichererCOMZugriff = varDefault
        On Error GoTo 0
        Exit Function
    End If

    ' Schneller Pfad: erster Versuch ohne Overhead
    Dim varResult As Variant
    varResult = CallByName(objCOM, strProperty, VbGet)

    If Err.Number = 0 Then
        SichererCOMZugriff = varResult
        On Error GoTo 0
        Exit Function
    End If

    ' Fehler -> Klassifizieren und reagieren
    Dim lngOrigErr As Long
    Dim strKlasse As String
    Dim i As Integer

    lngOrigErr = Err.Number
    Err.Clear

    strKlasse = KlassifiziereCOMFehler(lngOrigErr)

    Select Case strKlasse
        Case "ITEM"
            ' Element defekt: sofort Default zurueck
            LogDebug "COM Item-Fehler bei ." & strProperty & _
                     " (" & lngOrigErr & ") -> Default", "COM"
            SichererCOMZugriff = varDefault
            On Error GoTo 0
            Exit Function

        Case "FATAL"
            ' Outlook tot: Warten/Reconnect
            On Error GoTo 0
            If WarteAufOutlook() Then
                ' Nochmal versuchen (1x)
                On Error Resume Next
                varResult = CallByName(objCOM, strProperty, VbGet)
                If Err.Number = 0 Then
                    SichererCOMZugriff = varResult
                Else
                    SichererCOMZugriff = varDefault
                End If
                Err.Clear
                On Error GoTo 0
            Else
                SichererCOMZugriff = varDefault
            End If
            Exit Function

        Case "TRANSIENT", "UNKNOWN"
            ' Retry mit progressivem Backoff
            For i = 1 To MAX_COM_RETRIES
                m_lngCOMRetries = m_lngCOMRetries + 1
                Sleep COM_RETRY_DELAY_MS * i
                DoEvents

                varResult = CallByName(objCOM, strProperty, VbGet)
                If Err.Number = 0 Then
                    If i > 1 Then
                        LogDebug "COM-Retry erfolgreich nach " & i & _
                                 " Versuchen (." & strProperty & ")", "COM"
                    End If
                    SichererCOMZugriff = varResult
                    On Error GoTo 0
                    Exit Function
                End If

                lngOrigErr = Err.Number
                Err.Clear

                ' Eskalation: Nach Haelfte der Retries -> Fatal pruefen
                If i = MAX_COM_RETRIES \ 2 Then
                    strKlasse = KlassifiziereCOMFehler(lngOrigErr)
                    If strKlasse = "FATAL" Then
                        On Error GoTo 0
                        If Not WarteAufOutlook() Then
                            SichererCOMZugriff = varDefault
                            Exit Function
                        End If
                        On Error Resume Next
                    End If
                End If
            Next i
    End Select

    ' Alle Retries erschoepft
    LogWarn "COM-Zugriff ." & strProperty & " gescheitert nach " & _
            MAX_COM_RETRIES & " Retries (Err=" & lngOrigErr & ")", "COM"
    SichererCOMZugriff = varDefault
    On Error GoTo 0
End Function


' ---------------------------------------------------------------------------
' COM-STATISTIK
' ---------------------------------------------------------------------------
Public Function COMRetryZaehler() As Long
    COMRetryZaehler = m_lngCOMRetries
End Function

Public Function COMReconnectZaehler() As Long
    COMReconnectZaehler = m_lngCOMReconnects
End Function

Public Sub ResetCOMStatistik()
    m_lngCOMRetries = 0
    m_lngCOMReconnects = 0
End Sub


