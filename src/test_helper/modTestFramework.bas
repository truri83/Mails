Option Compare Database
Option Explicit

' ===========================================================================
' modTestIntegration - End-to-End Integration Tests
' ===========================================================================
' Testet den kompletten Sync-Workflow mit echtem Outlook/Redemption.
'
' VORAUSSETZUNGEN:
'   - Outlook muss geoeffnet und eingeloggt sein
'   - Redemption DLL muss registriert sein
'   - Schema muss erstellt sein (ErstelleAlleTabellen)
'   - Mindestens 1 Mail im Posteingang
'
' AUFRUF:
'   RunIntegrationTests    ' Alle Integration-Tests
'
' HINWEIS:
'   Diese Tests lesen NUR aus Outlook — sie veraendern keine Mails.
'   DB-Testdaten werden am Ende bereinigt.
' ===========================================================================

Private Const MODUL_NAME As String = "modTestIntegration"
Private Const TEST_PREFIX As String = "INTTEST_"


' ===========================================================================
' ENTRY POINT
' ===========================================================================
Public Sub RunIntegrationTests()
    TestRunStart "INTEGRATION-TESTS (Outlook + DB)"

    Test_OutlookVerfuegbarkeit
    Test_RDOVerbindung
    Test_OrdnerZugriff
    Test_MailExtraktion
    Test_BufferWorkflow
    Test_HashDeduplikationLive
    Test_CacheStatistik

    ' Aufraeumen
    CleanupIntegrationDaten

    TestRunEnd
End Sub


' ===========================================================================
' OUTLOOK-VERBINDUNG
' ===========================================================================

Private Sub Test_OutlookVerfuegbarkeit()
    SuiteStart "Outlook Verfuegbarkeit"

    ' Outlook.Application erstellen
    Dim blnOK As Boolean
    blnOK = ConnectOutlook()

    If Not blnOK Then
        AssertSkip "Outlook nicht verfuegbar — restliche Tests werden uebersprungen"
        SuiteEnd
        Exit Sub
    End If

    AssertIsTrue blnOK, "ConnectOutlook() erfolgreich"
    AssertIsTrue IstOutlookAktiv(), "IstOutlookAktiv() = True"

    ' Namespace pruefen
    On Error Resume Next
    Dim objNS As Object
    Set objNS = g_objOutlook.GetNamespace("MAPI")
    AssertIsTrue (Not objNS Is Nothing), "MAPI Namespace erreichbar"
    Set objNS = Nothing
    On Error GoTo 0

    SuiteEnd
End Sub


Private Sub Test_RDOVerbindung()
    SuiteStart "RDO/Redemption Verbindung"

    ' Pruefen ob Outlook laeuft
    If Not IstOutlookAktiv() Then
        AssertSkip "Outlook nicht verfuegbar"
        SuiteEnd
        Exit Sub
    End If

    ' RDO verbinden
    Dim blnOK As Boolean
    blnOK = ConnectRDO()

    If Not blnOK Then
        AssertSkip "Redemption nicht verfuegbar (DLL registriert?)"
        SuiteEnd
        Exit Sub
    End If

    AssertIsTrue blnOK, "ConnectRDO() erfolgreich"
    AssertIsTrue IstRDOAktiv(), "IstRDOAktiv() = True"

    ' RDO Session pruefen
    On Error Resume Next
    Dim strStore As String
    strStore = g_objRDO.Stores.DefaultStore.Name
    AssertIsNotEmpty strStore, "Default Store erreichbar: " & Left(strStore, 30)
    On Error GoTo 0

    SuiteEnd
End Sub


' ===========================================================================
' ORDNER-ZUGRIFF
' ===========================================================================

Private Sub Test_OrdnerZugriff()
    SuiteStart "Ordner-Zugriff"

    If Not IstRDOAktiv() Then
        AssertSkip "RDO nicht verbunden"
        SuiteEnd
        Exit Sub
    End If

    On Error Resume Next

    ' Posteingang oeffnen
    Dim objFolder As Object
    Set objFolder = g_objRDO.GetDefaultFolder(olFolderInbox)

    If objFolder Is Nothing Then
        AssertFail "Posteingang nicht erreichbar"
        SuiteEnd
        Exit Sub
    End If

    AssertIsTrue (Not objFolder Is Nothing), "Posteingang geoeffnet"
    AssertIsNotEmpty objFolder.Name, "Ordnername: " & objFolder.Name

    ' Elementanzahl
    Dim lngCount As Long
    lngCount = objFolder.Items.Count
    AssertGreaterThan lngCount, 0, "Posteingang hat " & lngCount & " Elemente"

    ' Gesendete Elemente
    Dim objSent As Object
    Set objSent = g_objRDO.GetDefaultFolder(olFolderSentMail)
    AssertIsTrue (Not objSent Is Nothing), "Gesendete Elemente erreichbar"

    Set objFolder = Nothing
    Set objSent = Nothing
    On Error GoTo 0

    SuiteEnd
End Sub


' ===========================================================================
' MAIL-EXTRAKTION (Kerntest)
' ===========================================================================

Private Sub Test_MailExtraktion()
    SuiteStart "Mail-Extraktion (ExtrahiereKomplett)"

    If Not IstRDOAktiv() Then
        AssertSkip "RDO nicht verbunden"
        SuiteEnd
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ' Erste Mail aus Posteingang holen
    Dim objFolder As Object
    Set objFolder = g_objRDO.GetDefaultFolder(olFolderInbox)

    If objFolder.Items.Count = 0 Then
        AssertSkip "Posteingang ist leer"
        Set objFolder = Nothing
        SuiteEnd
        Exit Sub
    End If

    Dim objMail As Object
    Set objMail = objFolder.Items(1)

    ' Nur Mails verarbeiten (keine Termine, Aufgaben etc.)
    If objMail.MessageClass <> "IPM.Note" Then
        ' Naechste Mail versuchen
        Dim i As Long
        Dim blnFound As Boolean
        blnFound = False
        For i = 1 To IIf(objFolder.Items.Count < 20, objFolder.Items.Count, 20)
            Set objMail = objFolder.Items(i)
            If Left(objMail.MessageClass, 8) = "IPM.Note" Then
                blnFound = True
                Exit For
            End If
            Set objMail = Nothing
        Next i
        If Not blnFound Then
            AssertSkip "Keine IPM.Note Mail in den ersten 20 Elementen"
            Set objFolder = Nothing
            SuiteEnd
            Exit Sub
        End If
    End If

    ' ExtrahiereKomplett aufrufen (Sub, kein Function-Aufruf!)
    Dim mk As TypMailKomplett
    ExtrahiereKomplett objMail, mk, False, False, False

    ' Ergebnisse pruefen
    AssertIsNotEmpty mk.Mail.Betreff, "Betreff extrahiert: " & Left(mk.Mail.Betreff, 40)
    AssertIsNotEmpty mk.Mail.AbsenderEmail, "Absender-Email extrahiert"
    AssertIsTrue mk.Mail.EmpfangenAm > #1/1/2000#, "Empfangsdatum plausibel"
    AssertIsTrue mk.Mail.Groesse > 0, "Groesse > 0 (" & mk.Mail.Groesse & " Bytes)"
    AssertIsNotEmpty mk.Mail.InternetMessageID, "InternetMessageID vorhanden"

    ' Hash generieren
    Dim strHash As String
    strHash = GeneriereMailHash(mk.Mail.Betreff, mk.Mail.AbsenderEmail, _
                               mk.Mail.DisplayTo, mk.Mail.EmpfangenAm)
    AssertAreEqual 64, Len(strHash), "Mail-Hash generiert (64 Zeichen)"
    mk.UniqueHash = strHash

    ' HTML/PlainText Body
    ' (mindestens einer sollte nicht leer sein)
    Dim blnHatBody As Boolean
    blnHatBody = (Len(Nz(mk.Mail.HTMLBody, "")) > 0) Or (Len(Nz(mk.Mail.PlainTextBody, "")) > 0)
    AssertIsTrue blnHatBody, "Body vorhanden (HTML oder PlainText)"

    ' Empfaenger
    AssertGreaterThan mk.EmpfaengerAnzahl, 0, "Mindestens 1 Empfaenger (" & mk.EmpfaengerAnzahl & ")"

    ' COM-Objekt freigeben
    Set objMail = Nothing
    Set objFolder = Nothing

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Extraktion Fehler: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    Set objMail = Nothing
    Set objFolder = Nothing
    On Error GoTo 0
    SuiteEnd
End Sub


' ===========================================================================
' BUFFER WORKFLOW (Extract -> Buffer -> Flush -> Verify)
' ===========================================================================

Private Sub Test_BufferWorkflow()
    SuiteStart "Buffer Workflow (E2E)"

    If Not IstRDOAktiv() Then
        AssertSkip "RDO nicht verbunden"
        SuiteEnd
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ' --- Setup ---
    InitGlobals

    ' SyncLauf starten
    Dim lngSyncID As Long
    lngSyncID = StarteSyncLauf(TEST_PREFIX & "\\TestInbox", TEST_PREFIX & "Projekt", TEST_PREFIX & "Phase")
    AssertGreaterThan lngSyncID, 0, "SyncLauf gestartet"

    ' Ordner registrieren
    Dim lngOrdnerID As Long
    lngOrdnerID = SpeichereOrdner(TEST_PREFIX & "TestInbox", TEST_PREFIX & "\\TestInbox")
    AssertGreaterThan lngOrdnerID, 0, "Ordner registriert"

    ' Buffer initialisieren
    BufferInit 10  ' Kleiner Buffer fuer Test
    BufferSetzeKontext lngSyncID, lngOrdnerID, _
                       TEST_PREFIX & "Projekt", TEST_PREFIX & "Phase", _
                       "", False, False, False

    ' --- 3 Mails extrahieren und puffern ---
    Dim objFolder As Object
    Set objFolder = g_objRDO.GetDefaultFolder(olFolderInbox)

    Dim lngMailCount As Long
    lngMailCount = IIf(objFolder.Items.Count < 3, objFolder.Items.Count, 3)

    If lngMailCount = 0 Then
        AssertSkip "Posteingang leer"
        Set objFolder = Nothing
        SuiteEnd
        Exit Sub
    End If

    Dim i As Long
    Dim lngGepuffert As Long
    lngGepuffert = 0

    For i = 1 To lngMailCount
        On Error Resume Next
        Dim objMail As Object
        Set objMail = objFolder.Items(i)

        ' Nur Mails
        If Left(Nz(objMail.MessageClass, ""), 8) = "IPM.Note" Then
            Dim mk As TypMailKomplett
            ExtrahiereKomplett objMail, mk, False, False, False

            If mk.Mail.IstGueltig Then
                mk.UniqueHash = GeneriereMailHash(mk.Mail.Betreff, mk.Mail.AbsenderEmail, _
                                                  mk.Mail.DisplayTo, mk.Mail.EmpfangenAm)
                BufferHinzufuegen mk
                lngGepuffert = lngGepuffert + 1
            End If
        End If

        Set objMail = Nothing
        On Error GoTo ErrHandler
    Next i

    Set objFolder = Nothing

    AssertGreaterThan lngGepuffert, 0, lngGepuffert & " Mails gepuffert"
    AssertAreEqual CLng(lngGepuffert), CLng(BufferGroesse()), "BufferGroesse = " & lngGepuffert

    ' --- Flush ---
    Dim lngFlushed As Long
    lngFlushed = BufferFlush()
    AssertIsTrue (lngFlushed >= 0), "BufferFlush ausgefuehrt (" & lngFlushed & " neue)"

    ' Buffer sollte jetzt leer sein
    AssertAreEqual 0, CLng(BufferGroesse()), "Buffer nach Flush leer"

    ' --- Statistik pruefen ---
    Dim lngVerarbeitet As Long
    lngVerarbeitet = BufferGesamtVerarbeitet()
    AssertGreaterThan lngVerarbeitet, 0, "Verarbeitet > 0 (" & lngVerarbeitet & ")"

    ' --- SyncLauf beenden ---
    BeendeSyncLauf lngSyncID, "Abgeschlossen", lngGepuffert, lngFlushed, _
                   lngGepuffert - lngFlushed, 0

    ' --- Verify: Daten in DB ---
    ' Mindestens ein SyncLauf-Eintrag mit unserem Prefix
    Dim lngDBCount As Long
    lngDBCount = DCount("*", TBL_SYNC_LAUF, "OrdnerPfad LIKE '" & TEST_PREFIX & "*'")
    AssertGreaterThan lngDBCount, 0, "SyncLauf in DB vorhanden"

    BufferLeeren
    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Buffer-Workflow Fehler: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    BufferLeeren
    Set objFolder = Nothing
    On Error GoTo 0
    SuiteEnd
End Sub


' ===========================================================================
' HASH-DEDUPLIKATION LIVE
' ===========================================================================

Private Sub Test_HashDeduplikationLive()
    SuiteStart "Hash-Deduplikation (Live)"

    If Not IstRDOAktiv() Then
        AssertSkip "RDO nicht verbunden"
        SuiteEnd
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ' Gleiche Mail zweimal extrahieren -> gleicher Hash
    Dim objFolder As Object
    Set objFolder = g_objRDO.GetDefaultFolder(olFolderInbox)

    If objFolder.Items.Count = 0 Then
        AssertSkip "Posteingang leer"
        Set objFolder = Nothing
        SuiteEnd
        Exit Sub
    End If

    ' Erste Mail finden
    Dim objMail As Object
    Dim i As Long
    For i = 1 To IIf(objFolder.Items.Count < 20, objFolder.Items.Count, 20)
        Set objMail = objFolder.Items(i)
        If Left(Nz(objMail.MessageClass, ""), 8) = "IPM.Note" Then Exit For
        Set objMail = Nothing
    Next i

    If objMail Is Nothing Then
        AssertSkip "Keine IPM.Note in Inbox"
        Set objFolder = Nothing
        SuiteEnd
        Exit Sub
    End If

    ' Erste Extraktion
    Dim mk1 As TypMailKomplett
    ExtrahiereKomplett objMail, mk1, False, False, False
    Dim strHash1 As String
    strHash1 = GeneriereMailHash(mk1.Mail.Betreff, mk1.Mail.AbsenderEmail, _
                                mk1.Mail.DisplayTo, mk1.Mail.EmpfangenAm)

    ' Zweite Extraktion (gleiche Mail)
    Dim mk2 As TypMailKomplett
    ExtrahiereKomplett objMail, mk2, False, False, False
    Dim strHash2 As String
    strHash2 = GeneriereMailHash(mk2.Mail.Betreff, mk2.Mail.AbsenderEmail, _
                                mk2.Mail.DisplayTo, mk2.Mail.EmpfangenAm)

    AssertAreEqual strHash1, strHash2, "Gleiche Mail = gleicher Hash (Dedup-Sicherheit)"

    Set objMail = Nothing
    Set objFolder = Nothing

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Hash-Dedup Fehler: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    Set objMail = Nothing
    Set objFolder = Nothing
    On Error GoTo 0
    SuiteEnd
End Sub


' ===========================================================================
' CACHE-STATISTIK
' ===========================================================================

Private Sub Test_CacheStatistik()
    SuiteStart "Cache-Statistik"

    ' Nach den vorherigen Tests sollte der Cache Eintraege haben
    On Error Resume Next
    CacheStatus  ' Gibt Statistik im Debug.Print aus
    If Err.Number = 0 Then
        AssertIsTrue True, "CacheStatus ohne Fehler"
    Else
        AssertFail "CacheStatus Fehler: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    SuiteEnd
End Sub


' ===========================================================================
' CLEANUP
' ===========================================================================

Private Sub CleanupIntegrationDaten()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb

    Debug.Print ""
    Debug.Print "--- Cleanup: Integrations-Testdaten entfernen ---"

    ' Emails und abhaengige Daten
    db.Execute "DELETE FROM [" & TBL_EMAIL_ANHAENGE & "] WHERE EmailID IN " & _
               "(SELECT EmailID FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*'" & _
               " OR SyncLaufID IN (SELECT SyncLaufID FROM [" & TBL_SYNC_LAUF & "] WHERE OrdnerPfad LIKE '" & TEST_PREFIX & "*'))"
    db.Execute "DELETE FROM [" & TBL_EMAIL_EMPFAENGER & "] WHERE EmailID IN " & _
               "(SELECT EmailID FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*'" & _
               " OR SyncLaufID IN (SELECT SyncLaufID FROM [" & TBL_SYNC_LAUF & "] WHERE OrdnerPfad LIKE '" & TEST_PREFIX & "*'))"
    db.Execute "DELETE FROM [" & TBL_EMAIL_CONTENT & "] WHERE EmailID IN " & _
               "(SELECT EmailID FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*'" & _
               " OR SyncLaufID IN (SELECT SyncLaufID FROM [" & TBL_SYNC_LAUF & "] WHERE OrdnerPfad LIKE '" & TEST_PREFIX & "*'))"
    db.Execute "DELETE FROM [" & TBL_EMAIL_STATUS & "] WHERE EmailID IN " & _
               "(SELECT EmailID FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*'" & _
               " OR SyncLaufID IN (SELECT SyncLaufID FROM [" & TBL_SYNC_LAUF & "] WHERE OrdnerPfad LIKE '" & TEST_PREFIX & "*'))"

    db.Execute "DELETE FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*'" & _
               " OR SyncLaufID IN (SELECT SyncLaufID FROM [" & TBL_SYNC_LAUF & "] WHERE OrdnerPfad LIKE '" & TEST_PREFIX & "*')"

    db.Execute "DELETE FROM [" & TBL_EMAIL_THREADS & "] WHERE ThreadBetreff LIKE '" & TEST_PREFIX & "*'"
    db.Execute "DELETE FROM [" & TBL_OUTLOOK_ORDNER & "] WHERE OrdnerPfad LIKE '" & TEST_PREFIX & "*'"
    db.Execute "DELETE FROM [" & TBL_KONTAKTE & "] WHERE Email LIKE 'unittest_*' OR Anzeigename LIKE '" & TEST_PREFIX & "*'"
    db.Execute "DELETE FROM [" & TBL_SYNC_LAUF & "] WHERE OrdnerPfad LIKE '" & TEST_PREFIX & "*'"

    Set db = Nothing
    Debug.Print "    Integrations-Testdaten bereinigt."
    Debug.Print ""

    ' Outlook aufraumen
    DisconnectAll
    CleanupGlobals

    On Error GoTo 0
End Sub


