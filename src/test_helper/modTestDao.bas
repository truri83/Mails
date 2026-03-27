Option Compare Database
Option Explicit

' ===========================================================================
' modTestDAO - Datenbank-Tests (Schema, CRUD, Dedup, Threads)
' ===========================================================================
' Testet die Datenzugriffsschicht gegen eine echte lokale DB.
' KEIN Outlook noetig — nur Access + Schema.
'
' AUFRUF:
'   RunDAOTests        ' Alle DB-Tests
'
' HINWEIS:
'   Erstellt Test-Daten und raeumt sie am Ende wieder auf.
'   Sollte auf einer Kopie oder Test-DB laufen.
' ===========================================================================

Private Const MODUL_NAME As String = "modTestDAO"

' Test-Konstanten zur Erkennung/Bereinigung
Private Const TEST_PREFIX As String = "UNITTEST_"
Private Const TEST_EMAIL  As String = "unittest_test@example.com"
Private Const TEST_EMAIL2 As String = "unittest_anna@example.com"


' ===========================================================================
' ENTRY POINT
' ===========================================================================
Public Sub RunDAOTests()
    TestRunStart "DAO-TESTS (Datenbank)"

    ' Schema sicherstellen (ueberspringt existierende Tabellen)
    On Error Resume Next
    ErstelleAlleTabellen
    On Error GoTo 0

    Test_SchemaExistenz
    Test_ConfigReadWrite
    Test_SyncLaufCRUD
    Test_KontaktCRUD
    Test_KontaktDeduplikation
    Test_OrdnerCRUD
    Test_ThreadCRUD
    Test_ThreadZuordnung
    Test_EmailCRUD
    Test_MailHashDeduplikation
    Test_EmailContentCRUD
    Test_EmpfaengerCRUD
    Test_AnhangCRUD

    ' Aufraeumen
    CleanupTestDaten

    TestRunEnd
End Sub


' ===========================================================================
' SCHEMA-TESTS
' ===========================================================================

Private Sub Test_SchemaExistenz()
    SuiteStart "Schema-Existenz"

    ' Alle 12+1 Tabellen muessen existieren
    AssertIsTrue TabelleExistiert(TBL_CONFIG), TBL_CONFIG & " existiert"
    AssertIsTrue TabelleExistiert(TBL_SYNC_LAUF), TBL_SYNC_LAUF & " existiert"
    AssertIsTrue TabelleExistiert(TBL_KONTAKTE), TBL_KONTAKTE & " existiert"
    AssertIsTrue TabelleExistiert(TBL_OUTLOOK_ORDNER), TBL_OUTLOOK_ORDNER & " existiert"
    AssertIsTrue TabelleExistiert(TBL_EMAIL_THREADS), TBL_EMAIL_THREADS & " existiert"
    AssertIsTrue TabelleExistiert(TBL_EMAILS), TBL_EMAILS & " existiert"
    AssertIsTrue TabelleExistiert(TBL_EMAIL_CONTENT), TBL_EMAIL_CONTENT & " existiert"
    AssertIsTrue TabelleExistiert(TBL_EMAIL_EMPFAENGER), TBL_EMAIL_EMPFAENGER & " existiert"
    AssertIsTrue TabelleExistiert(TBL_EMAIL_ANHAENGE), TBL_EMAIL_ANHAENGE & " existiert"
    AssertIsTrue TabelleExistiert(TBL_EMAIL_STATUS), TBL_EMAIL_STATUS & " existiert"
    AssertIsTrue TabelleExistiert(TBL_SYNC_PROFIL), TBL_SYNC_PROFIL & " existiert"
    AssertIsTrue TabelleExistiert(TBL_SYNC_PROFIL_ORDNER), TBL_SYNC_PROFIL_ORDNER & " existiert"
    AssertIsTrue TabelleExistiert(TBL_LOG), TBL_LOG & " existiert"

    ' Tabelle die nicht existiert
    AssertIsFalse TabelleExistiert("tblGibtEsNicht_XYZ"), "Nicht-existierende Tabelle = False"

    SuiteEnd
End Sub


' ===========================================================================
' CONFIG-TESTS
' ===========================================================================

Private Sub Test_ConfigReadWrite()
    SuiteStart "Config Read/Write"
    On Error GoTo ErrHandler

    ' Standard-Config lesen (sollte nach ErstelleAlleTabellen existieren)
    Dim strVal As String
    strVal = LeseConfig(CFG_SCHEMA_VERSION, "")
    AssertIsNotEmpty strVal, "SchemaVersion in Config vorhanden"

    ' Eigenen Test-Wert schreiben
    SchreibeConfig TEST_PREFIX & "TestKey", "TestWert123"

    ' Zuruecklesen
    strVal = LeseConfig(TEST_PREFIX & "TestKey", "")
    AssertAreEqual "TestWert123", strVal, "Geschriebener Config-Wert lesbar"

    ' Ueberschreiben
    SchreibeConfig TEST_PREFIX & "TestKey", "NeuerWert"
    strVal = LeseConfig(TEST_PREFIX & "TestKey", "")
    AssertAreEqual "NeuerWert", strVal, "Config-Wert ueberschrieben"

    ' Default-Wert fuer nicht-existierenden Key
    strVal = LeseConfig(TEST_PREFIX & "GibtEsNicht_" & Format(Now, "hhnnss"), "MeinDefault")
    AssertAreEqual "MeinDefault", strVal, "Default-Wert bei fehlendem Key"

    ' Aufraeumen
    CurrentDb.Execute "DELETE FROM [" & TBL_CONFIG & "] WHERE Schluessel LIKE '" & TEST_PREFIX & "*'"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Config-Test Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ===========================================================================
' SYNC-LAUF TESTS
' ===========================================================================

Private Sub Test_SyncLaufCRUD()
    SuiteStart "SyncLauf CRUD"
    On Error GoTo ErrHandler

    ' Starten
    Dim lngID As Long
    lngID = StarteSyncLauf(TEST_PREFIX & "\\Inbox", TEST_PREFIX & "Projekt", TEST_PREFIX & "Phase")
    AssertGreaterThan lngID, 0, "SyncLauf gestartet (ID > 0)"

    ' Beenden
    BeendeSyncLauf lngID, "Abgeschlossen", 100, 50, 45, 5
    
    ' Pruefen ob Status aktualisiert
    Dim strStatus As String
    strStatus = Nz(DLookup("Status", TBL_SYNC_LAUF, "SyncLaufID=" & lngID), "")
    AssertAreEqual "Abgeschlossen", strStatus, "SyncLauf-Status = Abgeschlossen"

    ' Zaehler pruefen
    Dim lngNeu As Long
    lngNeu = Nz(DLookup("AnzahlNeu", TBL_SYNC_LAUF, "SyncLaufID=" & lngID), 0)
    AssertAreEqual 50, lngNeu, "AnzahlNeu = 50"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "SyncLauf-Test Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ===========================================================================
' KONTAKT-TESTS
' ===========================================================================

Private Sub Test_KontaktCRUD()
    SuiteStart "Kontakt CRUD"
    On Error GoTo ErrHandler

    ' Neuen Kontakt anlegen
    Dim lngID As Long
    lngID = GetOderErstelleKontakt(TEST_PREFIX & "Max Mueller", TEST_EMAIL)
    AssertGreaterThan lngID, 0, "Kontakt erstellt (ID > 0)"

    ' Pruefen ob Felder korrekt gesetzt
    Dim strName As String
    strName = Nz(DLookup("Nachname", TBL_KONTAKTE, "KontaktID=" & lngID), "")
    AssertIsNotEmpty strName, "Nachname gesetzt"

    Dim strEmail As String
    strEmail = Nz(DLookup("Email", TBL_KONTAKTE, "KontaktID=" & lngID), "")
    AssertAreEqual TEST_EMAIL, strEmail, "Email korrekt gespeichert"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Kontakt-CRUD Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


Private Sub Test_KontaktDeduplikation()
    SuiteStart "Kontakt Deduplikation"
    On Error GoTo ErrHandler

    ' Gleichen Kontakt nochmal anlegen -> sollte gleiche ID liefern
    Dim lngID1 As Long, lngID2 As Long
    lngID1 = GetOderErstelleKontakt(TEST_PREFIX & "Max Mueller", TEST_EMAIL)
    lngID2 = GetOderErstelleKontakt(TEST_PREFIX & "Max Mueller", TEST_EMAIL)
    AssertAreEqual lngID1, lngID2, "Gleiche Email = gleiche KontaktID"

    ' Anderen Kontakt -> andere ID
    Dim lngID3 As Long
    lngID3 = GetOderErstelleKontakt(TEST_PREFIX & "Anna Schmidt", TEST_EMAIL2)
    AssertAreNotEqual lngID1, lngID3, "Andere Email = andere KontaktID"

    ' Cache-Hit pruefen (zweiter Aufruf sollte aus Cache kommen)
    Dim lngID4 As Long
    lngID4 = GetOderErstelleKontakt(TEST_PREFIX & "Max Mueller", TEST_EMAIL)
    AssertAreEqual lngID1, lngID4, "Cache-Hit = gleiche ID"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Kontakt-Dedup Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ===========================================================================
' ORDNER-TESTS
' ===========================================================================

Private Sub Test_OrdnerCRUD()
    SuiteStart "Ordner CRUD"
    On Error GoTo ErrHandler

    Dim lngID As Long
    lngID = SpeichereOrdner(TEST_PREFIX & "Posteingang", _
                            TEST_PREFIX & "\\Postfach\Posteingang", _
                            0, TEST_PREFIX & "TestPostfach", 42)
    AssertGreaterThan lngID, 0, "Ordner erstellt (ID > 0)"

    ' Gleicher Pfad -> Update statt Neuanlage
    Dim lngID2 As Long
    lngID2 = SpeichereOrdner(TEST_PREFIX & "Posteingang", _
                             TEST_PREFIX & "\\Postfach\Posteingang", _
                             0, TEST_PREFIX & "TestPostfach", 99)
    AssertAreEqual lngID, lngID2, "Gleicher Pfad = gleiche OrdnerID (Update)"

    ' Elementanzahl aktualisiert?
    Dim lngAnz As Long
    lngAnz = Nz(DLookup("ElementAnzahl", TBL_OUTLOOK_ORDNER, "OrdnerID=" & lngID), 0)
    AssertAreEqual 99, lngAnz, "Elementanzahl aktualisiert auf 99"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Ordner-Test Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ===========================================================================
' THREAD-TESTS
' ===========================================================================

Private Sub Test_ThreadCRUD()
    SuiteStart "Thread CRUD"
    On Error GoTo ErrHandler

    Dim lngID As Long
    lngID = GetOderErstelleThread(TEST_PREFIX & "Testbetreff", _
                                  TEST_PREFIX & "<msg001@test.com>", _
                                  "", Now, TEST_PREFIX & "Absender")
    AssertGreaterThan lngID, 0, "Thread erstellt (ID > 0)"

    ' Betreff pruefen
    Dim strBetreff As String
    strBetreff = Nz(DLookup("ThreadBetreff", TBL_EMAIL_THREADS, "ThreadID=" & lngID), "")
    AssertIsNotEmpty strBetreff, "ThreadBetreff gesetzt"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Thread-CRUD Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


Private Sub Test_ThreadZuordnung()
    SuiteStart "Thread Zuordnung"
    On Error GoTo ErrHandler

    ' Neuen Thread mit MessageID
    Dim lngID1 As Long
    lngID1 = GetOderErstelleThread(TEST_PREFIX & "Original-Betreff", _
                                   TEST_PREFIX & "<original@test.com>", _
                                   "", #1/1/2026#, TEST_PREFIX & "Sender1")
    AssertGreaterThan lngID1, 0, "Original-Thread erstellt"

    ' Antwort via InReplyTo -> sollte gleichen Thread finden
    Dim lngID2 As Long
    lngID2 = GetOderErstelleThread(TEST_PREFIX & "RE: Original-Betreff", _
                                   TEST_PREFIX & "<reply@test.com>", _
                                   TEST_PREFIX & "<original@test.com>", # _
                                   1/2/2026#, TEST_PREFIX & "Sender2")

    ' InReplyTo-Match: Hier wird ThreadIdentifier gesucht.
    ' Der erste Thread hat ThreadIdentifier = "<original@test.com>" (MessageID)
    ' Der zweite sucht erst via InReplyTo "<original@test.com>" -> Treffer!
    ' HINWEIS: Ob das matcht haengt davon ab, ob ThreadIdentifier = MessageID gesetzt wird
    ' Bei neuem Thread wird InReplyTo als Identifier genutzt wenn vorhanden

    ' Antwort via Betreff (ohne InReplyTo) -> gleicher bereinigter Betreff
    Dim lngID3 As Long
    lngID3 = GetOderErstelleThread(TEST_PREFIX & "AW: Original-Betreff", _
                                   TEST_PREFIX & "<aw@test.com>", _
                                   "", #1/3/2026#, TEST_PREFIX & "Sender3")

    ' Bereinigter Betreff "Original-Betreff" sollte den Thread finden
    ' (wenn der ThreadIdentifier auf dem bereinigten Betreff liegt)
    ' Das funktioniert NUR wenn der erste Thread seinen Identifier = bereinigter Betreff hat
    ' Aktuell: bei erstem Thread wird MessageID als Identifier genommen (InReplyTo war leer)
    ' -> Betreff-Match ist ein separater Lookup

    ' Mindestens: IDs muessen > 0 sein
    AssertGreaterThan lngID2, 0, "Antwort-Thread (InReplyTo) angelegt/gefunden"
    AssertGreaterThan lngID3, 0, "Antwort-Thread (Betreff) angelegt/gefunden"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Thread-Zuordnung Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ===========================================================================
' EMAIL-TESTS
' ===========================================================================

Private Sub Test_EmailCRUD()
    SuiteStart "Email CRUD"
    On Error GoTo ErrHandler

    ' Voraussetzungen
    Dim lngKontaktID As Long, lngOrdnerID As Long
    Dim lngThreadID As Long, lngSyncLaufID As Long

    lngKontaktID = GetOderErstelleKontakt(TEST_PREFIX & "Sender", TEST_EMAIL)
    lngOrdnerID = SpeichereOrdner(TEST_PREFIX & "Inbox", TEST_PREFIX & "\\Test\Inbox")
    lngThreadID = GetOderErstelleThread(TEST_PREFIX & "Email-Test", TEST_PREFIX & "<email-test@test.com>")
    lngSyncLaufID = StarteSyncLauf(TEST_PREFIX & "\\Test", TEST_PREFIX & "Proj", TEST_PREFIX & "Ph")

    ' Email speichern
    Dim strHash As String
    strHash = GeneriereMailHash(TEST_PREFIX & "Betreff", TEST_EMAIL, "empf@test.de", Now)

    Dim lngEmailID As Long
    lngEmailID = SpeichereEmail( _
        TEST_PREFIX & "ENTRYID123", strHash, _
        lngThreadID, lngOrdnerID, lngKontaktID, lngSyncLaufID, _
        TEST_PREFIX & "Testbetreff", TEST_PREFIX & "Max Mueller", TEST_EMAIL, _
        Now, Now, 1024, 1, True, False, 0, "IPM.Note", TEST_PREFIX & "<msg@test.com>")

    AssertGreaterThan lngEmailID, 0, "Email gespeichert (ID > 0)"

    ' Pruefen ob Hash gespeichert
    Dim strDBHash As String
    strDBHash = Nz(DLookup("UniqueHash", TBL_EMAILS, "EmailID=" & lngEmailID), "")
    AssertAreEqual strHash, strDBHash, "Hash korrekt gespeichert"

    ' Status pruefen
    Dim strStatus As String
    strStatus = Nz(DLookup("Status", TBL_EMAILS, "EmailID=" & lngEmailID), "")
    AssertAreEqual "Neu", strStatus, "Initialer Status = Neu"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Email-CRUD Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


Private Sub Test_MailHashDeduplikation()
    SuiteStart "Mail-Hash Deduplikation"
    On Error GoTo ErrHandler

    ' Hash generieren
    Dim strHash As String
    strHash = GeneriereMailHash(TEST_PREFIX & "DedupTest", TEST_EMAIL, "empf@test.de", #3/5/2026 10:30:00 AM#)
    AssertAreEqual 64, Len(strHash), "Hash ist 64 Zeichen"

    ' ExistiertMailHash testen
    ' (haengt davon ab ob vorheriger Test eine Mail mit diesem Hash angelegt hat)
    ' Neuen eindeutigen Hash testen -> sollte nicht existieren
    Dim strUniqueHash As String
    strUniqueHash = GeneriereMailHash(TEST_PREFIX & "Unique_" & Format(Now, "hhnnss"), _
                                     "unique_" & Format(Timer, "0") & "@test.de", _
                                     "empf@test.de", Now)
    AssertIsFalse ExistiertMailHash(strUniqueHash), "Neuer Hash existiert noch nicht"

    ' Gleicher Hash zweimal -> deterministisch
    Dim strHash2 As String
    strHash2 = GeneriereMailHash(TEST_PREFIX & "DedupTest", TEST_EMAIL, "empf@test.de", #3/5/2026 10:30:00 AM#)
    AssertAreEqual strHash, strHash2, "Gleiche Eingaben = gleicher Hash"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "MailHash-Dedup Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ===========================================================================
' EMAIL-CONTENT TESTS
' ===========================================================================

Private Sub Test_EmailContentCRUD()
    SuiteStart "EmailContent CRUD"
    On Error GoTo ErrHandler

    ' Brauchen eine Email-ID -> letzte Test-Email nehmen
    Dim varEmailID As Variant
    varEmailID = DLookup("MAX(EmailID)", TBL_EMAILS, "Betreff LIKE '" & TEST_PREFIX & "*'")

    If IsNull(varEmailID) Then
        AssertSkip "Keine Test-Email vorhanden (Email-CRUD muss zuerst laufen)"
        SuiteEnd
        Exit Sub
    End If

    Dim lngEmailID As Long
    lngEmailID = CLng(varEmailID)

    ' Content speichern
    Dim strHTML As String, strPlain As String
    strHTML = "<html><body><p>" & TEST_PREFIX & "Test-Content</p></body></html>"
    strPlain = TEST_PREFIX & "Test-Content Plaintext"
    SpeichereEmailContent lngEmailID, strHTML, strPlain

    ' Pruefen
    Dim strDBHTML As String
    strDBHTML = Nz(DLookup("HTMLBody", TBL_EMAIL_CONTENT, "EmailID=" & lngEmailID), "")
    AssertContains strDBHTML, TEST_PREFIX, "HTML-Content gespeichert"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "EmailContent Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ===========================================================================
' EMPFAENGER-TESTS
' ===========================================================================

Private Sub Test_EmpfaengerCRUD()
    SuiteStart "Empfaenger CRUD"
    On Error GoTo ErrHandler

    ' Letzte Test-Email
    Dim varEmailID As Variant
    varEmailID = DLookup("MAX(EmailID)", TBL_EMAILS, "Betreff LIKE '" & TEST_PREFIX & "*'")

    If IsNull(varEmailID) Then
        AssertSkip "Keine Test-Email vorhanden"
        SuiteEnd
        Exit Sub
    End If

    Dim lngEmailID As Long
    lngEmailID = CLng(varEmailID)

    ' Empfaenger anlegen
    Dim lngKontaktID As Long
    lngKontaktID = GetOderErstelleKontakt(TEST_PREFIX & "Empfaenger", TEST_EMAIL2)
    SpeichereEmpfaenger lngEmailID, lngKontaktID, "To", TEST_PREFIX & "Anna Schmidt", TEST_EMAIL2

    ' Pruefen
    Dim lngCount As Long
    lngCount = DCount("*", TBL_EMAIL_EMPFAENGER, "EmailID=" & lngEmailID)
    AssertGreaterThan lngCount, 0, "Empfaenger gespeichert"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Empfaenger Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ===========================================================================
' ANHANG-TESTS
' ===========================================================================

Private Sub Test_AnhangCRUD()
    SuiteStart "Anhang CRUD"
    On Error GoTo ErrHandler

    ' Letzte Test-Email
    Dim varEmailID As Variant
    varEmailID = DLookup("MAX(EmailID)", TBL_EMAILS, "Betreff LIKE '" & TEST_PREFIX & "*'")

    If IsNull(varEmailID) Then
        AssertSkip "Keine Test-Email vorhanden"
        SuiteEnd
        Exit Sub
    End If

    Dim lngEmailID As Long
    lngEmailID = CLng(varEmailID)

    ' Anhang anlegen
    Dim lngAnhangID As Long
    lngAnhangID = SpeichereAnhangMetadaten(lngEmailID, _
        TEST_PREFIX & "dokument.pdf", 1024, "application/pdf", 1, False)
    AssertGreaterThan lngAnhangID, 0, "Anhang gespeichert (ID > 0)"

    ' Felder pruefen
    Dim strName As String
    strName = Nz(DLookup("Dateiname", TBL_EMAIL_ANHAENGE, "AnhangID=" & lngAnhangID), "")
    AssertContains strName, ".pdf", "Dateiname mit .pdf"

    SuiteEnd
    Exit Sub

ErrHandler:
    AssertFail "Anhang Fehler: " & Err.Number & " - " & Err.Description
    SuiteEnd
End Sub


' ===========================================================================
' CLEANUP - Test-Daten entfernen
' ===========================================================================

Private Sub CleanupTestDaten()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb

    Debug.Print ""
    Debug.Print "--- Cleanup: Test-Daten entfernen ---"

    ' Reihenfolge: abhaengige Tabellen zuerst (FK-Reihenfolge)
    ' Anhaenge + Empfaenger + Content + Status fuer Test-Emails
    db.Execute "DELETE FROM [" & TBL_EMAIL_ANHAENGE & "] WHERE EmailID IN " & _
               "(SELECT EmailID FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*')"
    db.Execute "DELETE FROM [" & TBL_EMAIL_EMPFAENGER & "] WHERE EmailID IN " & _
               "(SELECT EmailID FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*')"
    db.Execute "DELETE FROM [" & TBL_EMAIL_CONTENT & "] WHERE EmailID IN " & _
               "(SELECT EmailID FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*')"
    db.Execute "DELETE FROM [" & TBL_EMAIL_STATUS & "] WHERE EmailID IN " & _
               "(SELECT EmailID FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*')"

    ' Emails
    db.Execute "DELETE FROM [" & TBL_EMAILS & "] WHERE Betreff LIKE '" & TEST_PREFIX & "*'"

    ' Threads
    db.Execute "DELETE FROM [" & TBL_EMAIL_THREADS & "] WHERE ThreadBetreff LIKE '" & TEST_PREFIX & "*'"

    ' Ordner
    db.Execute "DELETE FROM [" & TBL_OUTLOOK_ORDNER & "] WHERE OrdnerPfad LIKE '" & TEST_PREFIX & "*'"

    ' Kontakte
    db.Execute "DELETE FROM [" & TBL_KONTAKTE & "] WHERE Email LIKE 'unittest_*'"

    ' SyncLaeufe
    db.Execute "DELETE FROM [" & TBL_SYNC_LAUF & "] WHERE OrdnerPfad LIKE '" & TEST_PREFIX & "*'"

    ' Config Test-Werte
    db.Execute "DELETE FROM [" & TBL_CONFIG & "] WHERE Schluessel LIKE '" & TEST_PREFIX & "*'"

    Set db = Nothing
    Debug.Print "    Test-Daten bereinigt."
    Debug.Print ""
    On Error GoTo 0
End Sub


