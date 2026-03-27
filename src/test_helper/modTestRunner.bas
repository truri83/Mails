Option Compare Database
Option Explicit

' ===========================================================================
' modTestRunner - Zentraler Test-Runner fuer OutlookSync
' ===========================================================================
' Fuehrt alle Test-Suites aus und gibt eine Gesamtzusammenfassung.
'
' AUFRUF IM DIREKTBEREICH (Strg+G):
'
'   RunAlleTests               ' ALLES (Unit + DAO + Integration)
'   RunUnitTests               ' Nur reine Funktionen (kein Outlook noetig)
'   RunDAOTests                ' Nur DB-Tests (kein Outlook noetig)
'   RunIntegrationTests        ' Outlook + Redemption noetig
'   RunOfflineTests            ' Unit + DAO (alles ohne Outlook)
'
' VORAUSSETZUNGEN:
'   - ErstelleAlleTabellen muss einmalig gelaufen sein
'   - Fuer Integration: Outlook offen + Redemption registriert
'
' TEST-MODULE:
'   modTestFramework    - Assert-Funktionen, Suite-Verwaltung, Reporting
'   modTestUnit         - String, Crypto, Kontakte, Cache (~50 Tests)
'   modTestDAO          - Schema, Config, CRUD, Dedup, Threads (~30 Tests)
'   modTestIntegration  - Outlook, Extraktion, Buffer-Flush (~20 Tests)
'
' BESTEHENDE TEST-HELPER (unabhaengig, separat aufrufbar):
'   modOutlookTest          - Deep Outlook/RDO Access Tests
'   modTestKopiermethoden   - Datei-Kopiermethoden Benchmark
'   modTestPerformance      - DB/Outlook/Hash Performance
' ===========================================================================


' ===========================================================================
' HAUPTROUTINE: Alle Tests
' ===========================================================================
Public Sub RunAlleTests()
    Debug.Print ""
    Debug.Print String(70, "*")
    Debug.Print "  OutlookSync Test-Suite"
    Debug.Print "  " & Format(Now, "dd.mm.yyyy hh:nn:ss")
    Debug.Print String(70, "*")
    Debug.Print ""

    ' Phase 1: Unit-Tests (keine externen Abhaengigkeiten)
    Debug.Print ">>> PHASE 1: Unit-Tests (reine Funktionen)"
    Debug.Print ""
    RunUnitTests

    ' Phase 2: DAO-Tests (braucht Schema)
    Debug.Print ""
    Debug.Print ">>> PHASE 2: DAO-Tests (Datenbank)"
    Debug.Print ""
    RunDAOTests

    ' Phase 3: Integration (braucht Outlook + Redemption)
    Debug.Print ""
    Debug.Print ">>> PHASE 3: Integration-Tests (Outlook + DB)"
    Debug.Print ""
    RunIntegrationTests

    Debug.Print ""
    Debug.Print String(70, "*")
    Debug.Print "  Test-Suite abgeschlossen"
    Debug.Print String(70, "*")
End Sub


' ===========================================================================
' OFFLINE-TESTS (Unit + DAO, ohne Outlook)
' ===========================================================================
Public Sub RunOfflineTests()
    Debug.Print ""
    Debug.Print String(70, "*")
    Debug.Print "  OutlookSync OFFLINE Test-Suite (kein Outlook noetig)"
    Debug.Print "  " & Format(Now, "dd.mm.yyyy hh:nn:ss")
    Debug.Print String(70, "*")
    Debug.Print ""

    Debug.Print ">>> PHASE 1: Unit-Tests"
    Debug.Print ""
    RunUnitTests

    Debug.Print ""
    Debug.Print ">>> PHASE 2: DAO-Tests"
    Debug.Print ""
    RunDAOTests

    Debug.Print ""
    Debug.Print String(70, "*")
    Debug.Print "  Offline-Tests abgeschlossen"
    Debug.Print String(70, "*")
End Sub

