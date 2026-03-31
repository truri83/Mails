
Option Compare Database
Option Explicit

' ===========================================================================
' modTestUnit - Unit-Tests fuer reine Funktionen (KEIN Outlook noetig)
' ===========================================================================
' Testet alle Funktionen die ohne externe Abhaengigkeiten laufen:
'   - modStringUtils  (Email-Validierung, Bereinigung, Pfade)
'   - modCrypto       (SHA256, MailHash)
'   - modKontakte     (Name-Parsing, Geschlecht, Anrede)
'   - modCache        (Init, Get/Set, Reset, Statistik)
'
' AUFRUF:
'   RunUnitTests       ' Alle Unit-Tests
'
' Voraussetzungen: Nur Access (Schema muss fuer Cache-Tests existieren)
' ===========================================================================

Private Const MODUL_NAME As String = "modTestUnit"


' ===========================================================================
' ENTRY POINT
' ===========================================================================
Public Sub RunUnitTests()
    TestRunStart "UNIT-TESTS"

    Test_IstGueltigeEmail
    Test_BereinigeBetreff
    Test_BereinigeDateiname
    Test_SQLSafe
    Test_BereinigeAnhangName
    Test_GeneriereMailDateiname
    Test_NormalisierePfad
    Test_HoleEndung

    Test_SHA256_Hash
    Test_GeneriereMailHash

    Test_ParseKontaktName
    Test_FallbackNameAusEmail
    Test_IstSystemEmail
    Test_ErkenneGeschlecht
    Test_BildeAnrede
    Test_IstKontaktPlausibel

    Test_CacheInitReset
    Test_CacheConfigGetSet
    Test_CacheKontaktGetSet
    Test_CacheTabellenCache

    TestRunEnd
End Sub


' ===========================================================================
' modStringUtils TESTS
' ===========================================================================

Private Sub Test_IstGueltigeEmail()
    SuiteStart "IstGueltigeEmail"

    ' Gueltige Adressen
    AssertIsTrue IstGueltigeEmail("max@example.com"), "Standard-Email"
    AssertIsTrue IstGueltigeEmail("vorname.nachname@firma.de"), "Punkt im User"
    AssertIsTrue IstGueltigeEmail("user+tag@example.co.uk"), "Plus-Adresse + Subdomain"
    AssertIsTrue IstGueltigeEmail("a@b.cc"), "Minimale gueltige Adresse"

    ' Ungueltige Adressen
    AssertIsFalse IstGueltigeEmail(""), "Leerstring"
    AssertIsFalse IstGueltigeEmail("kein-at-zeichen"), "Kein @"
    AssertIsFalse IstGueltigeEmail("@domain.de"), "Kein User"
    AssertIsFalse IstGueltigeEmail("user@"), "Keine Domain"
    AssertIsFalse IstGueltigeEmail("user@domain"), "Keine TLD"
    AssertIsFalse IstGueltigeEmail("/O=EXCHANGELABS/OU=..."), "Exchange X500-Pfad"

    SuiteEnd
End Sub


Private Sub Test_BereinigeBetreff()
    SuiteStart "BereinigeBetreff"

    AssertAreEqual "Test", BereinigeBetreff("RE: Test"), "RE: entfernen"
    AssertAreEqual "Test", BereinigeBetreff("AW: Test"), "AW: entfernen"
    AssertAreEqual "Test", BereinigeBetreff("FW: Test"), "FW: entfernen"
    AssertAreEqual "Test", BereinigeBetreff("WG: Test"), "WG: entfernen"
    AssertAreEqual "Test", BereinigeBetreff("RE: RE: AW: Test"), "Mehrfach RE/AW"
    AssertAreEqual "Test", BereinigeBetreff("EXTERN RE: Test"), "EXTERN entfernen"
    AssertAreEqual "(Kein Betreff)", BereinigeBetreff(""), "Leerstring -> Fallback"
    AssertAreEqual "Hallo Welt", BereinigeBetreff("Hallo Welt"), "Ohne Prefix unveraendert"
    AssertAreEqual "Test", BereinigeBetreff("FWD: Test"), "FWD: entfernen"

    SuiteEnd
End Sub


Private Sub Test_BereinigeDateiname()
    SuiteStart "BereinigeDateiname"

    AssertAreEqual "Test", BereinigeDateiname("Test"), "Normaler Name"
    AssertAreEqual "A-B-C", BereinigeDateiname("A/B\C"), "Slashes werden Bindestriche"
    AssertAreEqual "Test", BereinigeDateiname("Test*?<>|"), "Sonderzeichen entfernt"
    AssertAreEqual "A-B", BereinigeDateiname("A:B"), "Doppelpunkt wird Bindestrich"

    ' Laengenbegrenzung
    Dim strLang As String
    strLang = String(200, "A")
    AssertIsTrue Len(BereinigeDateiname(strLang, 50)) <= 50, "Laengenlimit 50 Zeichen"

    SuiteEnd
End Sub


Private Sub Test_SQLSafe()
    SuiteStart "SQLSafe"

    AssertAreEqual "Test", SQLSafe("Test"), "Normaler Text"
    AssertAreEqual "O''Brien", SQLSafe("O'Brien"), "Apostroph verdoppelt"
    AssertAreEqual "Hans''s Test", SQLSafe("Hans's Test"), "Mittiger Apostroph"
    AssertAreEqual "", SQLSafe(""), "Leerstring"

    SuiteEnd
End Sub


Private Sub Test_BereinigeAnhangName()
    SuiteStart "BereinigeAnhangName"

    AssertAreEqual "Test.pdf", BereinigeAnhangName("Test.pdf"), "Normaler Anhangname"
    AssertIsNotEmpty BereinigeAnhangName(""), "Leerstring bekommt Fallback-Namen"

    ' Extension bleibt erhalten
    Dim strResult As String
    strResult = BereinigeAnhangName("Mein*Dokument.xlsx")
    AssertContains strResult, ".xlsx", "Extension bleibt erhalten"

    SuiteEnd
End Sub


Private Sub Test_GeneriereMailDateiname()
    SuiteStart "GeneriereMailDateiname"

    Dim strResult As String
    strResult = GeneriereMailDateiname("Max Mueller", "Testbetreff", #1/15/2026#, "eml")

    AssertContains strResult, "20260115", "Datum im Dateinamen"
    AssertContains strResult, "Mueller", "Absender im Dateinamen"
    AssertContains strResult, "Testbetreff", "Betreff im Dateinamen"

    ' MSG-Endung
    strResult = GeneriereMailDateiname("Test", "Betr", #3/5/2026#, "msg")
    AssertContains strResult, ".msg", "MSG-Endung"

    SuiteEnd
End Sub


Private Sub Test_NormalisierePfad()
    SuiteStart "NormalisierePfad"

    ' Trailing Backslash
    AssertAreEqual "C:\Test\", NormalisierePfad("C:\Test"), "Trailing Backslash hinzufuegen"
    AssertAreEqual "C:\Test\", NormalisierePfad("C:\Test\"), "Trailing Backslash bleibt"

    ' UNC-Pfade
    Dim strUNC As String
    strUNC = NormalisierePfad("\\Server\Share\Ordner")
    AssertContains strUNC, "\\", "UNC-Prefix erhalten"

    SuiteEnd
End Sub


Private Sub Test_HoleEndung()
    SuiteStart "HoleEndung"

    AssertAreEqual "pdf", HoleEndung("Test.pdf"), "PDF-Extension"
    AssertAreEqual "xlsx", HoleEndung("Dokument.xlsx"), "XLSX-Extension"
    AssertAreEqual "pdf", HoleEndung("TEST.PDF"), "Lowercase-Konvertierung"
    AssertAreEqual "", HoleEndung("Kein_Punkt"), "Keine Extension"

    SuiteEnd
End Sub


' ===========================================================================
' modCrypto TESTS
' ===========================================================================

Private Sub Test_SHA256_Hash()
    SuiteStart "SHA256_Hash"

    ' Bekannte Referenzwerte (verifiziert mit externen Tools)
    Dim strHash As String

    ' SHA256("") - leerer String
    strHash = SHA256_Hash("")
    AssertAreEqual "", strHash, "Leerstring -> Leerstring (by design)"

    ' SHA256("abc") = ba7816bf...
    strHash = SHA256_Hash("abc")
    AssertAreEqual 64, Len(strHash), "Hash-Laenge = 64 Zeichen"
    AssertAreEqual "ba7816bf", Left(strHash, 8), "SHA256('abc') beginnt mit ba7816bf"

    ' Determinismus: gleicher Input = gleicher Output
    AssertAreEqual SHA256_Hash("Hello World"), SHA256_Hash("Hello World"), "Deterministisch"

    ' Verschiedene Inputs = verschiedene Hashes
    AssertAreNotEqual SHA256_Hash("Test1"), SHA256_Hash("Test2"), "Verschiedene Inputs"

    SuiteEnd
End Sub


Private Sub Test_GeneriereMailHash()
    SuiteStart "GeneriereMailHash"

    Dim strHash As String
    strHash = GeneriereMailHash("Betreff", "absender@test.de", "empf@test.de", #3/5/2026#)

    AssertAreEqual 64, Len(strHash), "Mail-Hash ist 64 Zeichen"
    AssertIsNotEmpty strHash, "Mail-Hash nicht leer"

    ' Gleiche Eingaben = gleicher Hash (Deduplikation)
    Dim strHash2 As String
    strHash2 = GeneriereMailHash("Betreff", "absender@test.de", "empf@test.de", #3/5/2026#)
    AssertAreEqual strHash, strHash2, "Gleiche Mail = gleicher Hash"

    ' Anderer Absender = anderer Hash
    strHash2 = GeneriereMailHash("Betreff", "anderer@test.de", "empf@test.de", #3/5/2026#)
    AssertAreNotEqual strHash, strHash2, "Anderer Absender = anderer Hash"

    ' Anderes Datum = anderer Hash
    strHash2 = GeneriereMailHash("Betreff", "absender@test.de", "empf@test.de", #3/6/2026#)
    AssertAreNotEqual strHash, strHash2, "Anderes Datum = anderer Hash"

    SuiteEnd
End Sub


' ===========================================================================
' modKontakte TESTS
' ===========================================================================

Private Sub Test_ParseKontaktName()
    SuiteStart "ParseKontaktName"

    Dim kn As TypKontaktName

    ' Standard: "Vorname Nachname"
    kn = ParseKontaktName("Max Mueller")
    AssertAreEqual "Max", kn.Vorname, "Vorname aus 'Max Mueller'"
    AssertAreEqual "Mueller", kn.Nachname, "Nachname aus 'Max Mueller'"

    ' Mit Titel: "Dr. Anna Schmidt"
    kn = ParseKontaktName("Dr. Anna Schmidt")
    AssertAreEqual "Dr.", kn.Titel, "Titel erkannt"
    AssertAreEqual "Anna", kn.Vorname, "Vorname mit Titel"
    AssertAreEqual "Schmidt", kn.Nachname, "Nachname mit Titel"

    ' Nachname, Vorname Format
    kn = ParseKontaktName("Mueller, Max")
    AssertAreEqual "Max", kn.Vorname, "Vorname Komma-Format"
    AssertAreEqual "Mueller", kn.Nachname, "Nachname Komma-Format"

    ' Mit Klammer-Institution: "Jan Meier (LUBW)"
    kn = ParseKontaktName("Jan Meier (LUBW)")
    AssertAreEqual "Jan", kn.Vorname, "Vorname mit Institution"
    AssertAreEqual "Meier", kn.Nachname, "Nachname mit Institution"
    AssertIsNotEmpty kn.Institution, "Institution aus Klammer"

    ' Prof. Dr. mit Adelspraedikat
    kn = ParseKontaktName("Prof. Dr. Hans von Hohenheim")
    AssertAreEqual "Prof. Dr.", kn.Titel, "Prof. Dr. erkannt"
    AssertAreEqual "Hans", kn.Vorname, "Vorname bei 3+ Teilen"
    AssertAreEqual "Hohenheim", kn.Nachname, "Nachname bei 3+ Teilen"
    AssertIsNotEmpty kn.Namenszusatz, "Namenszusatz 'von' erkannt"

    ' Nur ein Name
    kn = ParseKontaktName("Mueller")
    AssertAreEqual "Mueller", kn.Nachname, "Einzelwort = Nachname"
    AssertAreEqual "", kn.Vorname, "Kein Vorname bei Einzelwort"

    ' Sortiername
    kn = ParseKontaktName("Max Mueller")
    AssertIsNotEmpty kn.Sortiername, "Sortiername gebildet"

    SuiteEnd
End Sub


Private Sub Test_FallbackNameAusEmail()
    SuiteStart "FallbackNameAusEmail"

    Dim kn As TypKontaktName

    ' vorname.nachname@domain.de
    kn = FallbackNameAusEmail("max.mueller@firma.de")
    AssertAreEqual "Max", kn.Vorname, "Vorname aus Email"
    AssertAreEqual "Mueller", kn.Nachname, "Nachname aus Email"

    ' info@domain.de -> Institution
    kn = FallbackNameAusEmail("info@lubw.bwl.de")
    AssertIsNotEmpty kn.Anzeigename, "Anzeigename fuer System-Email"

    ' Kein @ -> Default
    kn = FallbackNameAusEmail("kein-at-zeichen")
    AssertAreEqual DEFAULT_NAME, kn.Nachname, "Fallback bei fehlerhafter Email"

    SuiteEnd
End Sub


Private Sub Test_IstSystemEmail()
    SuiteStart "IstSystemEmail"

    AssertIsTrue IstSystemEmail("info@firma.de"), "info@ ist System"
    AssertIsTrue IstSystemEmail("noreply@example.com"), "noreply@ ist System"
    AssertIsTrue IstSystemEmail("support@test.de"), "support@ ist System"
    AssertIsTrue IstSystemEmail("newsletter@news.de"), "newsletter@ ist System"

    AssertIsFalse IstSystemEmail("max.mueller@firma.de"), "Persoenlich = kein System"
    AssertIsFalse IstSystemEmail("stefan@example.com"), "Vorname = kein System"
    AssertIsFalse IstSystemEmail(""), "Leerstring = kein System"

    SuiteEnd
End Sub


Private Sub Test_ErkenneGeschlecht()
    SuiteStart "ErkenneGeschlecht"

    AssertAreEqual "m", ErkenneGeschlecht("Thomas"), "Thomas = maennlich"
    AssertAreEqual "m", ErkenneGeschlecht("Michael"), "Michael = maennlich"
    AssertAreEqual "w", ErkenneGeschlecht("Anna"), "Anna = weiblich"
    AssertAreEqual "w", ErkenneGeschlecht("Sabine"), "Sabine = weiblich"
    AssertAreEqual "", ErkenneGeschlecht("Robin"), "Robin = unbekannt"
    AssertAreEqual "", ErkenneGeschlecht(""), "Leer = unbekannt"

    ' Case-insensitive
    AssertAreEqual "m", ErkenneGeschlecht("THOMAS"), "THOMAS uppercase"
    AssertAreEqual "w", ErkenneGeschlecht("julia"), "julia lowercase"

    SuiteEnd
End Sub


Private Sub Test_BildeAnrede()
    SuiteStart "BildeAnrede"

    AssertContains BildeAnrede("", "Thomas", "Mueller"), "Herr", "Maennliche Anrede"
    AssertContains BildeAnrede("", "Anna", "Schmidt"), "Frau", "Weibliche Anrede"
    AssertContains BildeAnrede("Dr.", "Thomas", "Mueller"), "Dr.", "Titel in Anrede"
    AssertContains BildeAnrede("", "", ""), "Damen und Herren", "Ohne Name generisch"

    SuiteEnd
End Sub


Private Sub Test_IstKontaktPlausibel()
    SuiteStart "IstKontaktPlausibel"

    AssertIsTrue IstKontaktPlausibel("Max", "Mueller"), "Normal plausibel"
    AssertIsTrue IstKontaktPlausibel("", "Mueller"), "Ohne Vorname plausibel"
    AssertIsFalse IstKontaktPlausibel("", "X"), "Zu kurzer Nachname"
    AssertIsFalse IstKontaktPlausibel("", ""), "Komplett leer"

    SuiteEnd
End Sub


' ===========================================================================
' modCache TESTS
' ===========================================================================

Private Sub Test_CacheInitReset()
    SuiteStart "Cache Init/Reset"

    ' Init
    CacheInit
    ' Kein Fehler = OK
    AssertIsTrue True, "CacheInit ohne Fehler"

    ' Reset
    CacheReset
    AssertIsTrue True, "CacheReset ohne Fehler"

    ' Erneuter Init nach Reset
    CacheInit
    AssertIsTrue True, "CacheInit nach Reset"

    SuiteEnd
End Sub


Private Sub Test_CacheConfigGetSet()
    SuiteStart "Cache Config Get/Set"

    CacheInit

    ' Default-Wert wenn Key nicht existiert
    Dim strVal As String
    strVal = CacheGetConfig("TEST_NICHT_EXISTIERT_" & Format(Now, "hhnnss"), "MeinDefault")
    AssertAreEqual "MeinDefault", strVal, "Default-Wert fuer unbekannten Key"

    ' Zweimal gleicher Key -> Cache-Hit
    Dim strVal2 As String
    strVal2 = CacheGetConfig("TEST_NICHT_EXISTIERT_" & Format(Now, "hhnnss"), "MeinDefault")
    AssertAreEqual strVal, strVal2, "Cache-Hit liefert gleichen Wert"

    ' Reset leert Cache
    CacheResetConfig
    AssertIsTrue True, "CacheResetConfig ohne Fehler"

    CacheReset
    SuiteEnd
End Sub


Private Sub Test_CacheKontaktGetSet()
    SuiteStart "Cache Kontakt Get/Set"

    CacheInit

    ' Unbekannter Kontakt -> 0
    Dim lngID As Long
    lngID = CacheGetKontaktID("test_gibts_nicht@example.com")
    AssertAreEqual 0, lngID, "Unbekannter Kontakt = 0"

    ' Kontakt cachen
    CacheSetKontaktID "test@example.com", 42
    lngID = CacheGetKontaktID("test@example.com")
    AssertAreEqual 42, lngID, "Gecachter Kontakt = 42"

    ' Case-insensitive
    lngID = CacheGetKontaktID("TEST@EXAMPLE.COM")
    AssertAreEqual 42, lngID, "Case-insensitive Lookup"

    ' Reset
    CacheResetKontakte
    lngID = CacheGetKontaktID("test@example.com")
    AssertAreEqual 0, lngID, "Nach Reset = 0"

    CacheReset
    SuiteEnd
End Sub


Private Sub Test_CacheTabellenCache()
    SuiteStart "Cache Tabellen"

    CacheInit

    ' Bekannte Tabelle testen (tblConfig sollte existieren wenn Schema steht)
    On Error Resume Next
    Dim blnExists As Boolean
    blnExists = CacheTabelleExistiert(TBL_CONFIG)
    If Err.Number <> 0 Then
        AssertSkip "tblConfig nicht verfuegbar (Schema nicht erstellt?)"
        Err.Clear
    Else
        ' Zweiter Aufruf = Cache-Hit (kein DB-Zugriff)
        Dim blnExists2 As Boolean
        blnExists2 = CacheTabelleExistiert(TBL_CONFIG)
        AssertAreEqual CStr(blnExists), CStr(blnExists2), "Tabellen-Cache konsistent"
    End If
    On Error GoTo 0

    ' Nicht-existierende Tabelle
    blnExists = CacheTabelleExistiert("tblGibtEsNicht_XYZ")
    AssertIsFalse blnExists, "Nicht-existierende Tabelle = False"

    CacheReset
    SuiteEnd
End Sub

