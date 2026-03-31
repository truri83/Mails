Option Compare Database
Option Explicit

' ===========================================================================
' modCache - Zentrales Cache-System fuer wiederkehrende Abfragen
' ===========================================================================
' v0.5: Vermeidet wiederholte DB-Abfragen fuer haeufig gelesene Daten.
'
' CACHES:
'   Config-Cache     - LeseConfig-Werte (vermeidet DLookup pro Aufruf)
'   Tabellen-Cache   - TabelleExistiert-Ergebnisse
'   Kontakt-Cache    - Email -> KontaktID (haeufigste Abfrage)
'
' STEUERUNG:
'   CacheInit          - Alle Caches initialisieren
'   CacheReset         - Alle Caches leeren
'   CacheResetConfig   - Nur Config-Cache leeren (nach SchreibeConfig)
'   CacheResetKontakte - Nur Kontakt-Cache leeren
'   CacheStatus        - Statistik ausgeben (Hits/Misses)
'
' NUTZUNG:
'   CacheGetConfig(strKey, strDefault) As String    - statt DLookup
'   CacheSetConfig(strKey, strVal)                  - Cache + DB aktualisieren
'   CacheTabelleExistiert(strName) As Boolean       - mit Cache
'   CacheGetKontaktID(strEmail) As Long             - 0 = nicht gefunden
'   CacheSetKontaktID(strEmail, lngID)              - Eintrag cachen
'
' Abhaengigkeiten: modGlobals (TBL_*, CFG_*), modSchema (LeseConfig),
'                  modLogging
' ===========================================================================


' ---------------------------------------------------------------------------
' CACHE-DICTIONARYS
' ---------------------------------------------------------------------------
Private m_dictConfig    As Object   ' Key -> Wert (String)
Private m_dictTabellen  As Object   ' Tabellenname -> Boolean
Private m_dictKontakte  As Object   ' Email (lowercase) -> KontaktID (Long)

' STATISTIK
Private m_lngConfigHit      As Long
Private m_lngConfigMiss     As Long
Private m_lngTabellenHit    As Long
Private m_lngTabellenMiss   As Long
Private m_lngKontaktHit     As Long
Private m_lngKontaktMiss    As Long

' Initialisiert-Flag
Private m_blnInitialized    As Boolean


' ===========================================================================
' INITIALISIERUNG / RESET
' ===========================================================================

Public Sub CacheInit()
    Set m_dictConfig = CreateObject("Scripting.Dictionary")
    Set m_dictTabellen = CreateObject("Scripting.Dictionary")
    Set m_dictKontakte = CreateObject("Scripting.Dictionary")
    m_dictKontakte.CompareMode = vbTextCompare  ' Case-insensitive fuer Emails

    m_lngConfigHit = 0:   m_lngConfigMiss = 0
    m_lngTabellenHit = 0: m_lngTabellenMiss = 0
    m_lngKontaktHit = 0:  m_lngKontaktMiss = 0
    m_blnInitialized = True

    LogTrace "Cache initialisiert", "CACHE"
End Sub

Public Sub CacheReset()
    If Not m_blnInitialized Then Exit Sub

    m_dictConfig.RemoveAll
    m_dictTabellen.RemoveAll
    m_dictKontakte.RemoveAll

    m_lngConfigHit = 0:   m_lngConfigMiss = 0
    m_lngTabellenHit = 0: m_lngTabellenMiss = 0
    m_lngKontaktHit = 0:  m_lngKontaktMiss = 0

    LogDebug "Alle Caches geleert", "CACHE"
End Sub

Public Sub CacheResetConfig()
    If Not m_blnInitialized Then Exit Sub
    m_dictConfig.RemoveAll
    m_lngConfigHit = 0: m_lngConfigMiss = 0
End Sub

Public Sub CacheResetKontakte()
    If Not m_blnInitialized Then Exit Sub
    m_dictKontakte.RemoveAll
    m_lngKontaktHit = 0: m_lngKontaktMiss = 0
End Sub


' ===========================================================================
' CONFIG-CACHE
' ===========================================================================

' Cached LeseConfig - vermeidet wiederholte DLookup-Aufrufe
Public Function CacheGetConfig(ByVal strKey As String, _
                                Optional ByVal strDefault As String = "") As String
    If Not m_blnInitialized Then CacheInit

    If m_dictConfig.Exists(strKey) Then
        m_lngConfigHit = m_lngConfigHit + 1
        CacheGetConfig = CStr(m_dictConfig(strKey))
        Exit Function
    End If

    ' Cache-Miss: aus DB laden
    m_lngConfigMiss = m_lngConfigMiss + 1
    Dim strVal As String
    strVal = LeseConfig(strKey, strDefault)
    m_dictConfig(strKey) = strVal
    CacheGetConfig = strVal
End Function

' Config-Wert setzen: Cache UND DB aktualisieren
Public Sub CacheSetConfig(ByVal strKey As String, ByVal strVal As String)
    If Not m_blnInitialized Then CacheInit

    SchreibeConfig strKey, strVal
    m_dictConfig(strKey) = strVal
End Sub


' ===========================================================================
' TABELLEN-CACHE
' ===========================================================================

' Cached TabelleExistiert - Ergebnis aendert sich zur Laufzeit quasi nie
Public Function CacheTabelleExistiert(ByVal strName As String) As Boolean
    If Not m_blnInitialized Then CacheInit

    If m_dictTabellen.Exists(strName) Then
        m_lngTabellenHit = m_lngTabellenHit + 1
        CacheTabelleExistiert = CBool(m_dictTabellen(strName))
        Exit Function
    End If

    ' Cache-Miss: wirklich pruefen
    m_lngTabellenMiss = m_lngTabellenMiss + 1
    Dim blnExists As Boolean
    blnExists = TabelleExistiert(strName)
    m_dictTabellen(strName) = blnExists
    CacheTabelleExistiert = blnExists
End Function


' ===========================================================================
' KONTAKT-CACHE
' ===========================================================================

' Kontakt-ID per Email aus Cache holen (0 = nicht im Cache)
Public Function CacheGetKontaktID(ByVal strEmail As String) As Long
    If Not m_blnInitialized Then CacheInit

    Dim strKey As String
    strKey = LCase(Trim(strEmail))

    If m_dictKontakte.Exists(strKey) Then
        m_lngKontaktHit = m_lngKontaktHit + 1
        CacheGetKontaktID = CLng(m_dictKontakte(strKey))
    Else
        m_lngKontaktMiss = m_lngKontaktMiss + 1
        CacheGetKontaktID = 0
    End If
End Function

' Kontakt-ID in Cache eintragen (nach DB-Lookup oder Neuanlage)
Public Sub CacheSetKontaktID(ByVal strEmail As String, ByVal lngID As Long)
    If Not m_blnInitialized Then CacheInit

    Dim strKey As String
    strKey = LCase(Trim(strEmail))
    m_dictKontakte(strKey) = lngID
End Sub


' ===========================================================================
' STATISTIK
' ===========================================================================

Public Sub CacheStatus()
    Debug.Print String(50, "-")
    Debug.Print "=== CACHE-STATISTIK ==="

    If Not m_blnInitialized Then
        Debug.Print "  Cache nicht initialisiert."
        Debug.Print String(50, "-")
        Exit Sub
    End If

    Debug.Print "  Config:   " & m_dictConfig.Count & " Eintraege, " & _
                m_lngConfigHit & " Hits / " & m_lngConfigMiss & " Misses" & _
                FormatHitRate(m_lngConfigHit, m_lngConfigMiss)

    Debug.Print "  Tabellen: " & m_dictTabellen.Count & " Eintraege, " & _
                m_lngTabellenHit & " Hits / " & m_lngTabellenMiss & " Misses" & _
                FormatHitRate(m_lngTabellenHit, m_lngTabellenMiss)

    Debug.Print "  Kontakte: " & m_dictKontakte.Count & " Eintraege, " & _
                m_lngKontaktHit & " Hits / " & m_lngKontaktMiss & " Misses" & _
                FormatHitRate(m_lngKontaktHit, m_lngKontaktMiss)

    Debug.Print String(50, "-")
End Sub


' Hilfsfunktion: Formatiert Hit-Rate sicher (kein Division-by-Zero)
' IIf() evaluiert in VBA IMMER beide Zweige -> Ueberlauf bei /0!
Private Function FormatHitRate(ByVal lngHit As Long, ByVal lngMiss As Long) As String
    Dim lngTotal As Long
    lngTotal = lngHit + lngMiss
    If lngTotal > 0 Then
        FormatHitRate = " (" & Format(CDbl(lngHit) / CDbl(lngTotal) * 100, "0") & "%)"
    Else
        FormatHitRate = ""
    End If
End Function


