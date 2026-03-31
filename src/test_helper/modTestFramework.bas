Option Compare Database
Option Explicit

' ===========================================================================
' modTestFramework - Assert-Funktionen, Suite-Verwaltung, Reporting
' ===========================================================================
' Stellt die gemeinsame Test-Infrastruktur fuer alle Test-Module bereit:
'   - TestRunStart / TestRunEnd      Gesamt-Lauf starten/beenden
'   - SuiteStart / SuiteEnd          Test-Suite starten/beenden
'   - AssertIsTrue / AssertIsFalse   Boolean-Pruefungen
'   - AssertAreEqual / AssertAreNotEqual   Gleichheits-Pruefungen
'   - AssertContains                 Teilstring-Pruefung
'   - AssertIsNotEmpty               Nicht-Leer-Pruefung
'   - AssertFail                     Test explizit als Fehler markieren
'   - AssertSkip                     Test uebersprungen markieren
'
' AUFRUF: Wird von modTestUnit, modTestDAO, modTestIntegration etc. verwendet.
' ===========================================================================

Private Const MODUL_NAME As String = "modTestFramework"

' ---------------------------------------------------------------------------
' Zaehler fuer den aktuellen Test-Lauf
' ---------------------------------------------------------------------------
Private m_lngPassed   As Long
Private m_lngFailed   As Long
Private m_lngSkipped  As Long
Private m_strLaufName As String

' ---------------------------------------------------------------------------
' Zaehler fuer die aktuelle Test-Suite
' ---------------------------------------------------------------------------
Private m_lngSuitePassed  As Long
Private m_lngSuiteFailed  As Long
Private m_lngSuiteSkipped As Long
Private m_strSuiteName    As String


' ===========================================================================
' LAUF-VERWALTUNG
' ===========================================================================

' Startet einen neuen Test-Lauf und setzt alle Zaehler zurueck.
Public Sub TestRunStart(ByVal strName As String)
    m_strLaufName  = strName
    m_lngPassed   = 0
    m_lngFailed   = 0
    m_lngSkipped  = 0
    Debug.Print ""
    Debug.Print String(60, "-")
    Debug.Print "  TEST-LAUF: " & strName
    Debug.Print "  " & Format(Now, "dd.mm.yyyy hh:nn:ss")
    Debug.Print String(60, "-")
End Sub

' Beendet den aktuellen Test-Lauf und gibt eine Zusammenfassung aus.
Public Sub TestRunEnd()
    Dim lngGesamt As Long
    lngGesamt = m_lngPassed + m_lngFailed + m_lngSkipped
    Debug.Print String(60, "-")
    Debug.Print "  ERGEBNIS: " & m_strLaufName
    Debug.Print "  Gesamt : " & lngGesamt
    Debug.Print "  OK     : " & m_lngPassed
    Debug.Print "  FAIL   : " & m_lngFailed
    Debug.Print "  SKIP   : " & m_lngSkipped
    If m_lngFailed = 0 Then
        Debug.Print "  >>> ALLE TESTS BESTANDEN <<<"
    Else
        Debug.Print "  >>> " & m_lngFailed & " TEST(S) FEHLGESCHLAGEN <<<"
    End If
    Debug.Print String(60, "-")
End Sub


' ===========================================================================
' SUITE-VERWALTUNG
' ===========================================================================

' Startet eine neue Test-Suite innerhalb des aktuellen Laufs.
Public Sub SuiteStart(ByVal strName As String)
    m_strSuiteName    = strName
    m_lngSuitePassed  = 0
    m_lngSuiteFailed  = 0
    m_lngSuiteSkipped = 0
    Debug.Print ""
    Debug.Print "  [Suite] " & strName
End Sub

' Beendet die aktuelle Test-Suite und gibt ein Kurz-Ergebnis aus.
Public Sub SuiteEnd()
    Dim strStatus As String
    If m_lngSuiteFailed = 0 Then
        strStatus = "OK"
    Else
        strStatus = "FAIL"
    End If
    Debug.Print "  [Suite] " & m_strSuiteName & " => " & strStatus & _
                " (" & m_lngSuitePassed & " OK / " & m_lngSuiteFailed & " FAIL / " & m_lngSuiteSkipped & " SKIP)"
End Sub


' ===========================================================================
' ASSERT-FUNKTIONEN
' ===========================================================================

' Prueft ob ein boolescher Ausdruck True ist.
Public Sub AssertIsTrue(ByVal blnErgebnis As Boolean, ByVal strBeschreibung As String)
    If blnErgebnis Then
        TestPass strBeschreibung
    Else
        TestFail strBeschreibung & " (erwartet: True, erhalten: False)"
    End If
End Sub

' Prueft ob ein boolescher Ausdruck False ist.
Public Sub AssertIsFalse(ByVal blnErgebnis As Boolean, ByVal strBeschreibung As String)
    If Not blnErgebnis Then
        TestPass strBeschreibung
    Else
        TestFail strBeschreibung & " (erwartet: False, erhalten: True)"
    End If
End Sub

' Prueft ob zwei Werte gleich sind (Vergleich als String).
Public Sub AssertAreEqual(ByVal varErwartet As Variant, ByVal varErhalten As Variant, _
                          ByVal strBeschreibung As String)
    If CStr(varErwartet) = CStr(varErhalten) Then
        TestPass strBeschreibung
    Else
        TestFail strBeschreibung & " (erwartet: [" & CStr(varErwartet) & "] erhalten: [" & CStr(varErhalten) & "])"
    End If
End Sub

' Prueft ob zwei Werte NICHT gleich sind.
Public Sub AssertAreNotEqual(ByVal varErwartet As Variant, ByVal varErhalten As Variant, _
                             ByVal strBeschreibung As String)
    If CStr(varErwartet) <> CStr(varErhalten) Then
        TestPass strBeschreibung
    Else
        TestFail strBeschreibung & " (beide gleich: [" & CStr(varErwartet) & "])"
    End If
End Sub

' Prueft ob ein String einen Teilstring enthaelt.
Public Sub AssertContains(ByVal strHaystack As String, ByVal strNeedle As String, _
                          ByVal strBeschreibung As String)
    If InStr(1, strHaystack, strNeedle, vbTextCompare) > 0 Then
        TestPass strBeschreibung
    Else
        TestFail strBeschreibung & " ([" & strNeedle & "] nicht in [" & Left$(strHaystack, 80) & "])"
    End If
End Sub

' Prueft ob ein String nicht leer ist.
Public Sub AssertIsNotEmpty(ByVal strWert As String, ByVal strBeschreibung As String)
    If Len(Trim$(strWert)) > 0 Then
        TestPass strBeschreibung
    Else
        TestFail strBeschreibung & " (Wert ist leer)"
    End If
End Sub

' Markiert einen Test explizit als fehlgeschlagen.
Public Sub AssertFail(ByVal strBeschreibung As String)
    TestFail strBeschreibung
End Sub

' Markiert einen Test als uebersprungen (z.B. fehlende Voraussetzung).
Public Sub AssertSkip(ByVal strBeschreibung As String)
    m_lngSkipped      = m_lngSkipped + 1
    m_lngSuiteSkipped = m_lngSuiteSkipped + 1
    Debug.Print "    [SKIP] " & strBeschreibung
End Sub


' ===========================================================================
' PRIVATE HILFSFUNKTIONEN
' ===========================================================================

Private Sub TestPass(ByVal strBeschreibung As String)
    m_lngPassed      = m_lngPassed + 1
    m_lngSuitePassed = m_lngSuitePassed + 1
    Debug.Print "    [OK  ] " & strBeschreibung
End Sub

Private Sub TestFail(ByVal strBeschreibung As String)
    m_lngFailed      = m_lngFailed + 1
    m_lngSuiteFailed = m_lngSuiteFailed + 1
    Debug.Print "    [FAIL] " & strBeschreibung
End Sub
