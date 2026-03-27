Option Compare Database
Option Explicit

' ===========================================================================
' modRedemptionLoader - Redemption DLL ohne COM-Registrierung laden
' ===========================================================================
' v0.5.1: Basiert auf dem offiziellen RedemptionLoader
'   (https://www.dimastr.com/redemption/security.htm#redemptionloader)
'
' Laedt Redemption-DLLs direkt per LoadLibrary statt ueber COM-Registry.
' Vorteile:
'   - Keine regsvr32-Registrierung noetig (kein Admin erforderlich)
'   - Portabler Einsatz (DLLs im Anwendungsverzeichnis)
'   - Fallback auf CreateObject wenn DLL nicht verfuegbar
'
' DLL-Suchpfad (Prioritaet):
'   1. CFG_RDO_PFAD aus tblConfig (konfigurierbar)
'   2. CurrentProject.Path (Verzeichnis der .accdb)
'
' OEFFENTLICH:
'   ErstelleRedemptionObjekt(strName) -> Redemption-Objekt erstellen
'     z.B. "RDOSession", "SafeMailItem"
'
' TECHNIK:
'   LoadLibrary laedt die DLL vorab in den Prozess. Danach finden die
'   Declare-Anweisungen (nur Dateiname, kein Pfad) die bereits geladene
'   DLL automatisch ueber den Windows-Modulnamen.
'
' Abhaengigkeiten: modGlobals (CFG_RDO_PFAD, RDO_DLL_*),
'                  modSchema (LeseConfig), modLogging
' ===========================================================================


' ---------------------------------------------------------------------------
' KERNEL32 API
' ---------------------------------------------------------------------------
Private Declare PtrSafe Function LoadLibraryA Lib "kernel32" _
    (ByVal lpFileName As String) As LongPtr


' ---------------------------------------------------------------------------
' REDEMPTION DLL-FUNKTIONEN
' Nur Dateiname (kein Pfad) - DLL wird vorab per LoadLibrary geladen,
' sodass Windows sie am Modulnamen erkennt.
' VBA loest Declare-Funktionen lazy auf (erst beim ersten Aufruf).
' ---------------------------------------------------------------------------

' 32-Bit
Private Declare PtrSafe Function GetRedemptionObject32 Lib "Redemption.dll" _
    Alias "GetRedemptionObject" _
    (ByVal ClassName As LongPtr, ByRef COMObject As Object) As Long

' 64-Bit
Private Declare PtrSafe Function GetRedemptionObject64 Lib "Redemption64.dll" _
    Alias "GetRedemptionObject" _
    (ByVal ClassName As LongPtr, ByRef COMObject As Object) As Long


' ---------------------------------------------------------------------------
' MODUL-STATUS
' ---------------------------------------------------------------------------
Private m_blnInitialized    As Boolean
Private m_blnDLLGeladen     As Boolean  ' True = DLL per LoadLibrary geladen


' ===========================================================================
' INITIALISIERUNG (lazy, wird beim ersten Aufruf ausgefuehrt)
' ===========================================================================
Private Sub InitLoader()
    If m_blnInitialized Then Exit Sub
    m_blnInitialized = True

    On Error Resume Next

    ' DLL-Pfad ermitteln: Config -> Projektverzeichnis
    Dim strPfad As String
    strPfad = LeseConfig(CFG_RDO_PFAD, "")
    If strPfad = "" Then strPfad = CurrentProject.Path
    If Right(strPfad, 1) <> "\" Then strPfad = strPfad & "\"

    ' DLL-Datei bestimmen (32/64-Bit)
    Dim strDLL As String
    #If Win64 Then
        strDLL = strPfad & RDO_DLL_64
    #Else
        strDLL = strPfad & RDO_DLL_32
    #End If

    ' DLL vorab laden - Declare-Anweisungen finden sie dann am Modulnamen
    Dim hLib As LongPtr
    hLib = LoadLibraryA(strDLL)
    m_blnDLLGeladen = (hLib <> 0)

    If m_blnDLLGeladen Then
        LogInfo "Redemption DLL geladen: " & strDLL, "RDO"
    Else
        LogDebug "Redemption DLL nicht unter " & strDLL & _
                 " - nutze COM-Registrierung", "RDO"
    End If

    On Error GoTo 0
End Sub


' ===========================================================================
' REDEMPTION-OBJEKT ERSTELLEN (Public)
' ===========================================================================
' Erstellt ein Redemption-Objekt. Versucht zuerst den DLL-Loader (kein
' regsvr32 noetig), dann Fallback auf CreateObject (COM-Registrierung).
'
' Parameter:
'   strName  - Klassenname ohne "Redemption."-Praefix
'              z.B. "RDOSession", "SafeMailItem", "SafeContactItem"
'
' Beispiel:
'   Set g_objRDO = ErstelleRedemptionObjekt("RDOSession")
'   Set objSafe = ErstelleRedemptionObjekt("SafeMailItem")
' ---------------------------------------------------------------------------
Public Function ErstelleRedemptionObjekt(ByVal strName As String) As Object
    InitLoader

    ' --- Versuch 1: DLL-Loader (ohne COM-Registrierung) ---
    If m_blnDLLGeladen Then
        On Error Resume Next

        Dim lngResult As Long
        Dim objResult As Object

        #If Win64 Then
            lngResult = GetRedemptionObject64(StrPtr(strName), objResult)
        #Else
            lngResult = GetRedemptionObject32(StrPtr(strName), objResult)
        #End If

        If Err.Number = 0 And lngResult = 0 And Not objResult Is Nothing Then
            Set ErstelleRedemptionObjekt = objResult
            On Error GoTo 0
            Exit Function
        End If

        ' DLL-Loader fehlgeschlagen, Fallback versuchen
        Err.Clear
        On Error GoTo 0
    End If

    ' --- Versuch 2: CreateObject (erfordert COM-Registrierung) ---
    On Error GoTo ErrHandler
    Set ErstelleRedemptionObjekt = CreateObject("Redemption." & strName)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "modRedemptionLoader", _
              "Redemption-Objekt '" & strName & "' nicht erstellbar. " & _
              "Weder DLL-Loader noch COM-Registrierung verfuegbar. " & _
              "Pruefen: DLL im Pfad (CFG_RDO_PFAD) oder regsvr32 ausfuehren."
End Function


