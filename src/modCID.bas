Option Compare Database
Option Explicit

' ===========================================================================
' modCID - Content-ID Inline-Bilder Verarbeitung
' ===========================================================================
' v0.4.2: Extrahiert CID-Referenzen aus HTML-Bodies und ersetzt sie
' durch lokale Dateipfade. Damit werden Inline-Bilder in gespeicherten
' Mails korrekt angezeigt.
'
' Workflow:
'   1. ExtrahiereCIDReferenzen() - Findet alle cid:xxx im HTML
'   2. MapCIDZuAnhaenge()        - Ordnet CID-Keys den Anhaengen zu
'   3. ErsetzeCIDLinks()         - Ersetzt cid:xxx durch file:///Pfad
'
' Convenience:
'   VerarbeiteCIDKomplett()      - Fuehrt 1-3 in einem Aufruf durch
'   BaueCIDSchluessel()          - CID-Key-Varianten erzeugen
'
' Abhaengigkeiten:
'   modStringUtils  (BereinigeAnhangName)
'   modLogging      (LogInfo, LogDebug, LogWarn, LogVBAError)
' ===========================================================================


' ===========================================================================
' CID-REFERENZEN AUS HTML EXTRAHIEREN
' ===========================================================================

' Sucht alle "cid:..." Referenzen im HTML-Body per RegExp.
' Gibt ein Dictionary zurueck: Key = CID-String (lowercase), Value = ""
' (Wert wird spaeter durch MapCIDZuAnhaenge befuellt)
'
' Rueckgabe: Scripting.Dictionary mit CID-Keys, oder Nothing bei Fehler
Public Function ExtrahiereCIDReferenzen(ByVal strHTML As String) As Object
    On Error GoTo ErrHandler

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Len(strHTML) = 0 Then
        Set ExtrahiereCIDReferenzen = dict
        Exit Function
    End If

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "cid:([^\"">' ]+)"

    Dim matches As Object
    Set matches = regex.Execute(strHTML)

    Dim m As Object
    Dim strKey As String
    For Each m In matches
        strKey = LCase(Trim(m.SubMatches(0)))
        If Len(strKey) > 0 And Not dict.Exists(strKey) Then
            dict.Add strKey, ""
        End If
    Next m

    LogDebug "CID-Referenzen gefunden: " & dict.Count, "CID"

    Set ExtrahiereCIDReferenzen = dict
    Set regex = Nothing
    Set matches = Nothing
    Exit Function

ErrHandler:
    HandleError "modCID", "ExtrahiereCIDReferenzen"
    Set ExtrahiereCIDReferenzen = CreateObject("Scripting.Dictionary")
End Function


' ===========================================================================
' CID-KEYS ZU ANHAENGEN ZUORDNEN
' ===========================================================================

' Ordnet gefundene CID-Keys den tatsaechlichen Anhaengen zu.
' Arbeitet mit TypAnhangDaten-Array (aus modMailExtract).
'
' Parameter:
'   dictCID       - Dictionary mit CID-Keys (aus ExtrahiereCIDReferenzen)
'   arrAnhaenge() - Array von TypAnhangDaten (Dateiname + TempPfad)
'   intAnzahl     - Anzahl der Anhaenge
'   strAnhangBasis - Basispfad fuer Anhang-Ablage (Netzwerk)
'
' Aktualisiert dictCID: Value wird auf den Dateipfad gesetzt.
' Rueckgabe: Anzahl erfolgreicher Zuordnungen
Public Function MapCIDZuAnhaenge(dictCID As Object, _
                                  arrAnhaenge() As TypAnhangDaten, _
                                  ByVal intAnzahl As Integer, _
                                  ByVal strAnhangBasis As String) As Long
    On Error GoTo ErrHandler

    If dictCID Is Nothing Or dictCID.Count = 0 Then
        MapCIDZuAnhaenge = 0
        Exit Function
    End If

    Dim lngMapped As Long: lngMapped = 0
    Dim cidKey As Variant
    Dim arrVarianten As Object
    Dim strDateiLower As String
    Dim i As Integer
    Dim v As Variant

    For Each cidKey In dictCID.Keys
        ' Bereits gemappt?
        If dictCID(cidKey) <> "" Then GoTo NaechsterCID

        ' Schluessel-Varianten erzeugen
        Set arrVarianten = BaueCIDSchluessel(CStr(cidKey))

        ' Gegen alle Anhaenge pruefen
        For i = 0 To intAnzahl - 1
            strDateiLower = LCase(arrAnhaenge(i).Dateiname)
            If Len(strDateiLower) = 0 Then GoTo NaechsterAnhang

            ' Jede CID-Variante pruefen
            For Each v In arrVarianten
                If strDateiLower = CStr(v) Or InStr(strDateiLower, CStr(v)) > 0 Then
                    ' Treffer - Pfad setzen
                    Dim strPfad As String
                    If Len(arrAnhaenge(i).TempPfad) > 0 Then
                        strPfad = arrAnhaenge(i).TempPfad
                    Else
                        strPfad = NormalisierePfad(strAnhangBasis) & arrAnhaenge(i).Dateiname
                    End If

                    dictCID(cidKey) = strPfad
                    lngMapped = lngMapped + 1
                    LogDebug "CID gemappt: " & cidKey & " -> " & arrAnhaenge(i).Dateiname, "CID"
                    GoTo NaechsterCID
                End If
            Next v
NaechsterAnhang:
        Next i

        LogTrace "Kein Anhang fuer CID: " & cidKey, "CID"
NaechsterCID:
    Next cidKey

    MapCIDZuAnhaenge = lngMapped
    LogDebug "CID-Mappings erstellt: " & lngMapped & "/" & dictCID.Count, "CID"
    Exit Function

ErrHandler:
    HandleError "modCID", "MapCIDZuAnhaenge"
    MapCIDZuAnhaenge = 0
End Function


' ===========================================================================
' CID-LINKS IM HTML ERSETZEN
' ===========================================================================

' Ersetzt alle cid:xxx-Referenzen im HTML durch file:///Pfad-URLs.
'
' Parameter:
'   strHTML     - HTML-Body der Mail
'   dictCID     - Dictionary: CID-Key -> Dateipfad (aus MapCIDZuAnhaenge)
'
' Rueckgabe: HTML mit ersetzten Links
Public Function ErsetzeCIDLinks(ByVal strHTML As String, _
                                 dictCID As Object) As String
    On Error GoTo ErrHandler

    If dictCID Is Nothing Or dictCID.Count = 0 Then
        ErsetzeCIDLinks = strHTML
        Exit Function
    End If

    Dim cidKey As Variant
    Dim strOld As String, strNew As String
    Dim lngErsetzt As Long: lngErsetzt = 0

    For Each cidKey In dictCID.Keys
        ' Nur ersetzen wenn Pfad vorhanden
        If Len(dictCID(cidKey)) = 0 Then GoTo NaechsterLink

        strOld = "cid:" & cidKey

        ' file:/// URL erzeugen (Backslash -> Slash, Leerzeichen -> %20)
        strNew = "file:///" & Replace(dictCID(cidKey), "\", "/")
        strNew = Replace(strNew, " ", "%20")

        If InStr(1, strHTML, strOld, vbTextCompare) > 0 Then
            strHTML = Replace(strHTML, strOld, strNew, , , vbTextCompare)
            lngErsetzt = lngErsetzt + 1
        End If

NaechsterLink:
    Next cidKey

    If lngErsetzt > 0 Then
        LogDebug "CID-Links ersetzt: " & lngErsetzt, "CID"
    End If

    ErsetzeCIDLinks = strHTML
    Exit Function

ErrHandler:
    HandleError "modCID", "ErsetzeCIDLinks"
    ErsetzeCIDLinks = strHTML
End Function


' ===========================================================================
' CONVENIENCE: KOMPLETT-VERARBEITUNG
' ===========================================================================

' Fuehrt CID-Extraktion, Mapping und Ersetzung in einem Schritt durch.
'
' Parameter:
'   strHTML        - HTML-Body (wird modifiziert zurueckgegeben)
'   arrAnhaenge()  - Anhang-Array (TypAnhangDaten)
'   intAnzahl      - Anzahl Anhaenge
'   strAnhangBasis - Basispfad fuer Anhaenge
'   blnHatInline   - (Out) True wenn Inline-Bilder gefunden wurden
'
' Rueckgabe: Modifizierter HTML-String
Public Function VerarbeiteCIDKomplett(ByVal strHTML As String, _
                                      arrAnhaenge() As TypAnhangDaten, _
                                      ByVal intAnzahl As Integer, _
                                      ByVal strAnhangBasis As String, _
                                      Optional ByRef blnHatInline As Boolean = False) As String
    On Error GoTo ErrHandler
    blnHatInline = False

    ' 1. CID-Referenzen finden
    Dim dictCID As Object
    Set dictCID = ExtrahiereCIDReferenzen(strHTML)

    If dictCID.Count = 0 Then
        VerarbeiteCIDKomplett = strHTML
        Exit Function
    End If

    ' 2. Anhaenge zuordnen
    Dim lngMapped As Long
    lngMapped = MapCIDZuAnhaenge(dictCID, arrAnhaenge, intAnzahl, strAnhangBasis)

    If lngMapped = 0 Then
        LogInfo "CID-Referenzen gefunden (" & dictCID.Count & ") aber keine Anhaenge zugeordnet", "CID"
        VerarbeiteCIDKomplett = strHTML
        Exit Function
    End If

    ' 3. Links ersetzen
    blnHatInline = True
    VerarbeiteCIDKomplett = ErsetzeCIDLinks(strHTML, dictCID)

    LogInfo "CID-Verarbeitung: " & lngMapped & " Inline-Bilder gemappt", "CID"
    Exit Function

ErrHandler:
    HandleError "modCID", "VerarbeiteCIDKomplett"
    VerarbeiteCIDKomplett = strHTML
End Function


' ===========================================================================
' CID-SCHLUESSEL-VARIANTEN ERZEUGEN
' ===========================================================================

' Erzeugt alternative Lookup-Keys fuer eine CID-Referenz:
'   "image001.png@01DA1234" -> ["image001.png@01da1234", "image001.png", "image001"]
'
' Rueckgabe: Collection mit Varianten (lowercase)
Public Function BaueCIDSchluessel(ByVal strCID As String) As Object
    On Error GoTo ErrHandler

    Dim col As Object
    Set col = CreateObject("Scripting.Dictionary")

    strCID = LCase(Trim(strCID))
    If Len(strCID) = 0 Then
        Set BaueCIDSchluessel = col
        Exit Function
    End If

    ' Original (lowercase)
    col.Add strCID, True

    ' Teil vor @ (haeufig: image001.png@01DA...)
    If InStr(strCID, "@") > 0 Then
        Dim strVorAt As String
        strVorAt = Left(strCID, InStr(strCID, "@") - 1)
        If Len(strVorAt) > 0 And Not col.Exists(strVorAt) Then
            col.Add strVorAt, True
        End If

        ' Ohne Extension (image001)
        If InStrRev(strVorAt, ".") > 0 Then
            Dim strOhneExt As String
            strOhneExt = Left(strVorAt, InStrRev(strVorAt, ".") - 1)
            If Len(strOhneExt) > 0 And Not col.Exists(strOhneExt) Then
                col.Add strOhneExt, True
            End If
        End If
    End If

    Set BaueCIDSchluessel = col
    Exit Function

ErrHandler:
    HandleError "modCID", "BaueCIDSchluessel"
    Set BaueCIDSchluessel = CreateObject("Scripting.Dictionary")
End Function


