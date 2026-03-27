Option Compare Database
Option Explicit

' ===========================================================================
' modProjekte - Projekt-Verwaltung und Email-Projekt-Zuordnung (v0.6)
' ===========================================================================
' Zentrale Logik fuer:
'   - Projekt-CRUD (Erstellen, Suchen, Auflisten)
'   - Email-Projekt n:m Zuordnung (Taggen/Entfernen)
'   - "Untagged"-Konzept (Mails ohne Projekt-Zuordnung)
'   - Datei-Verschiebung (_Eingang -> Projektordner)
'
' Abhaengigkeiten: modGlobals (Konstanten), modDAO (SpeichereEmailStatus),
'                  modBackend (PruefeBackendVorZugriff, BehandleNetzwerkFehler),
'                  modStringUtils (SQLSafe, BereinigeDateiname, NormalisierePfad),
'                  modFileManager (BaueMailOrdnerPfad, KopiereNachNetzwerk, ErstelleOrdner),
'                  modLogging, modDevUtils (HandleError)
' ===========================================================================


' ===================================================================
' SECTION: PROJEKT-CRUD
' ===================================================================

' Erstellt ein neues Projekt -> gibt ProjektID zurueck (0 bei Fehler)
Public Function ErstelleProjekt(ByVal strName As String, _
                                 Optional ByVal strKuerzel As String = "", _
                                 Optional ByVal strBeschreibung As String = "", _
                                 Optional ByVal strPhase As String = "") As Long
    On Error GoTo ErrHandler
    ErstelleProjekt = 0

    If Trim(Nz(strName, "")) = "" Then
        LogWarn "ErstelleProjekt: Leerer Name", "PROJEKT"
        Exit Function
    End If

    ' Netzwerk-Schutz
    If Not PruefeBackendVorZugriff() Then
        LogWarn "ErstelleProjekt: Backend offline", "PROJEKT"
        Exit Function
    End If

    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb

    Set rs = db.OpenRecordset(TBL_PROJEKTE, dbOpenDynaset)
    With rs
        .AddNew
        !Name = Left(strName, 100)
        If strKuerzel <> "" Then !Kuerzel = Left(strKuerzel, 20)
        If strBeschreibung <> "" Then !Beschreibung = Left(strBeschreibung, 255)
        If strPhase <> "" Then !Phase = Left(strPhase, 100)
        !Status = PROJ_STATUS_AKTIV
        !SortierNr = 0
        !ErstelltVon = Left(Environ("USERNAME"), 100)
        !ErstelltAm = Now
        !AktualisiertAm = Now
        .Update
        .Bookmark = .LastModified
        ErstelleProjekt = !ProjektID
    End With
    rs.Close: Set rs = Nothing

    LogInfo "Projekt erstellt: " & strName & " (ID=" & ErstelleProjekt & ")", "PROJEKT"
    Set db = Nothing
    Exit Function

ErrHandler:
    If Err.Number = 3022 Then
        ' UNIQUE Constraint verletzt -> Projekt existiert bereits
        Err.Clear
        ErstelleProjekt = HoleProjektID(strName)
        Exit Function
    End If
    If BehandleNetzwerkFehler(Err.Number, "modProjekte", "ErstelleProjekt") Then Exit Function
    HandleError "modProjekte", "ErstelleProjekt"
End Function


' Sucht ProjektID nach Name -> 0 wenn nicht gefunden
Public Function HoleProjektID(ByVal strName As String) As Long
    On Error Resume Next
    Dim varID As Variant
    varID = DLookup("ProjektID", TBL_PROJEKTE, "Name='" & SQLSafe(strName) & "'")
    HoleProjektID = Nz(varID, 0)
    If Err.Number <> 0 Then HoleProjektID = 0: Err.Clear
    On Error GoTo 0
End Function


' Sucht oder erstellt Projekt -> gibt ProjektID zurueck
Public Function GetOderErstelleProjekt(ByVal strName As String, _
                                        Optional ByVal strKuerzel As String = "", _
                                        Optional ByVal strPhase As String = "") As Long
    GetOderErstelleProjekt = HoleProjektID(strName)
    If GetOderErstelleProjekt = 0 Then
        GetOderErstelleProjekt = ErstelleProjekt(strName, strKuerzel, , strPhase)
    End If
End Function


' Gibt den Ordnernamen fuer ein Projekt zurueck (Kuerzel falls vorhanden, sonst Name)
Public Function HoleProjektOrdnerName(ByVal lngProjektID As Long) As String
    On Error Resume Next
    Dim varKuerzel As Variant, varName As Variant
    Dim strFilter As String
    strFilter = "ProjektID=" & lngProjektID

    varKuerzel = DLookup("Kuerzel", TBL_PROJEKTE, strFilter)
    If Nz(varKuerzel, "") <> "" Then
        HoleProjektOrdnerName = BereinigeDateiname(CStr(varKuerzel), 30)
    Else
        varName = DLookup("Name", TBL_PROJEKTE, strFilter)
        HoleProjektOrdnerName = BereinigeDateiname(Nz(varName, "_Eingang"), 30)
    End If

    If Err.Number <> 0 Then HoleProjektOrdnerName = "_Eingang": Err.Clear
    On Error GoTo 0
End Function


' Gibt alle aktiven Projekte als Recordset zurueck (fuer UI-Auswahl)
Public Function HoleAktiveProjekte() As DAO.Recordset
    On Error GoTo ErrHandler
    Set HoleAktiveProjekte = CurrentDb.OpenRecordset( _
        "SELECT ProjektID, Name, Kuerzel, Beschreibung, Phase, Status, Farbe " & _
        "FROM [" & TBL_PROJEKTE & "] " & _
        "WHERE Status='" & PROJ_STATUS_AKTIV & "' " & _
        "ORDER BY SortierNr, Name", dbOpenSnapshot)
    Exit Function
ErrHandler:
    Set HoleAktiveProjekte = Nothing
    HandleError "modProjekte", "HoleAktiveProjekte"
End Function


' ===================================================================
' SECTION: EMAIL-PROJEKT ZUORDNUNG (n:m)
' ===================================================================

' Ordnet eine Email einem Projekt zu
' Bei Mails in _Eingang wird automatisch VerschiebeNachProjekt aufgerufen
Public Sub OrdneEmailProjektZu(ByVal lngEmailID As Long, _
                                ByVal lngProjektID As Long, _
                                Optional ByVal strQuelle As String = "Manuell")
    On Error GoTo ErrHandler

    If lngEmailID = 0 Or lngProjektID = 0 Then Exit Sub

    ' Netzwerk-Schutz
    If Not PruefeBackendVorZugriff() Then
        LogWarn "OrdneEmailProjektZu: Backend offline", "PROJEKT"
        Exit Sub
    End If

    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb

    Set rs = db.OpenRecordset(TBL_EMAIL_PROJEKT, dbOpenDynaset)
    With rs
        .AddNew
        !EmailID = lngEmailID
        !ProjektID = lngProjektID
        !Quelle = Left(Nz(strQuelle, EP_QUELLE_MANUELL), 20)
        !ZugeordnetVon = Left(Environ("USERNAME"), 100)
        !ZugeordnetAm = Now
        .Update
    End With
    rs.Close: Set rs = Nothing

    LogTrace "Email " & lngEmailID & " -> Projekt " & lngProjektID & " (" & strQuelle & ")", "PROJEKT"

    ' --- Automatische Datei-Verschiebung wenn Mail in _Eingang liegt ---
    Dim varMSGPfad As Variant
    varMSGPfad = DLookup("MSGDateiPfad", TBL_EMAILS, "EmailID=" & lngEmailID)
    If Nz(varMSGPfad, "") <> "" Then
        If InStr(1, CStr(varMSGPfad), "_Eingang", vbTextCompare) > 0 Then
            VerschiebeNachProjekt lngEmailID, lngProjektID
        End If
    End If

    Set db = Nothing
    Exit Sub

ErrHandler:
    If Err.Number = 3022 Then
        ' Duplikat (EmailID+ProjektID existiert schon) -> OK, idempotent
        Err.Clear
        Exit Sub
    End If
    If BehandleNetzwerkFehler(Err.Number, "modProjekte", "OrdneEmailProjektZu") Then Exit Sub
    HandleError "modProjekte", "OrdneEmailProjektZu"
End Sub


' Entfernt eine Email aus einem Projekt
Public Sub EntferneEmailAusProjekt(ByVal lngEmailID As Long, _
                                    ByVal lngProjektID As Long)
    On Error GoTo ErrHandler

    If lngEmailID = 0 Or lngProjektID = 0 Then Exit Sub

    If Not PruefeBackendVorZugriff() Then
        LogWarn "EntferneEmailAusProjekt: Backend offline", "PROJEKT"
        Exit Sub
    End If

    CurrentDb.Execute "DELETE FROM [" & TBL_EMAIL_PROJEKT & "] " & _
        "WHERE EmailID=" & lngEmailID & " AND ProjektID=" & lngProjektID, dbFailOnError

    LogTrace "Email " & lngEmailID & " aus Projekt " & lngProjektID & " entfernt", "PROJEKT"
    Exit Sub

ErrHandler:
    If BehandleNetzwerkFehler(Err.Number, "modProjekte", "EntferneEmailAusProjekt") Then Exit Sub
    HandleError "modProjekte", "EntferneEmailAusProjekt"
End Sub


' Gibt komma-getrennte ProjektIDs fuer eine Email zurueck (leer wenn keine)
Public Function HoleProjekteFuerEmail(ByVal lngEmailID As Long) As String
    On Error Resume Next
    Dim rs As DAO.Recordset
    Dim strResult As String: strResult = ""

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT p.Name FROM [" & TBL_EMAIL_PROJEKT & "] ep " & _
        "INNER JOIN [" & TBL_PROJEKTE & "] p ON ep.ProjektID = p.ProjektID " & _
        "WHERE ep.EmailID=" & lngEmailID & " ORDER BY p.Name", dbOpenSnapshot)

    Do While Not rs.EOF
        If strResult <> "" Then strResult = strResult & ", "
        strResult = strResult & Nz(rs!Name, "")
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing

    HoleProjekteFuerEmail = strResult
    If Err.Number <> 0 Then HoleProjekteFuerEmail = "": Err.Clear
    On Error GoTo 0
End Function


' ===================================================================
' SECTION: UNTAGGED / STATUS-VERWALTUNG
' ===================================================================

' Markiert eine Email als irrelevant (verleast die Untagged-Queue)
Public Sub MarkiereAlsIrrelevant(ByVal lngEmailID As Long, _
                                  Optional ByVal strBemerkung As String = "")
    SpeichereEmailStatus lngEmailID, EMAIL_STATUS_IRRELEVANT, Environ("USERNAME"), strBemerkung
    LogTrace "Email " & lngEmailID & " als irrelevant markiert", "PROJEKT"
End Sub


' Markiert eine Email als archiviert
Public Sub MarkiereAlsArchiviert(ByVal lngEmailID As Long, _
                                  Optional ByVal strBemerkung As String = "")
    SpeichereEmailStatus lngEmailID, EMAIL_STATUS_ARCHIVIERT, Environ("USERNAME"), strBemerkung
    LogTrace "Email " & lngEmailID & " als archiviert markiert", "PROJEKT"
End Sub


' Zaehlt Mails ohne Projekt-Zuordnung (und nicht Irrelevant/Archiviert)
Public Function AnzahlUngetaggt() As Long
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Count(*) AS Anz FROM [" & TBL_EMAILS & "] e " & _
        "LEFT JOIN [" & TBL_EMAIL_PROJEKT & "] ep ON e.EmailID = ep.EmailID " & _
        "WHERE ep.EmailProjektID IS NULL " & _
        "AND e.Status NOT IN ('" & EMAIL_STATUS_IRRELEVANT & "','" & EMAIL_STATUS_ARCHIVIERT & "')", _
        dbOpenSnapshot)
    AnzahlUngetaggt = Nz(rs!Anz, 0)
    rs.Close: Set rs = Nothing
    If Err.Number <> 0 Then AnzahlUngetaggt = -1: Err.Clear
    On Error GoTo 0
End Function


' ===================================================================
' SECTION: DATEI-VERSCHIEBUNG (_Eingang -> Projektordner)
' ===================================================================

' Verschiebt Mail-Dateien (MSG + Anhaenge) von _Eingang in den Projektordner.
' Aktualisiert MSGDateiPfad und AnhangDateiPfad in der DB.
'
' Wird automatisch von OrdneEmailProjektZu() aufgerufen wenn die Mail
' aktuell in _Eingang liegt. Kann auch manuell aufgerufen werden.
'
' Returns: True wenn erfolgreich (oder nichts zu tun war)
Public Function VerschiebeNachProjekt(ByVal lngEmailID As Long, _
                                      ByVal lngProjektID As Long) As Boolean
    On Error GoTo ErrHandler
    VerschiebeNachProjekt = False

    ' --- Aktuellen MSG-Pfad lesen ---
    Dim varMSGPfad As Variant
    varMSGPfad = DLookup("MSGDateiPfad", TBL_EMAILS, "EmailID=" & lngEmailID)
    Dim strAlterMSGPfad As String
    strAlterMSGPfad = Nz(varMSGPfad, "")

    ' Nichts zu tun wenn kein Pfad oder nicht in _Eingang
    If strAlterMSGPfad = "" Then
        VerschiebeNachProjekt = True
        Exit Function
    End If
    If InStr(1, strAlterMSGPfad, "_Eingang", vbTextCompare) = 0 Then
        VerschiebeNachProjekt = True
        Exit Function
    End If

    ' --- Projektordner bestimmen ---
    Dim strProjektOrdner As String
    strProjektOrdner = HoleProjektOrdnerName(lngProjektID)
    If strProjektOrdner = "" Or strProjektOrdner = "_Eingang" Then
        VerschiebeNachProjekt = True
        Exit Function
    End If

    ' --- Neuen Pfad berechnen (ersetze _Eingang durch Projektordner) ---
    Dim strNeuerMSGPfad As String
    strNeuerMSGPfad = ErsetzePfadSegment(strAlterMSGPfad, "_Eingang", strProjektOrdner)

    ' --- MSG-Datei verschieben ---
    Dim blnMSGOK As Boolean: blnMSGOK = False
    If Dir(strAlterMSGPfad) <> "" Then
        ' Zielordner erstellen
        Dim strZielOrdner As String
        strZielOrdner = Left(strNeuerMSGPfad, InStrRev(strNeuerMSGPfad, "\"))
        ErstelleOrdner strZielOrdner

        ' Kopieren
        FileCopy strAlterMSGPfad, strNeuerMSGPfad
        blnMSGOK = True
    Else
        LogWarn "VerschiebeNachProjekt: MSG nicht gefunden: " & strAlterMSGPfad, "PROJEKT"
        blnMSGOK = True  ' Nicht abbrechen wegen fehlender Datei
    End If

    ' --- DB aktualisieren ---
    CurrentDb.Execute "UPDATE [" & TBL_EMAILS & "] SET MSGDateiPfad='" & _
        SQLSafe(strNeuerMSGPfad) & "' WHERE EmailID=" & lngEmailID

    ' --- Anhaenge verschieben ---
    Dim rsAnh As DAO.Recordset
    Set rsAnh = CurrentDb.OpenRecordset( _
        "SELECT AnhangID, DateiPfad FROM [" & TBL_EMAIL_ANHAENGE & "] " & _
        "WHERE EmailID=" & lngEmailID & " AND DateiPfad IS NOT NULL AND DateiPfad <> ''", _
        dbOpenSnapshot)

    Do While Not rsAnh.EOF
        Dim strAlterAnhPfad As String, strNeuerAnhPfad As String
        strAlterAnhPfad = Nz(rsAnh!DateiPfad, "")

        If strAlterAnhPfad <> "" And InStr(1, strAlterAnhPfad, "_Eingang", vbTextCompare) > 0 Then
            strNeuerAnhPfad = ErsetzePfadSegment(strAlterAnhPfad, "_Eingang", strProjektOrdner)

            If Dir(strAlterAnhPfad) <> "" Then
                Dim strAnhZielOrdner As String
                strAnhZielOrdner = Left(strNeuerAnhPfad, InStrRev(strNeuerAnhPfad, "\"))
                ErstelleOrdner strAnhZielOrdner
                FileCopy strAlterAnhPfad, strNeuerAnhPfad
            Else
                LogWarn "VerschiebeNachProjekt: Anhang nicht gefunden: " & strAlterAnhPfad, "PROJEKT"
            End If

            ' DB aktualisieren
            CurrentDb.Execute "UPDATE [" & TBL_EMAIL_ANHAENGE & "] SET DateiPfad='" & _
                SQLSafe(strNeuerAnhPfad) & "' WHERE AnhangID=" & rsAnh!AnhangID
        End If
        rsAnh.MoveNext
    Loop
    rsAnh.Close: Set rsAnh = Nothing

    ' --- Alte Dateien loeschen (erst nach DB-Update) ---
    On Error Resume Next
    If blnMSGOK And Dir(strAlterMSGPfad) <> "" And strAlterMSGPfad <> strNeuerMSGPfad Then
        Kill strAlterMSGPfad
    End If

    ' Alte Anhaenge loeschen (nochmal durchlaufen)
    Dim rsAnh2 As DAO.Recordset
    Set rsAnh2 = CurrentDb.OpenRecordset( _
        "SELECT DateiPfad FROM [" & TBL_EMAIL_ANHAENGE & "] " & _
        "WHERE EmailID=" & lngEmailID, dbOpenSnapshot)
    Do While Not rsAnh2.EOF
        ' Der alte Pfad: aus dem neuen zurueckrechnen
        Dim strAltRekon As String
        strAltRekon = ErsetzePfadSegment(Nz(rsAnh2!DateiPfad, ""), strProjektOrdner, "_Eingang")
        If strAltRekon <> Nz(rsAnh2!DateiPfad, "") And Dir(strAltRekon) <> "" Then
            Kill strAltRekon
        End If
        rsAnh2.MoveNext
    Loop
    rsAnh2.Close: Set rsAnh2 = Nothing

    ' Alten Ordner loeschen wenn leer
    Dim strAlterOrdner As String
    strAlterOrdner = Left(strAlterMSGPfad, InStrRev(strAlterMSGPfad, "\"))
    If Dir(strAlterOrdner & "*.*") = "" Then
        RmDir strAlterOrdner
    End If
    On Error GoTo 0

    VerschiebeNachProjekt = True
    LogInfo "Email " & lngEmailID & " nach Projekt-Ordner verschoben: " & strProjektOrdner, "PROJEKT"
    Exit Function

ErrHandler:
    LogError "VerschiebeNachProjekt: " & Err.Number & " - " & Err.Description, "PROJEKT"
    VerschiebeNachProjekt = False
End Function


' ===================================================================
' SECTION: HILFSFUNKTIONEN
' ===================================================================

' Ersetzt ein Pfadsegment (z.B. "_Eingang" durch "FLIWAS")
' Arbeitet case-insensitive auf dem ersten Vorkommen des Segments
Private Function ErsetzePfadSegment(ByVal strPfad As String, _
                                     ByVal strAlt As String, _
                                     ByVal strNeu As String) As String
    Dim lngPos As Long
    lngPos = InStr(1, strPfad, "\" & strAlt & "\", vbTextCompare)
    If lngPos > 0 Then
        ErsetzePfadSegment = Left(strPfad, lngPos) & strNeu & _
            Mid(strPfad, lngPos + Len(strAlt) + 1)
    Else
        ErsetzePfadSegment = strPfad
    End If
End Function


