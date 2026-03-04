Attribute VB_Name = "modDAO"
Option Compare Database
Option Explicit

' ===========================================================================
' modDAO - Datenzugriffsschicht (Data Access Object)
' ===========================================================================
' Zentraler Zugriff auf alle Tabellen der OutlookSync-Datenbank.
' Gruppiert nach Sektionen:
'   SYNC-LAUF  | KONTAKTE | ORDNER | THREADS | EMAILS |
'   EMAIL-CONTENT | EMPFAENGER | ANHAENGE | EMAIL-STATUS
'
' Alle Funktionen verwenden DAO (Built-in Access) und loggen Fehler.
' ===========================================================================


' ===================================================================
' SECTION: SYNC-LAUF
' ===================================================================

' Neuen Sync-Lauf starten -> gibt SyncLaufID zurueck
Public Function StarteSyncLauf(ByVal strOrdnerPfad As String, _
                                ByVal strProjekt As String, _
                                ByVal strPhase As String) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblSyncLauf", dbOpenDynaset)

    With rs
        .AddNew
        !StartZeit = Now
        !Status = "Gestartet"
        !OrdnerPfad = Left(Nz(strOrdnerPfad, ""), 255)
        !Projekt = Left(Nz(strProjekt, ""), 100)
        !Phase = Left(Nz(strPhase, ""), 100)
        .Update
        .Bookmark = .LastModified
        StarteSyncLauf = !SyncLaufID
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    LogInfo "SyncLauf gestartet: ID=" & StarteSyncLauf, "DAO"
    Exit Function

ErrHandler:
    LogVBAError "StarteSyncLauf"
    StarteSyncLauf = 0
End Function

' Sync-Lauf beenden mit Ergebnis-Zaehler
Public Sub BeendeSyncLauf(ByVal lngSyncLaufID As Long, _
                           ByVal strStatus As String, _
                           ByVal lngGelesen As Long, _
                           ByVal lngNeu As Long, _
                           ByVal lngDuplikate As Long, _
                           ByVal lngFehler As Long)
    On Error GoTo ErrHandler
    Dim db As DAO.Database

    Set db = CurrentDb
    db.Execute "UPDATE tblSyncLauf SET " & _
               "EndeZeit = #" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "#, " & _
               "Status = '" & SQLSafe(strStatus) & "', " & _
               "AnzahlGelesen = " & lngGelesen & ", " & _
               "AnzahlNeu = " & lngNeu & ", " & _
               "AnzahlDuplikate = " & lngDuplikate & ", " & _
               "AnzahlFehler = " & lngFehler & " " & _
               "WHERE SyncLaufID = " & lngSyncLaufID, dbFailOnError

    Set db = Nothing
    LogInfo "SyncLauf beendet: ID=" & lngSyncLaufID & " Status=" & strStatus & _
            " (Neu=" & lngNeu & " Dup=" & lngDuplikate & " Err=" & lngFehler & ")", "DAO"
    Exit Sub

ErrHandler:
    LogVBAError "BeendeSyncLauf"
End Sub


' ===================================================================
' SECTION: KONTAKTE
' ===================================================================

' Kontakt suchen oder neu anlegen -> gibt KontaktID zurueck
' v0.2: Nutzt ParseKontaktName fuer erweiterte Felder (Vorname, Nachname, etc.)
Public Function GetOderErstelleKontakt(ByVal strName As String, _
                                        ByVal strEmail As String, _
                                        Optional ByVal strEmailTyp As String = "SMTP") As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim lngID As Long
    Dim kontakt As Object

    ' Bereinigen
    Set kontakt = BereinigeKontaktInfo(strName, strEmail)
    strName = kontakt("Name")
    strEmail = kontakt("Email")

    Set db = CurrentDb

    ' 1. Zuerst nach Email suchen
    Set rs = db.OpenRecordset( _
        "SELECT KontaktID FROM tblKontakte WHERE Email='" & SQLSafe(strEmail) & "'", dbOpenSnapshot)
    If Not rs.EOF Then
        lngID = rs!KontaktID
        rs.Close: Set rs = Nothing: Set db = Nothing
        GetOderErstelleKontakt = lngID
        Exit Function
    End If
    rs.Close: Set rs = Nothing

    ' 2. Dann nach Anzeigename suchen
    Set rs = db.OpenRecordset( _
        "SELECT KontaktID FROM tblKontakte WHERE Anzeigename='" & SQLSafe(strName) & "'", dbOpenSnapshot)
    If Not rs.EOF Then
        lngID = rs!KontaktID
        rs.Close: Set rs = Nothing: Set db = Nothing
        GetOderErstelleKontakt = lngID
        Exit Function
    End If
    rs.Close: Set rs = Nothing

    ' 3. Neuen Kontakt anlegen mit erweitertem Namens-Parsing
    Dim kn As TypKontaktName
    kn = ParseKontaktName(strName, strEmail)

    ' Domain-basiertes Lernen: Institution von Geschwister-Kontakt uebernehmen
    If kn.Institution = "" Then
        Dim dictLern As Object
        Set dictLern = LerneVonDomain(strEmail)
        If dictLern.Exists("Institution") Then
            kn.Institution = dictLern("Institution")
        End If
    End If

    Set rs = db.OpenRecordset("tblKontakte", dbOpenDynaset)
    With rs
        .AddNew
        !Anzeigename = Left(IIf(kn.Anzeigename <> "", kn.Anzeigename, strName), 255)
        !Email = Left(strEmail, 255)
        !EmailTyp = Left(strEmailTyp, 10)
        !Vorname = Left(Nz(kn.Vorname, ""), 100)
        !Nachname = Left(Nz(kn.Nachname, ""), 100)
        !Titel = Left(Nz(kn.Titel, ""), 50)
        !Namenszusatz = Left(Nz(kn.Namenszusatz, ""), 100)
        !Institution = Left(Nz(kn.Institution, ""), 255)
        !Sortiername = Left(Nz(kn.Sortiername, ""), 255)
        !ErstelltAm = Now
        !AktualisiertAm = Now
        .Update
        .Bookmark = .LastModified
        lngID = !KontaktID
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    LogDebug "Neuer Kontakt: ID=" & lngID & " " & strEmail & _
             " [" & kn.Vorname & " " & kn.Nachname & "]", "DAO"
    GetOderErstelleKontakt = lngID
    Exit Function

ErrHandler:
    LogVBAError "GetOderErstelleKontakt"
    GetOderErstelleKontakt = 0
End Function


' ===================================================================
' SECTION: ORDNER
' ===================================================================

' Ordner in DB speichern oder aktualisieren -> gibt OrdnerID zurueck
Public Function SpeichereOrdner(ByVal strName As String, _
                                 ByVal strPfad As String, _
                                 Optional ByVal lngParentID As Long = 0, _
                                 Optional ByVal strPostfach As String = "", _
                                 Optional ByVal lngElemente As Long = 0) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim varID As Variant

    Set db = CurrentDb

    ' Pruefen ob Ordner bereits existiert (per Pfad)
    varID = DLookup("OrdnerID", "tblOutlookOrdner", "OrdnerPfad='" & SQLSafe(strPfad) & "'")
    If Not IsNull(varID) Then
        ' Aktualisieren: Elementanzahl + LetzterSync
        db.Execute "UPDATE tblOutlookOrdner SET " & _
                   "ElementAnzahl=" & lngElemente & ", " & _
                   "LetzterSync=#" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "# " & _
                   "WHERE OrdnerID=" & varID, dbFailOnError
        SpeichereOrdner = CLng(varID)
        Set db = Nothing
        Exit Function
    End If

    ' Neu anlegen
    Set rs = db.OpenRecordset("tblOutlookOrdner", dbOpenDynaset)
    With rs
        .AddNew
        !OrdnerName = Left(Nz(strName, ""), 255)
        !OrdnerPfad = Left(Nz(strPfad, ""), 255)
        !ParentID = lngParentID
        !PostfachName = Left(Nz(strPostfach, ""), 255)
        !ElementAnzahl = lngElemente
        !LetzterSync = Now
        .Update
        .Bookmark = .LastModified
        SpeichereOrdner = !OrdnerID
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    Exit Function

ErrHandler:
    LogVBAError "SpeichereOrdner"
    SpeichereOrdner = 0
End Function


' ===================================================================
' SECTION: THREADS
' ===================================================================

' Thread suchen oder neu anlegen -> gibt ThreadID zurueck
Public Function GetOderErstelleThread(ByVal strBetreff As String, _
                                       ByVal strMessageID As String, _
                                       Optional ByVal strInReplyTo As String = "", _
                                       Optional ByVal dtMailDatum As Date = 0, _
                                       Optional ByVal strAbsender As String = "") As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim lngID As Long
    Dim strIdent As String

    Set db = CurrentDb

    ' 1. Zuerst via In-Reply-To suchen (genaueste Zuordnung)
    If strInReplyTo <> "" Then
        Set rs = db.OpenRecordset( _
            "SELECT ThreadID FROM tblEmailThreads WHERE ThreadIdentifier='" & SQLSafe(strInReplyTo) & "'", _
            dbOpenSnapshot)
        If Not rs.EOF Then
            lngID = rs!ThreadID
            rs.Close: Set rs = Nothing
            Call AktualisiereThread(db, lngID, dtMailDatum)
            Set db = Nothing
            GetOderErstelleThread = lngID
            Exit Function
        End If
        rs.Close: Set rs = Nothing
    End If

    ' 2. Ueber bereinigten Betreff suchen
    strIdent = BereinigeBetreff(strBetreff)

    Set rs = db.OpenRecordset( _
        "SELECT ThreadID FROM tblEmailThreads WHERE ThreadIdentifier='" & SQLSafe(strIdent) & "'", _
        dbOpenSnapshot)
    If Not rs.EOF Then
        lngID = rs!ThreadID
        rs.Close: Set rs = Nothing
        Call AktualisiereThread(db, lngID, dtMailDatum)
        Set db = Nothing
        GetOderErstelleThread = lngID
        Exit Function
    End If
    rs.Close: Set rs = Nothing

    ' 3. Neuen Thread anlegen
    If dtMailDatum = 0 Then dtMailDatum = Now

    Set rs = db.OpenRecordset("tblEmailThreads", dbOpenDynaset)
    With rs
        .AddNew
        !ThreadBetreff = Left(strIdent, 255)
        ' Verwende InReplyTo als Identifier wenn vorhanden, sonst bereinigten Betreff
        If strInReplyTo <> "" Then
            !ThreadIdentifier = Left(strInReplyTo, 255)
        Else
            !ThreadIdentifier = Left(strIdent, 255)
        End If
        !Antwortanzahl = 1
        !ErsterAbsender = Left(Nz(strAbsender, ""), 255)
        !ErstesMailDatum = dtMailDatum
        !LetztesMailDatum = dtMailDatum
        !ErstelltAm = Now
        .Update
        .Bookmark = .LastModified
        lngID = !ThreadID
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    LogDebug "Neuer Thread: ID=" & lngID & " '" & Left(strIdent, 40) & "'", "DAO"
    GetOderErstelleThread = lngID
    Exit Function

ErrHandler:
    LogVBAError "GetOderErstelleThread"
    GetOderErstelleThread = 0
End Function

' Thread aktualisieren: Antwortanzahl erhoehen + LetztesMailDatum anpassen
Private Sub AktualisiereThread(db As DAO.Database, lngThreadID As Long, dtDatum As Date)
    On Error Resume Next
    Dim rs As DAO.Recordset

    If dtDatum = 0 Then dtDatum = Now

    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblEmailThreads WHERE ThreadID=" & lngThreadID, dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        rs!Antwortanzahl = rs!Antwortanzahl + 1
        If dtDatum > Nz(rs!LetztesMailDatum, #1/1/1900#) Then
            rs!LetztesMailDatum = dtDatum
        End If
        rs.Update
    End If
    rs.Close: Set rs = Nothing

    On Error GoTo 0
End Sub


' ===================================================================
' SECTION: EMAILS
' ===================================================================

' Pruefen ob Mail-Hash bereits in DB existiert
Public Function ExistiertMailHash(ByVal strHash As String) As Boolean
    On Error Resume Next
    ExistiertMailHash = (DCount("*", "tblEmails", "UniqueHash='" & SQLSafe(strHash) & "'") > 0)
    If Err.Number <> 0 Then ExistiertMailHash = False: Err.Clear
    On Error GoTo 0
End Function

' Email-Metadaten in Datenbank speichern -> gibt EmailID zurueck (0 bei Fehler)
Public Function SpeichereEmail(ByVal strEntryID As String, _
                                ByVal strHash As String, _
                                ByVal lngThreadID As Long, _
                                ByVal lngOrdnerID As Long, _
                                ByVal lngKontaktID As Long, _
                                ByVal lngSyncLaufID As Long, _
                                ByVal strBetreff As String, _
                                ByVal strAbsenderName As String, _
                                ByVal strAbsenderEmail As String, _
                                ByVal dtEmpfangen As Date, _
                                ByVal dtGesendet As Date, _
                                ByVal lngGroesse As Long, _
                                ByVal intWichtigkeit As Integer, _
                                ByVal blnGelesen As Boolean, _
                                ByVal blnHatAnhaenge As Boolean, _
                                ByVal intAnhangAnzahl As Integer, _
                                ByVal strMessageClass As String, _
                                ByVal strInternetMsgID As String) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblEmails", dbOpenDynaset)

    With rs
        .AddNew
        !OutlookEntryID = Left(Nz(strEntryID, ""), 255)
        !UniqueHash = Left(strHash, 64)
        !ThreadID = lngThreadID
        !OrdnerID = lngOrdnerID
        !KontaktID_Absender = lngKontaktID
        !SyncLaufID = lngSyncLaufID
        !Betreff = Left(Nz(strBetreff, ""), 255)
        !BetreffBereinigt = Left(BereinigeBetreff(Nz(strBetreff, "")), 255)
        !AbsenderName = Left(Nz(strAbsenderName, ""), 255)
        !AbsenderEmail = Left(Nz(strAbsenderEmail, ""), 255)
        !EmpfangenAm = dtEmpfangen
        !GesendetAm = dtGesendet
        !Groesse = lngGroesse
        !Wichtigkeit = intWichtigkeit
        !Gelesen = blnGelesen
        !HatAnhaenge = blnHatAnhaenge
        !AnhangAnzahl = intAnhangAnzahl
        !MessageClass = Left(Nz(strMessageClass, ""), 50)
        !InternetMessageID = Left(Nz(strInternetMsgID, ""), 255)
        !Status = "Neu"
        !ErstelltAm = Now
        .Update
        .Bookmark = .LastModified
        SpeichereEmail = !EmailID
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    Exit Function

ErrHandler:
    LogVBAError "SpeichereEmail"
    SpeichereEmail = 0
End Function


' ===================================================================
' SECTION: EMAIL-CONTENT
' ===================================================================

' Email-Inhalt (HTML + Plaintext) in separate Tabelle speichern
Public Sub SpeichereEmailContent(ByVal lngEmailID As Long, _
                                  ByVal strHTML As String, _
                                  ByVal strPlain As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim blnHatHTML As Boolean

    blnHatHTML = (Len(Trim(Nz(strHTML, ""))) > 0)
    If Not blnHatHTML Then strHTML = ""
    If Trim(Nz(strPlain, "")) = "" Then strPlain = ""

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblEmailContent", dbOpenDynaset)

    With rs
        .AddNew
        !EmailID = lngEmailID
        !HTMLBody = strHTML
        !PlainTextBody = strPlain
        !HatHTML = blnHatHTML
        !GroesseHTML = Len(strHTML)
        !GroesseText = Len(strPlain)
        .Update
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    Exit Sub

ErrHandler:
    LogVBAError "SpeichereEmailContent"
End Sub


' ===================================================================
' SECTION: EMPFAENGER
' ===================================================================

' Einzelnen Empfaenger speichern
Public Sub SpeichereEmpfaenger(ByVal lngEmailID As Long, _
                                ByVal lngKontaktID As Long, _
                                ByVal strTyp As String, _
                                ByVal strName As String, _
                                ByVal strEmail As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblEmailEmpfaenger", dbOpenDynaset)

    With rs
        .AddNew
        !EmailID = lngEmailID
        !KontaktID = lngKontaktID
        !Typ = Left(strTyp, 5)
        !Anzeigename = Left(Nz(strName, ""), 255)
        !Email = Left(Nz(strEmail, ""), 255)
        .Update
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    Exit Sub

ErrHandler:
    LogVBAError "SpeichereEmpfaenger"
End Sub


' ===================================================================
' SECTION: ANHAENGE
' ===================================================================

' Anhang-Metadaten speichern -> gibt AnhangID zurueck
Public Function SpeichereAnhangMetadaten(ByVal lngEmailID As Long, _
                                          ByVal strDateiname As String, _
                                          ByVal lngGroesse As Long, _
                                          ByVal strMimeType As String, _
                                          ByVal intAnhangTyp As Integer, _
                                          ByVal blnVersteckt As Boolean, _
                                          Optional ByVal strDateiPfad As String = "") As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblEmailAnhaenge", dbOpenDynaset)

    With rs
        .AddNew
        !EmailID = lngEmailID
        !Dateiname = Left(Nz(strDateiname, ""), 255)
        !DateinameBereinigt = Left(BereinigeDateiname(Nz(strDateiname, "")), 255)
        !Erweiterung = Left(HoleEndung(Nz(strDateiname, "")), 20)
        !Groesse = lngGroesse
        !MimeType = Left(Nz(strMimeType, ""), 100)
        !AnhangTyp = intAnhangTyp
        !IstVersteckt = blnVersteckt
        !IstGespeichert = (strDateiPfad <> "")
        !DateiPfad = Left(Nz(strDateiPfad, ""), 255)
        !ErstelltAm = Now
        .Update
        .Bookmark = .LastModified
        SpeichereAnhangMetadaten = !AnhangID
    End With

    rs.Close: Set rs = Nothing: Set db = Nothing
    Exit Function

ErrHandler:
    LogVBAError "SpeichereAnhangMetadaten"
    SpeichereAnhangMetadaten = 0
End Function

' Anhang-Dateipfad nachtraeglich aktualisieren (nach erfolgreichem SaveAsFile)
Public Sub AktualisiereAnhangPfad(ByVal lngAnhangID As Long, ByVal strDateiPfad As String)
    On Error GoTo ErrHandler
    CurrentDb.Execute "UPDATE tblEmailAnhaenge SET " & _
                      "DateiPfad='" & SQLSafe(strDateiPfad) & "', " & _
                      "IstGespeichert=True " & _
                      "WHERE AnhangID=" & lngAnhangID, dbFailOnError
    Exit Sub
ErrHandler:
    LogVBAError "AktualisiereAnhangPfad"
End Sub


' ===================================================================
' SECTION: EMAIL-STATUS
' ===================================================================

' Status-Eintrag in Historie-Tabelle + Haupt-Status aktualisieren
Public Sub SpeichereEmailStatus(ByVal lngEmailID As Long, _
                                 ByVal strStatus As String, _
                                 Optional ByVal strVon As String = "System", _
                                 Optional ByVal strBemerkung As String = "")
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset

    Set db = CurrentDb

    ' 1. Status in History-Tabelle eintragen
    Set rs = db.OpenRecordset("tblEmailStatus", dbOpenDynaset)
    With rs
        .AddNew
        !EmailID = lngEmailID
        !Status = Left(strStatus, 50)
        !GeaendertVon = Left(strVon, 100)
        !Bemerkung = Left(strBemerkung, 255)
        !GeaendertAm = Now
        .Update
    End With
    rs.Close: Set rs = Nothing

    ' 2. Status in Haupt-Tabelle aktualisieren
    db.Execute "UPDATE tblEmails SET Status='" & SQLSafe(strStatus) & _
               "' WHERE EmailID=" & lngEmailID

    Set db = Nothing
    Exit Sub

ErrHandler:
    LogVBAError "SpeichereEmailStatus"
End Sub

' MSG-Dateipfad in tblEmails nachtraeglich setzen
Public Sub SetzeEmailMSGPfad(ByVal lngEmailID As Long, ByVal strPfad As String)
    On Error GoTo ErrHandler
    CurrentDb.Execute "UPDATE tblEmails SET MSGDateiPfad='" & SQLSafe(strPfad) & _
                      "' WHERE EmailID=" & lngEmailID
    Exit Sub
ErrHandler:
    LogVBAError "SetzeEmailMSGPfad"
End Sub
