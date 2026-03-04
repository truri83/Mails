Attribute VB_Name = "modSync"
Option Compare Database
Option Explicit

' ===========================================================================
' modSync - Synchronisations-Orchestrierung
' ===========================================================================
' Hauptmodul fuer die Outlook-zu-Access-Synchronisation.
' Liest Mails aus Outlook-Ordnern und speichert alles in der Datenbank.
'
' Oeffentliche Routinen:
'   SyncPosteingang [Projekt, Phase, MaxMails]   - Posteingang synchronisieren
'   SyncOrdner "Pfad" [, Projekt, Phase, Max]    - Beliebigen Ordner synchronisieren
'   SyncOrdnerStruktur [Tiefe]                   - Ordnerstruktur in DB einlesen
'
' Aufruf-Beispiele (Direktbereich STRG+G):
'   SyncPosteingang "FLIWAS", "Test", 50
'   SyncOrdner "Torsten.Kugler@rps.bwl.de\Posteingang\FLIWAS", "FLIWAS", "Prod"
'   SyncOrdnerStruktur 3
' ===========================================================================


' ---------------------------------------------------------------------------
' POSTEINGANG SYNCHRONISIEREN (Convenience-Wrapper)
' ---------------------------------------------------------------------------
Public Sub SyncPosteingang(Optional ByVal strProjekt As String = "Standard", _
                            Optional ByVal strPhase As String = "Standard", _
                            Optional ByVal lngMaxMails As Long = 0, _
                            Optional ByVal blnSubfolder As Boolean = False)

    If Not ConnectRDO() Then
        LogError "SyncPosteingang: Keine RDO-Verbindung", "SYNC"
        Exit Sub
    End If

    Dim objInbox As Object
    Set objInbox = g_objRDO.GetDefaultFolder(olFolderInbox)

    If objInbox Is Nothing Then
        LogError "SyncPosteingang: Posteingang nicht erreichbar", "SYNC"
        Exit Sub
    End If

    ' Maximale Mailanzahl aus Config lesen wenn nicht angegeben
    If lngMaxMails <= 0 Then
        lngMaxMails = CLng(LeseConfig("MaxMailsProSync", "500"))
    End If

    Call SyncFolder(objInbox, strProjekt, strPhase, lngMaxMails, blnSubfolder)

    Set objInbox = Nothing
End Sub


' ---------------------------------------------------------------------------
' ORDNER PER PFAD SYNCHRONISIEREN
' ---------------------------------------------------------------------------
Public Sub SyncOrdner(ByVal strOrdnerPfad As String, _
                       Optional ByVal strProjekt As String = "Standard", _
                       Optional ByVal strPhase As String = "Standard", _
                       Optional ByVal lngMaxMails As Long = 0, _
                       Optional ByVal blnSubfolder As Boolean = False)

    Dim objFolder As Object
    Set objFolder = OeffneOrdner(strOrdnerPfad)

    If objFolder Is Nothing Then
        LogError "SyncOrdner: Ordner '" & strOrdnerPfad & "' nicht gefunden", "SYNC"
        Exit Sub
    End If

    If lngMaxMails <= 0 Then
        lngMaxMails = CLng(LeseConfig("MaxMailsProSync", "500"))
    End If

    Call SyncFolder(objFolder, strProjekt, strPhase, lngMaxMails, blnSubfolder)

    Set objFolder = Nothing
End Sub


' ---------------------------------------------------------------------------
' KERN-ROUTINE: Ordner-Objekt synchronisieren
' v0.2: Optional mit rekursiver Subfolder-Verarbeitung
' ---------------------------------------------------------------------------
Public Sub SyncFolder(objFolder As Object, _
                       ByVal strProjekt As String, _
                       ByVal strPhase As String, _
                       ByVal lngMaxMails As Long, _
                       Optional ByVal blnSubfolder As Boolean = False)

    On Error GoTo ErrHandler

    Dim lngSyncLaufID   As Long
    Dim lngOrdnerID     As Long
    Dim lngGesamt       As Long
    Dim lngVerarbeitet  As Long
    Dim lngNeu          As Long
    Dim lngDuplikate    As Long
    Dim lngFehler       As Long
    Dim lngMailCount    As Long
    Dim objItem         As Object
    Dim i               As Long
    Dim lngResult       As Long

    ' --- Initialisierung ---
    InitGlobals
    g_blnAbbrechen = False
    g_dtSyncStart = Timer

    If Nz(strProjekt, "") = "" Then strProjekt = "Standard"
    If Nz(strPhase, "") = "" Then strPhase = "Standard"

    lngGesamt = objFolder.Items.Count
    lngVerarbeitet = 0: lngNeu = 0: lngDuplikate = 0: lngFehler = 0: lngMailCount = 0

    Debug.Print String(70, "=")
    Debug.Print "=== SYNC START: " & objFolder.Name & " ==="
    Debug.Print "    Ordner     : " & objFolder.FolderPath
    Debug.Print "    Elemente   : " & lngGesamt
    Debug.Print "    Max. Mails : " & lngMaxMails
    Debug.Print "    Projekt    : " & strProjekt
    Debug.Print "    Phase      : " & strPhase
    Debug.Print "    " & Now()
    Debug.Print String(70, "=")

    ' Sync-Lauf in DB starten
    lngSyncLaufID = StarteSyncLauf(objFolder.FolderPath, strProjekt, strPhase)

    ' Ordner in DB registrieren
    lngOrdnerID = SpeichereOrdner(objFolder.Name, objFolder.FolderPath, 0, "", lngGesamt)

    ' Config-Flags lesen
    Dim blnAnhaenge As Boolean
    Dim blnMSG      As Boolean
    Dim blnFilter   As Boolean
    blnAnhaenge = (LeseConfig("AnhaengeExtrahieren", "1") = "1")
    blnMSG = (LeseConfig("MSGExportieren", "1") = "1")
    blnFilter = (LeseConfig("SignaturBilderFiltern", "1") = "1")

    ' Export-Basispfad
    Dim strExportBase As String
    strExportBase = NormalisierePfad(LeseConfig("ExportBasisPfad", _
                    Environ("USERPROFILE") & "\OutlookSync\"))

    ' --- Iteration ueber Mails ---
    For i = 1 To lngGesamt

        ' Abbruch-Pruefung
        If g_blnAbbrechen Then
            LogWarn "Sync abgebrochen nach " & lngMailCount & " Mails", "SYNC"
            Exit For
        End If

        ' Mail-Objekt holen
        On Error Resume Next
        Set objItem = objFolder.Items(i)
        If Err.Number <> 0 Or objItem Is Nothing Then
            Err.Clear
            lngFehler = lngFehler + 1
            On Error GoTo ErrHandler
            GoTo NaechsteMail
        End If
        On Error GoTo ErrHandler

        ' Nur echte E-Mails (IPM.Note)
        If Left(objItem.MessageClass, 8) <> "IPM.Note" Then GoTo NaechsteMail

        lngMailCount = lngMailCount + 1

        ' Fortschritt ausgeben (alle 10 Mails + erste Mail)
        If lngMailCount Mod 10 = 0 Or lngMailCount = 1 Then
            Debug.Print "  [" & Format(lngMailCount, "000") & "/" & _
                        Format(lngGesamt, "000") & "] Verarbeite..."
        End If

        ' Einzelne Mail verarbeiten
        lngResult = VerarbeiteEinzelmailRDO(objItem, lngSyncLaufID, lngOrdnerID, _
                                             strProjekt, strPhase, strExportBase, _
                                             blnAnhaenge, blnMSG, blnFilter)

        Select Case lngResult
            Case Is > 0: lngNeu = lngNeu + 1
            Case 0:      lngDuplikate = lngDuplikate + 1
            Case -1:     lngFehler = lngFehler + 1
        End Select

        lngVerarbeitet = lngVerarbeitet + 1

        ' Maximum erreicht?
        If lngMailCount >= lngMaxMails Then
            LogInfo "Maximum erreicht (" & lngMaxMails & " Mails)", "SYNC"
            Exit For
        End If

        DoEvents

NaechsteMail:
    Next i

    ' --- Abschluss ---
    Dim strStatus As String
    If g_blnAbbrechen Then
        strStatus = "Abgebrochen"
    Else
        strStatus = "Abgeschlossen"
    End If

    Call BeendeSyncLauf(lngSyncLaufID, strStatus, lngVerarbeitet, lngNeu, lngDuplikate, lngFehler)

    Debug.Print String(70, "-")
    Debug.Print "=== SYNC ERGEBNIS ==="
    Debug.Print "  Verarbeitet : " & lngVerarbeitet
    Debug.Print "  Neu         : " & lngNeu
    Debug.Print "  Duplikate   : " & lngDuplikate
    Debug.Print "  Fehler      : " & lngFehler
    Debug.Print "  Dauer       : " & Format((Timer - g_dtSyncStart) / 60, "0.0") & " min"
    Debug.Print "  Status      : " & strStatus
    Debug.Print String(70, "=")

    ' --- Subfolder rekursiv verarbeiten (v0.2) ---
    If blnSubfolder And Not g_blnAbbrechen Then
        Call SyncSubfolder(objFolder, strProjekt, strPhase, lngMaxMails)
    End If

    Set objItem = Nothing
    Exit Sub

ErrHandler:
    LogVBAError "SyncFolder"
    If lngSyncLaufID > 0 Then
        Call BeendeSyncLauf(lngSyncLaufID, "Fehler", lngVerarbeitet, lngNeu, lngDuplikate, lngFehler + 1)
    End If
End Sub


' ---------------------------------------------------------------------------
' EINZELNE MAIL VERARBEITEN (RDO)
' Rueckgabe: >0 = EmailID (neu gespeichert), 0 = Duplikat, -1 = Fehler
' ---------------------------------------------------------------------------
Private Function VerarbeiteEinzelmailRDO(objRDOMail As Object, _
                                          ByVal lngSyncLaufID As Long, _
                                          ByVal lngOrdnerID As Long, _
                                          ByVal strProjekt As String, _
                                          ByVal strPhase As String, _
                                          ByVal strExportBase As String, _
                                          ByVal blnAnhaenge As Boolean, _
                                          ByVal blnMSG As Boolean, _
                                          ByVal blnFilter As Boolean) As Long
    On Error GoTo ErrHandler

    Dim strHash         As String
    Dim strEntryID      As String
    Dim strBetreff      As String
    Dim strAbsenderName As String
    Dim strAbsenderEmail As String
    Dim strEmpfaenger   As String
    Dim dtEmpfangen     As Date
    Dim dtGesendet      As Date
    Dim lngEmailID      As Long
    Dim lngKontaktID    As Long
    Dim lngThreadID     As Long
    Dim strMessageID    As String
    Dim strInReplyTo    As String

    ' --- Basisdaten lesen ---
    strEntryID = Nz(objRDOMail.EntryID, "")
    strBetreff = Nz(objRDOMail.Subject, "")
    strAbsenderName = Nz(objRDOMail.SenderName, "")
    dtEmpfangen = Nz(objRDOMail.ReceivedTime, Now)
    dtGesendet = Nz(objRDOMail.SentOn, dtEmpfangen)

    ' Absender-SMTP aufloesen
    strAbsenderEmail = GetAbsenderSMTP(objRDOMail)

    ' Empfaenger-Feld fuer Hash
    On Error Resume Next
    strEmpfaenger = objRDOMail.Fields(PR_DISPLAY_TO)
    If Err.Number <> 0 Then strEmpfaenger = "": Err.Clear
    On Error GoTo ErrHandler

    ' --- Duplikat-Pruefung ---
    strHash = GeneriereMailHash(strBetreff, strAbsenderEmail, strEmpfaenger, dtEmpfangen)
    If ExistiertMailHash(strHash) Then
        LogTrace "Duplikat: " & Left(strBetreff, 40), "SYNC"
        VerarbeiteEinzelmailRDO = 0
        Exit Function
    End If

    ' --- Internet-Message-ID & In-Reply-To ---
    On Error Resume Next
    strMessageID = objRDOMail.Fields(PR_INTERNET_MESSAGE_ID)
    If Err.Number <> 0 Then strMessageID = "": Err.Clear

    Dim strHeaders As String
    strHeaders = objRDOMail.Fields(PR_TRANSPORT_MESSAGE_HEADERS)
    If Err.Number <> 0 Then strHeaders = "": Err.Clear
    On Error GoTo ErrHandler

    strInReplyTo = ParseHeaderField(strHeaders, "In-Reply-To")

    ' --- Kontakt + Thread ermitteln/erstellen ---
    Dim strEmailTyp As String
    On Error Resume Next
    strEmailTyp = objRDOMail.SenderEmailType
    If Err.Number <> 0 Then strEmailTyp = "SMTP": Err.Clear
    On Error GoTo ErrHandler

    lngKontaktID = GetOderErstelleKontakt(strAbsenderName, strAbsenderEmail, _
                                           IIf(strEmailTyp = "EX", "EX", "SMTP"))
    lngThreadID = GetOderErstelleThread(strBetreff, strMessageID, strInReplyTo, _
                                         dtEmpfangen, strAbsenderName)

    ' --- Weitere Metadaten ---
    Dim blnUnread   As Boolean
    Dim lngSize     As Long
    Dim intImp      As Integer
    Dim intAttCount As Integer

    On Error Resume Next
    blnUnread = objRDOMail.UnRead
    lngSize = objRDOMail.Size
    intImp = objRDOMail.Importance
    intAttCount = objRDOMail.Attachments.Count
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrHandler

    ' --- Email in DB speichern ---
    lngEmailID = SpeichereEmail( _
        strEntryID, strHash, lngThreadID, lngOrdnerID, lngKontaktID, lngSyncLaufID, _
        strBetreff, strAbsenderName, strAbsenderEmail, dtEmpfangen, dtGesendet, _
        lngSize, intImp, Not blnUnread, (intAttCount > 0), intAttCount, _
        objRDOMail.MessageClass, strMessageID)

    If lngEmailID = 0 Then
        VerarbeiteEinzelmailRDO = -1
        Exit Function
    End If

    ' --- Content speichern ---
    On Error Resume Next
    Dim strHTML As String, strPlain As String
    strHTML = Nz(objRDOMail.HTMLBody, "")
    strPlain = Nz(objRDOMail.Body, "")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrHandler

    Call SpeichereEmailContent(lngEmailID, strHTML, strPlain)

    ' --- Empfaenger speichern ---
    Call VerarbeiteEmpfaenger(objRDOMail, lngEmailID)

    ' --- Anhaenge verarbeiten ---
    If blnAnhaenge And intAttCount > 0 Then
        Call VerarbeiteAnhaenge(objRDOMail, lngEmailID, strProjekt, strPhase, _
                                strExportBase, blnFilter)
    End If

    ' --- MSG exportieren ---
    If blnMSG Then
        Call ExportiereMSG(objRDOMail, lngEmailID, strProjekt, strPhase, strExportBase)
    End If

    ' --- Status setzen ---
    Call SpeichereEmailStatus(lngEmailID, "Verarbeitet", "Sync", "Import via SyncFolder")

    VerarbeiteEinzelmailRDO = lngEmailID
    Exit Function

ErrHandler:
    LogVBAError "VerarbeiteEinzelmailRDO [" & Left(Nz(strBetreff, "?"), 30) & "]"
    VerarbeiteEinzelmailRDO = -1
End Function


' ---------------------------------------------------------------------------
' EMPFAENGER EINER MAIL VERARBEITEN
' v0.2: Aktualisiert Kontakt-E-Mail wenn bisherige ungueltig
' ---------------------------------------------------------------------------
Private Sub VerarbeiteEmpfaenger(objRDOMail As Object, ByVal lngEmailID As Long)
    On Error GoTo ErrHandler
    Dim objRec      As Object
    Dim strName     As String
    Dim strEmail    As String
    Dim strTyp      As String
    Dim lngKontaktID As Long

    For Each objRec In objRDOMail.Recipients
        On Error Resume Next
        strName = objRec.Name
        strEmail = GetSMTPFromRecipient(objRec)
        If Err.Number <> 0 Then Err.Clear: strName = "": strEmail = ""
        On Error GoTo ErrHandler

        Select Case objRec.Type
            Case 1: strTyp = "To"
            Case 2: strTyp = "CC"
            Case 3: strTyp = "BCC"
            Case Else: strTyp = "?"
        End Select

        ' Kontakt ermitteln/erstellen
        If IstGueltigeEmail(strEmail) Then
            lngKontaktID = GetOderErstelleKontakt(strName, strEmail)

            ' v0.2: E-Mail aktualisieren wenn Kontakt noch "Unbekannt"-Adresse hat
            If lngKontaktID > 0 Then
                Call AktualisiereKontaktEmail(lngKontaktID, strEmail)
            End If
        Else
            lngKontaktID = 0
        End If

        Call SpeichereEmpfaenger(lngEmailID, lngKontaktID, strTyp, strName, strEmail)
    Next objRec

    Exit Sub

ErrHandler:
    LogVBAError "VerarbeiteEmpfaenger"
End Sub


' ---------------------------------------------------------------------------
' ANHAENGE EINER MAIL VERARBEITEN
' ---------------------------------------------------------------------------
Private Sub VerarbeiteAnhaenge(objRDOMail As Object, _
                                ByVal lngEmailID As Long, _
                                ByVal strProjekt As String, _
                                ByVal strPhase As String, _
                                ByVal strExportBase As String, _
                                ByVal blnFilter As Boolean)
    On Error GoTo ErrHandler
    Dim objAtt      As Object
    Dim a           As Integer
    Dim strZielDir  As String
    Dim strDateiPfad As String
    Dim strEndung   As String
    Dim lngAnhangID As Long

    ' Zielordner: {Base}\{Projekt}\{Phase}\Anhaenge\EmailID_{id}\
    strZielDir = NormalisierePfad(strExportBase & strProjekt & "\" & strPhase & _
                                  "\Anhaenge\EmailID_" & lngEmailID & "\")

    For a = 1 To objRDOMail.Attachments.Count
        Set objAtt = objRDOMail.Attachments(a)

        ' Filter: Signatur-Bilder ueberspringen (nur Metadaten speichern)
        If blnFilter And (objAtt.Hidden Or objAtt.Type <> ATT_BY_VALUE) Then
            Call SpeichereAnhangMetadaten(lngEmailID, Nz(objAtt.FileName, ""), _
                                          objAtt.Size, Nz(objAtt.MimeTag, ""), _
                                          objAtt.Type, objAtt.Hidden, "")
            GoTo NaechsterAnhang
        End If

        ' Ordner erstellen
        ErstelleOrdner strZielDir

        ' Dateipfad generieren
        strEndung = HoleEndung(Nz(objAtt.FileName, ""))
        If strEndung = "" Then strEndung = "bin"
        strDateiPfad = EindeutigerDateipfad(strZielDir, Nz(objAtt.FileName, "Anhang"), strEndung)

        ' Metadaten speichern (noch ohne Dateipfad)
        lngAnhangID = SpeichereAnhangMetadaten(lngEmailID, Nz(objAtt.FileName, ""), _
                                                 objAtt.Size, Nz(objAtt.MimeTag, ""), _
                                                 objAtt.Type, objAtt.Hidden)

        ' Datei auf Festplatte speichern
        On Error Resume Next
        objAtt.SaveAsFile strDateiPfad
        If Err.Number = 0 Then
            Call AktualisiereAnhangPfad(lngAnhangID, strDateiPfad)
            LogDebug "Anhang: " & objAtt.FileName & " -> " & strDateiPfad, "SYNC"
        Else
            LogWarn "Anhang fehlgeschlagen: " & Nz(objAtt.FileName, "?") & " - " & Err.Description, "SYNC"
            Err.Clear
        End If
        On Error GoTo ErrHandler

NaechsterAnhang:
    Next a

    Exit Sub

ErrHandler:
    LogVBAError "VerarbeiteAnhaenge"
End Sub


' ---------------------------------------------------------------------------
' MSG-DATEI EXPORTIEREN
' ---------------------------------------------------------------------------
Private Sub ExportiereMSG(objRDOMail As Object, _
                           ByVal lngEmailID As Long, _
                           ByVal strProjekt As String, _
                           ByVal strPhase As String, _
                           ByVal strExportBase As String)
    On Error GoTo ErrHandler

    Dim strZielDir  As String
    Dim strDateiname As String
    Dim strPfad     As String

    ' Zielordner: {Base}\{Projekt}\{Phase}\MSG\
    strZielDir = NormalisierePfad(strExportBase & strProjekt & "\" & strPhase & "\MSG\")
    ErstelleOrdner strZielDir

    ' Dateiname: yyyymmdd_hhnn_Betreff.msg
    strDateiname = Format(objRDOMail.ReceivedTime, "yyyymmdd_hhnn") & "_" & _
                   BereinigeDateiname(Nz(objRDOMail.Subject, "Kein_Betreff"), 50)
    strPfad = EindeutigerDateipfad(strZielDir, strDateiname, "msg")

    On Error Resume Next
    objRDOMail.SaveAs strPfad
    If Err.Number = 0 Then
        Call SetzeEmailMSGPfad(lngEmailID, strPfad)
        LogDebug "MSG: " & strPfad, "SYNC"
    Else
        LogWarn "MSG-Export fehlgeschlagen: " & Err.Description, "SYNC"
        Err.Clear
    End If
    On Error GoTo ErrHandler

    Exit Sub

ErrHandler:
    LogVBAError "ExportiereMSG"
End Sub


' ---------------------------------------------------------------------------
' HEADER-FELD AUS INTERNET-HEADERN PARSEN
' ---------------------------------------------------------------------------
Private Function ParseHeaderField(ByVal strHeaders As String, _
                                   ByVal strFieldName As String) As String
    On Error Resume Next
    Dim pos As Long, posEnd As Long
    Dim strSearch As String

    ParseHeaderField = ""
    If Len(strHeaders) = 0 Then Exit Function

    ' Im Header suchen (Format: "FieldName: value\r\n")
    strSearch = vbCrLf & strFieldName & ": "
    pos = InStr(1, strHeaders, strSearch, vbTextCompare)

    ' Auch am Anfang der Headers pruefen
    If pos = 0 Then
        strSearch = strFieldName & ": "
        If LCase(Left(strHeaders, Len(strSearch))) = LCase(strSearch) Then
            pos = 1
        Else
            Exit Function
        End If
    Else
        pos = pos + 2  ' vbCrLf ueberspringen
    End If

    pos = pos + Len(strFieldName) + 2  ' "FieldName: " ueberspringen
    posEnd = InStr(pos, strHeaders, vbCrLf)
    If posEnd = 0 Then posEnd = Len(strHeaders) + 1

    ParseHeaderField = Trim(Mid(strHeaders, pos, posEnd - pos))
End Function


' ===========================================================================
' ORDNERSTRUKTUR IN DB EINLESEN
' ===========================================================================

' Alle Postfaecher/Stores + Ordner rekursiv in tblOutlookOrdner speichern
Public Sub SyncOrdnerStruktur(Optional ByVal intMaxTiefe As Integer = 5)
    On Error GoTo ErrHandler

    If Not ConnectRDO() Then
        LogError "SyncOrdnerStruktur: Keine RDO-Verbindung", "SYNC"
        Exit Sub
    End If

    Debug.Print String(70, "=")
    Debug.Print "=== Ordnerstruktur einlesen (Tiefe=" & intMaxTiefe & ") ==="
    Debug.Print String(70, "=")

    Dim objStore As Object
    Dim objRoot As Object
    Dim lngRootID As Long

    For Each objStore In g_objRDO.Stores
        Set objRoot = objStore.RootFolder
        Debug.Print "[STORE] " & objStore.DisplayName

        lngRootID = SpeichereOrdner(objRoot.Name, objRoot.FolderPath, 0, _
                                     objStore.DisplayName, 0)

        Call SyncOrdnerRekursiv(objRoot, lngRootID, objStore.DisplayName, _
                                 1, intMaxTiefe)
    Next objStore

    Debug.Print String(70, "=")
    Debug.Print "=== Ordnerstruktur eingelesen ==="
    Debug.Print String(70, "=")
    Exit Sub

ErrHandler:
    LogVBAError "SyncOrdnerStruktur"
End Sub

' Rekursiver Ordner-Scan
Private Sub SyncOrdnerRekursiv(objParent As Object, _
                                ByVal lngParentID As Long, _
                                ByVal strPostfach As String, _
                                ByVal intLevel As Integer, _
                                ByVal intMaxDepth As Integer)
    If intLevel > intMaxDepth Then Exit Sub

    On Error Resume Next
    Dim objSub      As Object
    Dim lngElements As Long
    Dim lngID       As Long

    For Each objSub In objParent.Folders
        lngElements = objSub.Items.Count
        If Err.Number <> 0 Then lngElements = 0: Err.Clear

        lngID = SpeichereOrdner(objSub.Name, objSub.FolderPath, lngParentID, _
                                 strPostfach, lngElements)
        Debug.Print String(intLevel * 2, " ") & "|- " & objSub.Name & " (" & lngElements & ")"

        If intLevel < intMaxDepth Then
            Call SyncOrdnerRekursiv(objSub, lngID, strPostfach, intLevel + 1, intMaxDepth)
        End If
    Next objSub
    On Error GoTo 0
End Sub


' ===========================================================================
' SUBFOLDER EINER MAIL-SYNC REKURSIV VERARBEITEN (v0.2)
' ===========================================================================

' Iteriert rekursiv ueber Unterordner und ruft SyncFolder fuer jeden auf.
' Stoppt bei g_blnAbbrechen = True.
Private Sub SyncSubfolder(objParent As Object, _
                           ByVal strProjekt As String, _
                           ByVal strPhase As String, _
                           ByVal lngMaxMails As Long, _
                           Optional ByVal intLevel As Integer = 1, _
                           Optional ByVal intMaxDepth As Integer = 5)
    If intLevel > intMaxDepth Then Exit Sub
    If g_blnAbbrechen Then Exit Sub

    On Error Resume Next
    Dim objSub As Object

    For Each objSub In objParent.Folders
        If g_blnAbbrechen Then Exit For
        If Err.Number <> 0 Then Err.Clear: GoTo NaechsterSub

        ' Nur Ordner mit E-Mails (nicht Kalender, Kontakte etc.)
        If objSub.DefaultItemType = olMail Then
            Debug.Print String(intLevel * 2, " ") & "|- SYNC: " & objSub.Name & _
                        " (" & objSub.Items.Count & " Elemente)"

            On Error GoTo 0
            Call SyncFolder(objSub, strProjekt, strPhase, lngMaxMails, False)
            On Error Resume Next

            ' Weitere Tiefe
            If intLevel < intMaxDepth Then
                Call SyncSubfolder(objSub, strProjekt, strPhase, lngMaxMails, _
                                    intLevel + 1, intMaxDepth)
            End If
        End If

NaechsterSub:
    Next objSub

    On Error GoTo 0
End Sub


' ===========================================================================
' GANZES POSTFACH SYNCHRONISIEREN (v0.2)
' ===========================================================================

' Synchronisiert alle Ordner eines Postfachs inkl. Unterordner.
' Aufruf: SyncPostfach "Torsten.Kugler@rps.bwl.de", "FLIWAS", "Test"
Public Sub SyncPostfach(ByVal strPostfachName As String, _
                         Optional ByVal strProjekt As String = "Standard", _
                         Optional ByVal strPhase As String = "Standard", _
                         Optional ByVal lngMaxMailsProOrdner As Long = 0, _
                         Optional ByVal intMaxTiefe As Integer = 5)
    On Error GoTo ErrHandler

    If Not ConnectRDO() Then
        LogError "SyncPostfach: Keine RDO-Verbindung", "SYNC"
        Exit Sub
    End If

    If lngMaxMailsProOrdner <= 0 Then
        lngMaxMailsProOrdner = CLng(LeseConfig("MaxMailsProSync", "500"))
    End If

    ' Store finden
    Dim objStore As Object
    Dim blnGefunden As Boolean: blnGefunden = False

    For Each objStore In g_objRDO.Stores
        If InStr(1, objStore.DisplayName, strPostfachName, vbTextCompare) > 0 Then
            blnGefunden = True
            Exit For
        End If
    Next objStore

    If Not blnGefunden Then
        LogError "SyncPostfach: Store '" & strPostfachName & "' nicht gefunden", "SYNC"
        Exit Sub
    End If

    Debug.Print String(70, "=")
    Debug.Print "=== POSTFACH-SYNC: " & objStore.DisplayName & " ==="
    Debug.Print "    Tiefe     : " & intMaxTiefe
    Debug.Print "    Max/Ordner: " & lngMaxMailsProOrdner
    Debug.Print "    " & Now()
    Debug.Print String(70, "=")

    g_blnAbbrechen = False
    g_dtSyncStart = Timer

    ' Root-Ordner und rekursiv alle Unterordner synchronisieren
    Dim objRoot As Object
    Set objRoot = objStore.RootFolder

    Call SyncSubfolder(objRoot, strProjekt, strPhase, lngMaxMailsProOrdner, 1, intMaxTiefe)

    Debug.Print String(70, "=")
    Debug.Print "=== POSTFACH-SYNC ABGESCHLOSSEN ==="
    Debug.Print "    Dauer: " & Format((Timer - g_dtSyncStart) / 60, "0.0") & " min"
    Debug.Print String(70, "=")

    Set objRoot = Nothing
    Exit Sub

ErrHandler:
    LogVBAError "SyncPostfach"
End Sub
