Attribute VB_Name = "modOutlookTest"
Option Compare Database
Option Explicit

' ===========================================================================
' modOutlookTest - Tiefgreifendes Testmodul fuer Outlook-Zugriff via Access
' ===========================================================================
' Getestete Zugriffswege:
'   1. Outlook Object Model (OOM)     - Outlook.Application
'   2. Redemption RDOSession          - Redemption.RDOSession  (EMPFOHLEN)
'   3. Redemption SafeMailItem        - Redemption.SafeMailItem
'   4. CDO 1.21                       - MAPI.Session (nicht verfuegbar)
'
' BESTAETIGT (04.03.2026): Outlook 16.0.0.17932 / Redemption 6.7.0.6412 / 64-Bit / Exchange
'   CheckOutlookDeepAccess: 35 OK / 6 FAIL (85%)
'   Alle FAIL = OOM Security Guard (erwartetes Verhalten, KEIN Fehler im Code)
'
'   OOM verfuegbar  : GetNamespace, GetDefaultFolder, Items.Count, Subject, EntryID
'   OOM blockiert   : SenderEmailAddress, Body, HTMLBody, PropertyAccessor, SaveAs  (Error 287)
'   RDO verfuegbar  : Logon, Folder, Subject, Body, HTMLBody, Fields[], SaveAs  --> KEIN Guard
'   Safe verfuegbar : Subject, SenderEmailAddress, Body, HTMLBody               --> KEIN Guard
'   CDO 1.21        : Nicht installiert (Error 429) - nicht benoetigt
'
' Voraussetzung:
'   - D:\Redemption64.dll (64-Bit Office) bzw. D:\Redemption.dll (32-Bit)
'   - Einmalige Registrierung: regsvr32 "D:\Redemption64.dll" (als Admin)
'
' Routinen:
'   CheckOutlookAccessMethods          - Schnelltest Verfuegbarkeit
'   CheckOutlookDeepAccess             - 11-Test Suite mit Zusammenfassung
'   AnalyzeSingleEmail                 - Mail vollstaendig sezieren (Metadaten/Header/Body/Anhaenge)
'   ExtractAttachmentsViaRedemption    - Dateinanhaenge extrahieren (Signatur-Bilder werden gefiltert)
'   ExportMailsAsMSG [n, "Unterordner"]- Batch MSG-Export
'   ShowFolderTree [Tiefe]             - Ordnerstruktur ausgeben
'
' Import in Access: Datei -> Importieren -> modOutlookTest.bas
' Aufruf:          Direktbereich (Strg+G):  CheckOutlookDeepAccess
' ===========================================================================

' ---------------------------------------------------------------------------
' KONSTANTEN
' ---------------------------------------------------------------------------
Private Const RDO_DLL_64    As String = "D:\Redemption64.dll"
Private Const RDO_DLL_32    As String = "D:\Redemption.dll"
Private Const EXPORT_BASE   As String = "\OutlookExport\"
Private Const ATT_BASE      As String = "\OutlookAttachments\"

' MAPI-Property-Tags fuer PropertyAccessor / RDO Fields[]
Private Const PR_SUBJECT                    As Long = &H37001E
Private Const PR_TRANSPORT_MESSAGE_HEADERS  As Long = &H7D001E
Private Const PR_INTERNET_MESSAGE_ID        As Long = &H1035001E
Private Const PR_MESSAGE_SIZE               As Long = &HE080003
Private Const PR_SENDER_EMAIL_ADDRESS       As Long = &HC1F001E
Private Const PR_SENDER_NAME                As Long = &HC1A001E
Private Const PR_DISPLAY_TO                 As Long = &HE04001E
Private Const PR_DISPLAY_CC                 As Long = &HE03001E

' Attachment-Typen
Private Const ATT_TYPE_DATA     As Integer = 1   ' Echte Datei (mapiattOLE)
Private Const ATT_TYPE_EMBEDDED As Integer = 5   ' Eingebettetes OLE
Private Const ATT_TYPE_MSG      As Integer = 6   ' Weitergeleitete Mail als Anlage

' OOM Folder-IDs
Private Const olFolderInbox         As Integer = 6
Private Const olFolderSentMail      As Integer = 5
Private Const olFolderDeletedItems  As Integer = 4
Private Const olFolderDrafts        As Integer = 16
Private Const olFolderCalendar      As Integer = 9
Private Const olFolderContacts      As Integer = 10

' Sonstiges
Private Const olMSG     As Integer = 3
Private Const olMail    As Integer = 43


' ===========================================================================
' 1. SCHNELLTEST - Verfuegbarkeit aller Zugriffswege pruefen
' ===========================================================================
Public Sub CheckOutlookAccessMethods()
    Dim objOutlook  As Object
    Dim objRDO      As Object
    Dim objSafe     As Object
    Dim objCDO      As Object

    Debug.Print String(70, "=")
    Debug.Print "=== SCHNELLTEST: Verfuegbarkeit der Outlook-Zugriffswege ==="
    Debug.Print "    " & Now()
    Debug.Print String(70, "=")

    On Error Resume Next

    ' --- OOM ---
    Err.Clear
    Set objOutlook = CreateObject("Outlook.Application")
    If Err.Number = 0 And Not objOutlook Is Nothing Then
        Debug.Print "[ OK   ] Outlook.Application  - Version " & objOutlook.Version
    Else
        Debug.Print "[ FAIL ] Outlook.Application  - " & Err.Number & ": " & Err.Description
    End If
    Err.Clear

    ' --- Redemption RDOSession ---
    Set objRDO = CreateObject("Redemption.RDOSession")
    If Err.Number = 0 And Not objRDO Is Nothing Then
        Debug.Print "[ OK   ] Redemption.RDOSession    - verfuegbar"
    Else
        Debug.Print "[ FAIL ] Redemption.RDOSession    - " & Err.Number & ": " & Err.Description
        Debug.Print "         TIPP: regsvr32 """ & RDO_DLL_64 & """ als Admin"
    End If
    Err.Clear

    ' --- Redemption SafeMailItem ---
    Set objSafe = CreateObject("Redemption.SafeMailItem")
    If Err.Number = 0 And Not objSafe Is Nothing Then
        Debug.Print "[ OK   ] Redemption.SafeMailItem  - verfuegbar"
    Else
        Debug.Print "[ FAIL ] Redemption.SafeMailItem  - " & Err.Number & ": " & Err.Description
    End If
    Err.Clear

    ' --- CDO 1.21 ---
    Set objCDO = CreateObject("MAPI.Session")
    If Err.Number = 0 And Not objCDO Is Nothing Then
        Debug.Print "[ OK   ] MAPI.Session (CDO 1.21) - verfuegbar"
    Else
        Debug.Print "[ FAIL ] MAPI.Session (CDO 1.21) - " & Err.Number & ": " & Err.Description
    End If
    Err.Clear

    On Error GoTo 0
    Set objOutlook = Nothing: Set objRDO = Nothing
    Set objSafe = Nothing:    Set objCDO = Nothing

    Debug.Print String(70, "=")
End Sub


' ===========================================================================
' 2. TIEFGREIFENDE PRUEFROUTINE - 11 Tests
'    Aufruf: CheckOutlookDeepAccess
' ===========================================================================
Public Sub CheckOutlookDeepAccess()

    Dim objOutlook          As Object
    Dim objNS               As Object
    Dim objFolder           As Object
    Dim objItem             As Object
    Dim objRefItem          As Object
    Dim objMSGItem          As Object
    Dim objAtt              As Object
    Dim objPA               As Object
    Dim objRedemptionRDO    As Object
    Dim objRDOFolder        As Object
    Dim objRDOItem          As Object
    Dim objSafeItem         As Object
    Dim objCDOSession       As Object
    Dim objCDOInbox         As Object
    Dim objCDOMsg           As Object
    Dim strTempFolder       As String
    Dim strTestMsgPath      As String
    Dim strVal              As String
    Dim strEntryID          As String
    Dim strAttPath          As String
    Dim lngOK               As Long
    Dim lngFail             As Long
    Dim lngAttCount         As Long
    Dim lngExportCount      As Long
    Dim i                   As Long
    Dim jj                  As Long
    Dim strExportPath       As String
    Dim strFileName         As String

    strTempFolder = Environ("TEMP") & "\"
    lngOK = 0: lngFail = 0

    Debug.Print String(70, "=")
    Debug.Print "=== START: Tiefergehende Pruefroutine der Outlook-Zugriffswege ==="
    Debug.Print "    Datum/Zeit : " & Now()
    Debug.Print "    TEMP-Pfad  : " & strTempFolder
    Debug.Print String(70, "=")

    On Error Resume Next

    ' -----------------------------------------------------------------------
    ' TEST 1: Outlook.Application + MAPI-Namespace
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 1 : Outlook.Application (CreateObject)"
    Debug.Print String(60, "-")

    Err.Clear
    Set objOutlook = CreateObject("Outlook.Application")
    If Err.Number = 0 And Not objOutlook Is Nothing Then
        Call Log_OK("Outlook.Application erstellt. Version: " & objOutlook.Version, lngOK)
    Else
        Call Log_FAIL("Outlook.Application: " & Err.Number & " - " & Err.Description, lngFail)
        GoTo SkipOOM
    End If

    Err.Clear
    Set objNS = objOutlook.GetNamespace("MAPI")
    If Err.Number = 0 And Not objNS Is Nothing Then
        Call Log_OK("GetNamespace(MAPI): OK. CurrentUser: " & objNS.CurrentUser.Name, lngOK)
    Else
        Call Log_FAIL("GetNamespace(MAPI): " & Err.Number & " - " & Err.Description, lngFail)
        GoTo SkipOOM
    End If

    ' -----------------------------------------------------------------------
    ' TEST 2: Standard-Ordner erreichbar? (OOM)
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 2 : Standard-Ordner erreichbar? (OOM)"
    Debug.Print String(60, "-")

    Dim aFolderID(5)    As Integer
    Dim aFolderName(5)  As String
    aFolderID(0) = olFolderInbox:        aFolderName(0) = "Posteingang        (olFolderInbox=6)"
    aFolderID(1) = olFolderSentMail:     aFolderName(1) = "Gesendete Elemente (olFolderSentMail=5)"
    aFolderID(2) = olFolderDeletedItems: aFolderName(2) = "Geloeschte El.     (olFolderDeletedItems=4)"
    aFolderID(3) = olFolderDrafts:       aFolderName(3) = "Entworfe           (olFolderDrafts=16)"
    aFolderID(4) = olFolderCalendar:     aFolderName(4) = "Kalender           (olFolderCalendar=9)"
    aFolderID(5) = olFolderContacts:     aFolderName(5) = "Kontakte           (olFolderContacts=10)"

    For i = 0 To 5
        Err.Clear
        Set objFolder = Nothing
        Set objFolder = objNS.GetDefaultFolder(aFolderID(i))
        If Err.Number = 0 And Not objFolder Is Nothing Then
            Call Log_OK(aFolderName(i) & " -> '" & objFolder.Name & "' (" & objFolder.Items.Count & " Elemente)", lngOK)
        Else
            Call Log_FAIL(aFolderName(i) & " -> " & Err.Number & " - " & Err.Description, lngFail)
        End If
    Next i

    ' -----------------------------------------------------------------------
    ' TEST 3: E-Mail-Inhalte lesen (OOM) - Subject, Body, HTMLBody
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 3 : E-Mail-Inhalte lesen (OOM) - Posteingang Item 1"
    Debug.Print String(60, "-")

    Err.Clear
    Set objFolder = objNS.GetDefaultFolder(olFolderInbox)
    If Err.Number <> 0 Or objFolder Is Nothing Then
        Call Log_FAIL("Posteingang nicht erreichbar (Test 3): " & Err.Number & " - " & Err.Description, lngFail)
        GoTo SkipAttachment
    End If

    If objFolder.Items.Count = 0 Then
        Debug.Print "  [ INFO ] Posteingang leer -> Tests 3-7 uebersprungen."
        GoTo SkipAttachment
    End If

    Err.Clear
    Set objItem = objFolder.Items(1)
    If Err.Number <> 0 Or objItem Is Nothing Then
        Call Log_FAIL("Items(1) nicht lesbar: " & Err.Number & " - " & Err.Description, lngFail)
        GoTo SkipAttachment
    End If

    Err.Clear: strVal = objItem.Subject
    If Err.Number = 0 Then Call Log_OK("Subject: '" & Left(strVal, 60) & "'", lngOK) _
    Else Call Log_FAIL("Subject: " & Err.Number & " - " & Err.Description, lngFail)

    Err.Clear: strVal = objItem.SenderEmailAddress
    If Err.Number = 0 Then Call Log_OK("SenderEmailAddress: '" & strVal & "'", lngOK) _
    Else Call Log_FAIL("SenderEmailAddress: " & Err.Number & " - " & Err.Description, lngFail)

    Err.Clear: strVal = CStr(objItem.ReceivedTime)
    If Err.Number = 0 Then Call Log_OK("ReceivedTime: " & strVal, lngOK) _
    Else Call Log_FAIL("ReceivedTime: " & Err.Number & " - " & Err.Description, lngFail)

    Err.Clear: strVal = objItem.Body
    If Err.Number = 0 Then Call Log_OK("Body (Plaintext) gelesen, " & Len(strVal) & " Zeichen", lngOK) _
    Else Call Log_FAIL("Body: " & Err.Number & " - " & Err.Description, lngFail)

    Err.Clear: strVal = objItem.HTMLBody
    If Err.Number = 0 Then Call Log_OK("HTMLBody gelesen, " & Len(strVal) & " Zeichen", lngOK) _
    Else Call Log_FAIL("HTMLBody: " & Err.Number & " - " & Err.Description, lngFail)

    Err.Clear: strEntryID = objItem.EntryID
    If Err.Number = 0 Then Call Log_OK("EntryID: " & Left(strEntryID, 30) & "...", lngOK) _
    Else Call Log_FAIL("EntryID: " & Err.Number & " - " & Err.Description, lngFail)

    ' -----------------------------------------------------------------------
    ' TEST 4: Anhaenge lesen und speichern (OOM)
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 4 : Anhaenge lesen und speichern (OOM)"
    Debug.Print String(60, "-")

    Err.Clear: lngAttCount = objItem.Attachments.Count
    If Err.Number = 0 Then
        Call Log_OK("Attachments.Count: " & lngAttCount, lngOK)
        If lngAttCount > 0 Then
            Err.Clear
            Set objAtt = objItem.Attachments(1)
            If Err.Number = 0 Then
                Debug.Print "         Anhang 1: '" & objAtt.FileName & "' (" & objAtt.Size & " Bytes)"
                strAttPath = strTempFolder & "OOM_Attachment_1_" & objAtt.FileName
                Err.Clear: objAtt.SaveAsFile strAttPath
                If Err.Number = 0 Then Call Log_OK("SaveAsFile: " & strAttPath, lngOK) _
                Else Call Log_FAIL("SaveAsFile: " & Err.Number & " - " & Err.Description, lngFail)
            Else
                Call Log_FAIL("Attachments(1): " & Err.Number & " - " & Err.Description, lngFail)
            End If
        Else
            Debug.Print "         [ INFO ] Diese Mail hat keine Anhaenge."
        End If
    Else
        Call Log_FAIL("Attachments.Count: " & Err.Number & " - " & Err.Description, lngFail)
    End If

    ' -----------------------------------------------------------------------
    ' TEST 5: Low-Level PropertyAccessor - MAPI-Props (OOM)
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 5 : PropertyAccessor - Low-Level MAPI-Props (OOM)"
    Debug.Print String(60, "-")

    Const PA_HEADERS  As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    Const PA_MSG_SIZE As String = "http://schemas.microsoft.com/mapi/proptag/0x0E080003"
    Const PA_MSG_ID   As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"

    Err.Clear
    Set objPA = objItem.PropertyAccessor
    If Err.Number = 0 And Not objPA Is Nothing Then
        Call Log_OK("PropertyAccessor erhalten", lngOK)

        Err.Clear: strVal = objPA.GetProperty(PA_HEADERS)
        If Err.Number = 0 Then Call Log_OK("PR_TRANSPORT_MESSAGE_HEADERS (" & Len(strVal) & " Zeichen)", lngOK) _
        Else Call Log_FAIL("PR_HEADERS: " & Err.Number & " - " & Err.Description, lngFail)

        Dim lngMsgSize As Long
        Err.Clear: lngMsgSize = objPA.GetProperty(PA_MSG_SIZE)
        If Err.Number = 0 Then Call Log_OK("PR_MESSAGE_SIZE: " & lngMsgSize & " Bytes", lngOK) _
        Else Call Log_FAIL("PR_MESSAGE_SIZE: " & Err.Number & " - " & Err.Description, lngFail)

        Err.Clear: strVal = objPA.GetProperty(PA_MSG_ID)
        If Err.Number = 0 Then Call Log_OK("PR_INTERNET_MESSAGE_ID: '" & strVal & "'", lngOK) _
        Else Call Log_FAIL("PR_INTERNET_MESSAGE_ID: " & Err.Number & " - " & Err.Description, lngFail)
    Else
        Call Log_FAIL("PropertyAccessor: " & Err.Number & " - " & Err.Description, lngFail)
    End If

    ' -----------------------------------------------------------------------
    ' TEST 6: MSG speichern und wieder oeffnen (OOM)
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 6 : MSG speichern und wieder oeffnen (OOM)"
    Debug.Print String(60, "-")

    strTestMsgPath = strTempFolder & "OOM_TestMail.msg"
    Err.Clear: objItem.SaveAs strTestMsgPath, olMSG
    If Err.Number = 0 Then
        Call Log_OK("SaveAs MSG: " & strTestMsgPath, lngOK)
        If Len(Dir(strTestMsgPath)) > 0 Then
            Call Log_OK("MSG physisch vorhanden (" & FileLen(strTestMsgPath) & " Bytes)", lngOK)
        Else
            Call Log_FAIL("MSG-Datei nach SaveAs nicht gefunden!", lngFail)
        End If
        Err.Clear
        Set objMSGItem = objOutlook.CreateItemFromTemplate(strTestMsgPath)
        If Err.Number = 0 And Not objMSGItem Is Nothing Then
            Call Log_OK("MSG wieder geoeffnet. Subject: '" & Left(objMSGItem.Subject, 50) & "'", lngOK)
            objMSGItem.Close 1   ' 1 = olDiscard
        Else
            Call Log_FAIL("CreateItemFromTemplate: " & Err.Number & " - " & Err.Description, lngFail)
        End If
    Else
        Call Log_FAIL("SaveAs MSG: " & Err.Number & " - " & Err.Description, lngFail)
    End If

SkipAttachment:
    Err.Clear

    ' -----------------------------------------------------------------------
    ' TEST 7: GetItemFromID ueber EntryID (OOM)
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 7 : GetItemFromID ueber EntryID (OOM)"
    Debug.Print String(60, "-")

    If Len(strEntryID) > 0 And Not objNS Is Nothing Then
        Err.Clear
        Set objRefItem = objNS.GetItemFromID(strEntryID)
        If Err.Number = 0 And Not objRefItem Is Nothing Then
            Call Log_OK("GetItemFromID OK. Subject: '" & Left(objRefItem.Subject, 50) & "'", lngOK)
        Else
            Call Log_FAIL("GetItemFromID: " & Err.Number & " - " & Err.Description, lngFail)
        End If
    Else
        Debug.Print "  [ INFO ] Test 7 uebersprungen (keine EntryID verfuegbar)."
    End If
    Err.Clear

SkipOOM:
    Err.Clear

    ' -----------------------------------------------------------------------
    ' TEST 8: CDO 1.21 (MAPI.Session)
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 8 : CDO 1.21 - MAPI-Zugriff (MAPI.Session)"
    Debug.Print String(60, "-")

    Err.Clear
    Set objCDOSession = CreateObject("MAPI.Session")
    If Err.Number = 0 And Not objCDOSession Is Nothing Then
        Call Log_OK("MAPI.Session (CDO 1.21) erstellt", lngOK)
        Err.Clear: objCDOSession.Logon "", "", False, False
        If Err.Number = 0 Then
            Call Log_OK("CDO Logon OK. Name: " & objCDOSession.CurrentUser.Name, lngOK)
            Err.Clear
            Set objCDOInbox = objCDOSession.Inbox
            If Err.Number = 0 And Not objCDOInbox Is Nothing Then
                Call Log_OK("CDO Inbox: '" & objCDOInbox.Name & "' (" & objCDOInbox.Messages.Count & " Nachrichten)", lngOK)
                If objCDOInbox.Messages.Count > 0 Then
                    Err.Clear
                    Set objCDOMsg = objCDOInbox.Messages.Item(1)
                    If Err.Number = 0 Then
                        Call Log_OK("CDO Msg Subject: '" & Left(objCDOMsg.Subject, 50) & "'", lngOK)
                        Call Log_OK("CDO Msg Sender : " & objCDOMsg.Sender.Name, lngOK)
                    Else
                        Call Log_FAIL("CDO Messages.Item(1): " & Err.Number & " - " & Err.Description, lngFail)
                    End If
                Else
                    Debug.Print "         [ INFO ] CDO Inbox ist leer."
                End If
            Else
                Call Log_FAIL("CDO Inbox: " & Err.Number & " - " & Err.Description, lngFail)
            End If
            objCDOSession.Logoff
        Else
            Call Log_FAIL("CDO Logon: " & Err.Number & " - " & Err.Description, lngFail)
        End If
    Else
        Call Log_FAIL("MAPI.Session (CDO 1.21 nicht installiert?): " & Err.Number & " - " & Err.Description, lngFail)
    End If
    Err.Clear

    ' -----------------------------------------------------------------------
    ' TEST 9: Redemption RDOSession - voller MAPI-Zugriff
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 9 : Redemption.RDOSession"
    Debug.Print String(60, "-")

    Err.Clear
    Set objRedemptionRDO = CreateObject("Redemption.RDOSession")
    If Err.Number = 0 And Not objRedemptionRDO Is Nothing Then
        Call Log_OK("Redemption.RDOSession erstellt", lngOK)
        Err.Clear: objRedemptionRDO.Logon
        If Err.Number = 0 Then
            Call Log_OK("RDO Logon OK. Version: " & objRedemptionRDO.Version, lngOK)

            Err.Clear
            Set objRDOFolder = objRedemptionRDO.GetDefaultFolder(olFolderInbox)
            If Err.Number = 0 And Not objRDOFolder Is Nothing Then
                Call Log_OK("RDO Inbox: '" & objRDOFolder.Name & "' (" & objRDOFolder.Items.Count & " Elemente)", lngOK)

                If objRDOFolder.Items.Count > 0 Then
                    Err.Clear
                    Set objRDOItem = objRDOFolder.Items(1)
                    If Err.Number = 0 And Not objRDOItem Is Nothing Then

                        Err.Clear: strVal = objRDOItem.Subject
                        If Err.Number = 0 Then Call Log_OK("RDO Subject: '" & Left(strVal, 50) & "'", lngOK) _
                        Else Call Log_FAIL("RDO Subject: " & Err.Number & " - " & Err.Description, lngFail)

                        Err.Clear: strVal = objRDOItem.Fields(PR_SUBJECT)
                        If Err.Number = 0 Then Call Log_OK("RDO Fields[PR_SUBJECT]: '" & Left(strVal, 50) & "'", lngOK) _
                        Else Call Log_FAIL("RDO Fields[PR_SUBJECT]: " & Err.Number & " - " & Err.Description, lngFail)

                        Err.Clear: strVal = objRDOItem.Fields(PR_TRANSPORT_MESSAGE_HEADERS)
                        If Err.Number = 0 Then Call Log_OK("RDO Fields[PR_HEADERS] (" & Len(strVal) & " Zeichen)", lngOK) _
                        Else Call Log_FAIL("RDO Fields[PR_HEADERS]: " & Err.Number & " - " & Err.Description, lngFail)

                        Err.Clear: strVal = objRDOItem.Body
                        If Err.Number = 0 Then Call Log_OK("RDO Body (" & Len(strVal) & " Zeichen)", lngOK) _
                        Else Call Log_FAIL("RDO Body: " & Err.Number & " - " & Err.Description, lngFail)

                        Err.Clear: strVal = objRDOItem.HTMLBody
                        If Err.Number = 0 Then Call Log_OK("RDO HTMLBody (" & Len(strVal) & " Zeichen)", lngOK) _
                        Else Call Log_FAIL("RDO HTMLBody: " & Err.Number & " - " & Err.Description, lngFail)

                        strTestMsgPath = strTempFolder & "RDO_TestMail.msg"
                        Err.Clear: objRDOItem.SaveAs strTestMsgPath
                        If Err.Number = 0 Then Call Log_OK("RDO SaveAs MSG: " & strTestMsgPath, lngOK) _
                        Else Call Log_FAIL("RDO SaveAs: " & Err.Number & " - " & Err.Description, lngFail)

                    Else
                        Call Log_FAIL("RDO Items(1): " & Err.Number & " - " & Err.Description, lngFail)
                    End If
                Else
                    Debug.Print "         [ INFO ] RDO Inbox ist leer."
                End If
            Else
                Call Log_FAIL("RDO GetDefaultFolder(6): " & Err.Number & " - " & Err.Description, lngFail)
            End If
        Else
            Call Log_FAIL("RDO Logon: " & Err.Number & " - " & Err.Description, lngFail)
        End If
    Else
        Call Log_FAIL("Redemption.RDOSession nicht registriert: " & Err.Number & " - " & Err.Description, lngFail)
    End If
    Err.Clear

    ' -----------------------------------------------------------------------
    ' TEST 10: Redemption SafeMailItem (OOM-Wrapper ohne Security-Guard)
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 10: Redemption SafeMailItem (OOM-Wrapper)"
    Debug.Print String(60, "-")

    If Not objItem Is Nothing Then
        Err.Clear
        Dim objSafeItem As Object
        Set objSafeItem = CreateObject("Redemption.SafeMailItem")
        If Err.Number = 0 And Not objSafeItem Is Nothing Then
            Call Log_OK("Redemption.SafeMailItem erstellt", lngOK)
            Err.Clear: objSafeItem.Item = objItem
            If Err.Number = 0 Then
                Call Log_OK("SafeMailItem.Item gesetzt", lngOK)

                Err.Clear: strVal = objSafeItem.Subject
                If Err.Number = 0 Then Call Log_OK("SafeMailItem.Subject: '" & Left(strVal, 50) & "'", lngOK) _
                Else Call Log_FAIL("SafeMailItem.Subject: " & Err.Number & " - " & Err.Description, lngFail)

                Err.Clear: strVal = objSafeItem.SenderEmailAddress
                If Err.Number = 0 Then Call Log_OK("SafeMailItem.SenderEmailAddress: '" & strVal & "'", lngOK) _
                Else Call Log_FAIL("SafeMailItem.SenderEmailAddress: " & Err.Number & " - " & Err.Description, lngFail)

                Err.Clear: strVal = objSafeItem.Body
                If Err.Number = 0 Then Call Log_OK("SafeMailItem.Body (" & Len(strVal) & " Zeichen)", lngOK) _
                Else Call Log_FAIL("SafeMailItem.Body: " & Err.Number & " - " & Err.Description, lngFail)

                Err.Clear: strVal = objSafeItem.HTMLBody
                If Err.Number = 0 Then Call Log_OK("SafeMailItem.HTMLBody (" & Len(strVal) & " Zeichen)", lngOK) _
                Else Call Log_FAIL("SafeMailItem.HTMLBody: " & Err.Number & " - " & Err.Description, lngFail)
            Else
                Call Log_FAIL("SafeMailItem.Item setzen: " & Err.Number & " - " & Err.Description, lngFail)
            End If
        Else
            Call Log_FAIL("Redemption.SafeMailItem: " & Err.Number & " - " & Err.Description, lngFail)
        End If
        Set objSafeItem = Nothing
    Else
        Debug.Print "  [ INFO ] Test 10 uebersprungen (kein OOM-Item verfuegbar)."
    End If
    Err.Clear

    ' -----------------------------------------------------------------------
    ' TEST 11: Redemption Batch-Export - erste 5 IPM.Note Mails als MSG
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(60, "-")
    Debug.Print "TEST 11: Redemption Batch-Export (erste 5 E-Mails als MSG)"
    Debug.Print String(60, "-")

    strExportPath = Environ("TEMP") & EXPORT_BASE
    lngExportCount = 0

    Err.Clear
    If Dir(strExportPath, vbDirectory) = "" Then MkDir strExportPath
    If Err.Number = 0 Then
        Call Log_OK("Export-Verzeichnis bereit: " & strExportPath, lngOK)
    Else
        Call Log_FAIL("MkDir " & strExportPath & ": " & Err.Number & " - " & Err.Description, lngFail)
        GoTo SkipBatchExport
    End If

    If objRedemptionRDO Is Nothing Then
        Err.Clear
        Set objRedemptionRDO = CreateObject("Redemption.RDOSession")
        If Err.Number = 0 Then objRedemptionRDO.Logon
    End If

    If Err.Number <> 0 Or objRedemptionRDO Is Nothing Then
        Call Log_FAIL("RDO-Session fuer Batch-Export nicht verfuegbar: " & Err.Number & " - " & Err.Description, lngFail)
        GoTo SkipBatchExport
    End If

    Err.Clear
    Set objRDOFolder = objRedemptionRDO.GetDefaultFolder(olFolderInbox)
    If Err.Number <> 0 Or objRDOFolder Is Nothing Then
        Call Log_FAIL("Batch-Export GetDefaultFolder(6): " & Err.Number & " - " & Err.Description, lngFail)
        GoTo SkipBatchExport
    End If

    Call Log_OK("Ordner: '" & objRDOFolder.Name & "' (" & objRDOFolder.Items.Count & " Elemente)", lngOK)

    If objRDOFolder.Items.Count = 0 Then
        Debug.Print "  [ INFO ] Posteingang leer - Batch-Export uebersprungen."
        GoTo SkipBatchExport
    End If

    For jj = 1 To objRDOFolder.Items.Count
        Err.Clear
        Set objRDOItem = objRDOFolder.Items(jj)
        If Err.Number <> 0 Or objRDOItem Is Nothing Then
            Call Log_FAIL("Batch Items(" & jj & "): " & Err.Number & " - " & Err.Description, lngFail)
        Else
            If Left(objRDOItem.MessageClass, 8) = "IPM.Note" Then
                strFileName = CleanFileName(objRDOItem.Subject, objRDOItem.ReceivedTime)
                Err.Clear
                objRDOItem.SaveAs strExportPath & strFileName
                If Err.Number = 0 Then
                    lngExportCount = lngExportCount + 1
                    Call Log_OK("Export " & lngExportCount & ": " & strFileName, lngOK)
                Else
                    Call Log_FAIL("SaveAs '" & strFileName & "': " & Err.Number & " - " & Err.Description, lngFail)
                End If
                If lngExportCount >= 5 Then Exit For
            End If
        End If
    Next jj

    If lngExportCount > 0 Then
        Call Log_OK("Batch-Export abgeschlossen: " & lngExportCount & " MSG-Dateien -> " & strExportPath, lngOK)
    Else
        Debug.Print "  [ INFO ] Keine IPM.Note-Nachrichten gefunden."
    End If

SkipBatchExport:
    Err.Clear

    On Error GoTo 0

    ' -----------------------------------------------------------------------
    ' ERKENNTNISSE
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(70, "=")
    Debug.Print "=== ERKENNTNISSE ==="
    Debug.Print String(70, "-")
    Debug.Print "OOM (Outlook.Application)"
    Debug.Print "  VERFUEGBAR : GetNamespace, GetDefaultFolder, Items.Count, Subject, EntryID"
    Debug.Print "  BLOCKIERT  : SenderEmailAddress, Body, HTMLBody, PropertyAccessor, SaveAs"
    Debug.Print "  URSACHE    : OOM Security Guard (Error 287 / HRESULT E_FAIL)"
    Debug.Print "  UMGEHUNG   : Redemption SafeMailItem oder RDOSession"
    Debug.Print String(70, "-")
    Debug.Print "CDO 1.21 (MAPI.Session)"
    Debug.Print "  STATUS     : Nicht installiert (Error 429)"
    Debug.Print "  ALTERNATIVE: Redemption RDOSession"
    Debug.Print String(70, "-")
    Debug.Print "Redemption.RDOSession  (EMPFOHLEN fuer Vollzugriff)"
    Debug.Print "  VERFUEGBAR : Logon, Folder, Subject, Body, HTMLBody, Fields[], SaveAs"
    Debug.Print "  BESONDERHEIT: Kein OOM Guard - direkter MAPI-Zugriff"
    Debug.Print "  FIELDS[]   : MAPI-Props per &H-Hex-Tag (z.B. &H7D001E = PR_HEADERS)"
    Debug.Print String(70, "-")
    Debug.Print "Redemption.SafeMailItem  (fuer OOM-Objekte im Umlauf)"
    Debug.Print "  VERFUEGBAR : Subject, SenderEmailAddress, Body, HTMLBody"
    Debug.Print "  BESONDERHEIT: OOM-Wrapper - kein eigener Logon noetig"
    Debug.Print "  EINSATZ    : Wenn OOM-Item bereits vorliegt (z.B. aus Items-Collection)"
    Debug.Print String(70, "-")
    Debug.Print "Batch-Export (RDO SaveAs)"
    Debug.Print "  ERGEBNIS   : Zuverlaessig, kein Guard"
    Debug.Print "  DATEINAME  : yyyymmdd_hhmm_Betreff.msg"
    Debug.Print "  FILTER     : Left(MessageClass,8) = 'IPM.Note'"
    Debug.Print String(70, "=")

    ' -----------------------------------------------------------------------
    ' ZUSAMMENFASSUNG
    ' -----------------------------------------------------------------------
    Debug.Print vbCrLf & String(70, "=")
    Debug.Print "=== ZUSAMMENFASSUNG ==="
    Debug.Print "  Erfolgreich [OK]  : " & lngOK
    Debug.Print "  Fehlgeschlagen    : " & lngFail
    Debug.Print "  Gesamt            : " & (lngOK + lngFail)
    If lngFail = 0 Then
        Debug.Print "  Ergebnis          : ALLE TESTS BESTANDEN"
    ElseIf lngOK = 0 Then
        Debug.Print "  Ergebnis          : ALLE TESTS FEHLGESCHLAGEN"
    Else
        Debug.Print "  Ergebnis          : TEILWEISE ERFOLGREICH (" & _
                    Format(lngOK / (lngOK + lngFail) * 100, "0") & "% OK)"
    End If
    Debug.Print String(70, "=")

    ' Aufraeum
    Set objAtt = Nothing:         Set objPA = Nothing
    Set objMSGItem = Nothing:     Set objRefItem = Nothing
    Set objItem = Nothing:        Set objFolder = Nothing
    Set objNS = Nothing:          Set objOutlook = Nothing
    Set objRDOItem = Nothing:     Set objRDOFolder = Nothing
    Set objRedemptionRDO = Nothing
    Set objCDOMsg = Nothing:      Set objCDOInbox = Nothing
    Set objCDOSession = Nothing
End Sub


' ===========================================================================
' 3. EINZELNE MAIL VOLLSTAENDIG ANALYSIEREN
'    Aufruf: AnalyzeSingleEmail
' ===========================================================================
Public Sub AnalyzeSingleEmail()

    Dim objRDO          As Object
    Dim objInbox        As Object
    Dim objItem         As Object
    Dim objRecipient    As Object
    Dim objAtt          As Object
    Dim i               As Integer
    Dim a               As Integer
    Dim strSMTP         As String
    Dim strTo           As String
    Dim strCC           As String
    Dim strBCC          As String
    Dim strBodySnippet  As String

    Debug.Print String(70, "=")
    Debug.Print "=== START: Vollanalyse einer E-Mail via Redemption ==="
    Debug.Print "    " & Now()
    Debug.Print String(70, "=")

    On Error GoTo ErrHandler

    Set objRDO = CreateObject("Redemption.RDOSession")
    objRDO.Logon
    Set objInbox = objRDO.GetDefaultFolder(olFolderInbox)

    ' Erste echte E-Mail suchen
    For i = 1 To objInbox.Items.Count
        Set objItem = objInbox.Items(i)
        If Left(objItem.MessageClass, 8) = "IPM.Note" Then

            ' --- METADATEN ---
            Debug.Print vbCrLf & "--- ALLGEMEINE METADATEN ---"
            Debug.Print "Betreff:            " & objItem.Subject
            Debug.Print "MessageClass:       " & objItem.MessageClass
            Debug.Print "EntryID:            " & Left(objItem.EntryID, 40) & "..."
            Debug.Print "Erstellt am:        " & objItem.CreationTime
            Debug.Print "Empfangen am:       " & objItem.ReceivedTime
            Debug.Print "Gesendet am:        " & objItem.SentOn
            Debug.Print "Groesse (Bytes):    " & objItem.Size
            Debug.Print "Gelesen:            " & Not (objItem.UnRead)
            Debug.Print "Wichtigkeit:        " & objItem.Importance & "  (0=Niedrig 1=Normal 2=Hoch)"
            Debug.Print "Kategorie(n):       " & objItem.Categories

            ' --- ABSENDER ---
            Debug.Print vbCrLf & "--- ABSENDER ---"
            Debug.Print "Anzeigename:        " & objItem.SenderName
            Debug.Print "Mail-Typ (EX/SMTP): " & objItem.SenderEmailType
            ' Exchange-Adressen korrekt aufloesen
            If objItem.SenderEmailType = "EX" Then
                strSMTP = GetSMTP(objItem.Sender)
            Else
                strSMTP = objItem.SenderEmailAddress
            End If
            Debug.Print "Echte SMTP:         " & strSMTP
            Debug.Print "RDO PR_SENDER_NAME: " & objItem.Fields(PR_SENDER_NAME)
            Debug.Print "RDO PR_SENDER_ADDR: " & objItem.Fields(PR_SENDER_EMAIL_ADDRESS)

            ' --- EMPFAENGER ---
            Debug.Print vbCrLf & "--- EMPFAENGER ---"
            strTo = "": strCC = "": strBCC = ""

            Dim recSMTP As String
            For Each objRecipient In objItem.Recipients
                If objRecipient.AddressEntry.Type = "EX" Then
                    recSMTP = GetSMTP(objRecipient.AddressEntry)
                Else
                    recSMTP = objRecipient.Address
                End If
                Select Case objRecipient.Type
                    Case 1: strTo  = strTo  & objRecipient.Name & " <" & recSMTP & ">; "
                    Case 2: strCC  = strCC  & objRecipient.Name & " <" & recSMTP & ">; "
                    Case 3: strBCC = strBCC & objRecipient.Name & " <" & recSMTP & ">; "
                End Select
            Next objRecipient

            If Len(strTo)  > 2 Then Debug.Print "AN (To):  " & Left(strTo,  Len(strTo)  - 2) Else Debug.Print "AN (To):  -"
            If Len(strCC)  > 2 Then Debug.Print "CC:       " & Left(strCC,  Len(strCC)  - 2) Else Debug.Print "CC:       -"
            If Len(strBCC) > 2 Then Debug.Print "BCC:      " & Left(strBCC, Len(strBCC) - 2) Else Debug.Print "BCC:      -"
            Debug.Print "RDO PR_DISPLAY_TO:  " & objItem.Fields(PR_DISPLAY_TO)
            Debug.Print "RDO PR_DISPLAY_CC:  " & objItem.Fields(PR_DISPLAY_CC)

            ' --- INHALT ---
            Debug.Print vbCrLf & "--- INHALT ---"
            Debug.Print "Hat HTML-Body:      " & (Len(objItem.HTMLBody) > 0)
            Debug.Print "Laenge Plaintext:   " & Len(objItem.Body) & " Zeichen"
            Debug.Print "Laenge HTMLBody:    " & Len(objItem.HTMLBody) & " Zeichen"
            strBodySnippet = Replace(Replace(Left(objItem.Body, 120), vbCr, ""), vbLf, " | ")
            Debug.Print "Text-Vorschau:      " & strBodySnippet

            ' --- INTERNET-HEADER ---
            Debug.Print vbCrLf & "--- INTERNET-HEADER (erste 400 Zeichen) ---"
            Dim strHeader As String
            strHeader = objItem.Fields(PR_TRANSPORT_MESSAGE_HEADERS)
            If Len(strHeader) > 0 Then
                Debug.Print Left(strHeader, 400)
            Else
                Debug.Print "(keine Header-Daten verfuegbar)"
            End If

            ' --- ANHAENGE ---
            Debug.Print vbCrLf & "--- ANHAENGE (" & objItem.Attachments.Count & " gesamt) ---"
            For a = 1 To objItem.Attachments.Count
                Set objAtt = objItem.Attachments(a)
                Debug.Print "  " & a & ". Dateiname : " & objAtt.FileName
                Debug.Print "     Groesse   : " & objAtt.Size & " Bytes"
                Debug.Print "     Typ       : " & objAtt.Type & _
                            "  (1=Datei  5=OLE/eingebettet  6=angehaengte Mail)"
                Debug.Print "     MimeType  : " & objAtt.MimeTag
                ' Hidden ist True bei Inline-Bildern (z.B. Signatur-Logo)
                Debug.Print "     Versteckt : " & objAtt.Hidden
            Next a
            If objItem.Attachments.Count = 0 Then Debug.Print "  (keine Anhaenge)"

            Exit For    ' Nur die erste Mail
        End If
    Next i

    Cleanup:
    Set objAtt = Nothing:  Set objRecipient = Nothing
    Set objItem = Nothing: Set objInbox = Nothing
    Set objRDO = Nothing
    Debug.Print vbCrLf & String(70, "=")
    Debug.Print "=== ENDE ==="
    Debug.Print String(70, "=")
    Exit Sub

ErrHandler:
    Debug.Print "[ FEHLER ] " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Sub


' ===========================================================================
' 4. ANHAENGE EXTRAHIEREN (RDO - ohne OOM Guard)
'    Extraktion aller echten Datei-Anhaenge aus den ersten N Mails
'    Aufruf: ExtractAttachmentsViaRedemption
' ===========================================================================
Public Sub ExtractAttachmentsViaRedemption(Optional intMaxMails As Integer = 20)

    Dim objRDO          As Object
    Dim objInbox        As Object
    Dim objItem         As Object
    Dim objAtt          As Object
    Dim strAttDir       As String
    Dim strAttPath      As String
    Dim strFileName     As String
    Dim lngMailCount    As Long
    Dim lngAttCount     As Long
    Dim lngSkipped      As Long
    Dim i               As Integer
    Dim a               As Integer

    strAttDir = Environ("TEMP") & ATT_BASE

    Debug.Print String(70, "=")
    Debug.Print "=== START: Anhang-Extraktion via Redemption ==="
    Debug.Print "    Ziel-Verzeichnis : " & strAttDir
    Debug.Print "    Max. Mails        : " & intMaxMails
    Debug.Print "    " & Now()
    Debug.Print String(70, "=")

    On Error GoTo ErrHandler

    ' Zielordner anlegen
    If Dir(strAttDir, vbDirectory) = "" Then MkDir strAttDir
    Debug.Print "[ OK ] Verzeichnis bereit."

    ' RDO initialisieren
    Set objRDO = CreateObject("Redemption.RDOSession")
    objRDO.Logon
    Set objInbox = objRDO.GetDefaultFolder(olFolderInbox)
    Debug.Print "[ OK ] RDO Logon. Inbox: " & objInbox.Items.Count & " Elemente."

    ' Iteration
    lngMailCount = 0: lngAttCount = 0: lngSkipped = 0

    For i = 1 To objInbox.Items.Count
        Set objItem = objInbox.Items(i)

        ' Nur echte E-Mails
        If Left(objItem.MessageClass, 8) = "IPM.Note" Then
            lngMailCount = lngMailCount + 1

            ' Nach echten Anhaengen suchen
            For a = 1 To objItem.Attachments.Count
                Set objAtt = objItem.Attachments(a)

                ' ----------------------------------------------------------------
                '  Filter: Signatur-Bilder und eingebettete OLE-Objekte ueberspringen
                '  Hidden=True:  inline-Bild in HTML-Signatur
                '  Type<>1:      kein normaler Datei-Anhang
                ' ----------------------------------------------------------------
                If objAtt.Hidden Or objAtt.Type <> ATT_TYPE_DATA Then
                    lngSkipped = lngSkipped + 1
                    ' Debug.Print "    -> SKIP: '" & objAtt.FileName & "' (typ=" & objAtt.Type & " hidden=" & objAtt.Hidden & ")"
                Else
                    ' Dateiname bereinigen und eindeutig machen
                    strFileName = Format(objItem.ReceivedTime, "yyyymmdd_hhnn") & "_" & _
                                  Left(CleanFileNameSimple(objItem.Subject), 30) & "_" & _
                                  CleanFileNameSimple(objAtt.FileName)

                    strAttPath = strAttDir & strFileName

                    On Error Resume Next
                    objAtt.SaveAsFile strAttPath
                    If Err.Number = 0 Then
                        lngAttCount = lngAttCount + 1
                        Debug.Print "[ OK ] Anhang " & lngAttCount & ": " & strFileName
                        Debug.Print "       Pfad  : " & strAttPath
                        Debug.Print "       Groesse: " & objAtt.Size & " Bytes"
                    Else
                        Debug.Print "[ FAIL] '" & strFileName & "': " & Err.Number & " - " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo ErrHandler
                End If
            Next a

            ' Oberlimit pruefen
            If lngMailCount >= intMaxMails Then Exit For
        End If
    Next i

    Debug.Print String(70, "-")
    Debug.Print "Durchsuchte Mails    : " & lngMailCount
    Debug.Print "Echte Anhaenge       : " & lngAttCount
    Debug.Print "Uebersprungen        : " & lngSkipped & "  (Signatur-Bilder / eingebettet)"
    Debug.Print "Speicherort          : " & strAttDir

    Cleanup:
    Set objAtt = Nothing:  Set objItem = Nothing
    Set objInbox = Nothing: Set objRDO = Nothing
    Debug.Print String(70, "=")
    Debug.Print "=== ENDE ==="
    Debug.Print String(70, "=")
    Exit Sub

ErrHandler:
    Debug.Print "[ FEHLER ] " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Sub


' ===========================================================================
' 5. BATCH-EXPORT MAILS ALS MSG
'    Aufruf: ExportMailsAsMSG  (optional: Anzahl und Unterordner-Name)
' ===========================================================================
Public Sub ExportMailsAsMSG(Optional intMaxItems As Integer = 10, _
                            Optional strSubFolder As String = "")
    Dim objRDO      As Object
    Dim objInbox    As Object
    Dim objFolder   As Object
    Dim objItem     As Object
    Dim strDir      As String
    Dim strFileName As String
    Dim lngCount    As Long
    Dim i           As Integer

    strDir = Environ("TEMP") & EXPORT_BASE

    Debug.Print String(70, "=")
    Debug.Print "=== START: Batch-Export Mails als MSG via Redemption ==="
    If strSubFolder <> "" Then Debug.Print "    Unterordner: " & strSubFolder
    Debug.Print "    Max. Mails : " & intMaxItems
    Debug.Print "    Ziel       : " & strDir
    Debug.Print String(70, "=")

    On Error GoTo ErrHandler

    If Dir(strDir, vbDirectory) = "" Then MkDir strDir

    Set objRDO = CreateObject("Redemption.RDOSession")
    objRDO.Logon
    Set objInbox = objRDO.GetDefaultFolder(olFolderInbox)

    ' Unterordner oder Posteingang verwenden
    If strSubFolder <> "" Then
        Set objFolder = objInbox.Folders(strSubFolder)
        If objFolder Is Nothing Then
            Debug.Print "[ FAIL ] Unterordner '" & strSubFolder & "' nicht gefunden!"
            Debug.Print "         Verfuegbare Unterordner:"
            Dim sf As Object
            For Each sf In objInbox.Folders
                Debug.Print "           - " & sf.Name
            Next sf
            GoTo Cleanup
        End If
        Debug.Print "[ OK ] Unterordner: '" & objFolder.Name & "' (" & objFolder.Items.Count & " Elemente)"
    Else
        Set objFolder = objInbox
        Debug.Print "[ OK ] Posteingang: " & objFolder.Items.Count & " Elemente"
    End If

    lngCount = 0
    For i = 1 To objFolder.Items.Count
        Set objItem = objFolder.Items(i)
        If Left(objItem.MessageClass, 8) = "IPM.Note" Then
            strFileName = CleanFileName(objItem.Subject, objItem.ReceivedTime)
            On Error Resume Next
            objItem.SaveAs strDir & strFileName
            If Err.Number = 0 Then
                lngCount = lngCount + 1
                Debug.Print "[ OK ] " & lngCount & ": " & strFileName
            Else
                Debug.Print "[ FAIL] " & strFileName & ": " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
            If lngCount >= intMaxItems Then Exit For
        End If
    Next i

    Debug.Print String(70, "-")
    Debug.Print "Exportiert: " & lngCount & " MSG-Dateien -> " & strDir

    Cleanup:
    Set objItem = Nothing:  Set objFolder = Nothing
    Set objInbox = Nothing: Set objRDO = Nothing
    Debug.Print String(70, "=")
    Debug.Print "=== ENDE ==="
    Debug.Print String(70, "=")
    Exit Sub

ErrHandler:
    Debug.Print "[ FEHLER ] " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Sub


' ===========================================================================
' 6. ORDNERSTRUKTUR DES POSTFACHS AUSGEBEN
'    Aufruf: ShowFolderTree  (optional: Tiefe)
' ===========================================================================
Public Sub ShowFolderTree(Optional intDepth As Integer = 2)

    Dim objRDO      As Object
    Dim objRoot     As Object

    Debug.Print String(70, "=")
    Debug.Print "=== Ordnerstruktur (Redemption RDO, Tiefe=" & intDepth & ") ==="
    Debug.Print "    " & Now()
    Debug.Print String(70, "=")

    On Error GoTo ErrHandler

    Set objRDO = CreateObject("Redemption.RDOSession")
    objRDO.Logon

    ' Postfach-Stammordner korrekt ermitteln (DefaultStore.RootFolder, nicht Parent von Inbox)
    Set objRoot = objRDO.Stores.DefaultStore.RootFolder
    Debug.Print "Postfach: " & objRoot.Name & "  [Store: " & objRDO.Stores.DefaultStore.DisplayName & "]"
    Call PrintFolderTree(objRoot, 1, intDepth)

    Cleanup:
    Set objRoot = Nothing: Set objRDO = Nothing
    Debug.Print String(70, "=")
    Debug.Print "=== ENDE ==="
    Debug.Print String(70, "=")
    Exit Sub

ErrHandler:
    Debug.Print "[ FEHLER ] " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Sub

Private Sub PrintFolderTree(objParent As Object, intLevel As Integer, intMaxDepth As Integer)
    Dim objSub  As Object
    Dim strPad  As String
    Dim strInfo As String

    If intLevel > intMaxDepth Then Exit Sub
    strPad = String(intLevel * 2, " ") & "|- "

    On Error Resume Next
    For Each objSub In objParent.Folders
        strInfo = strPad & objSub.Name
        ' Elementanzahl nur bis Tiefe 2 anzeigen (Performance)
        If intLevel <= 2 Then strInfo = strInfo & "  (" & objSub.Items.Count & " Elemente)"
        Debug.Print strInfo
        If intLevel < intMaxDepth Then
            Call PrintFolderTree(objSub, intLevel + 1, intMaxDepth)
        End If
    Next objSub
    On Error GoTo 0
End Sub


' ===========================================================================
' HILFSFUNKTIONEN
' ===========================================================================

' SMTP-Adresse aus Exchange-Adressbuch-Eintrag holen
Private Function GetSMTP(objEntry As Object) As String
    On Error Resume Next
    Dim strResult As String
    strResult = objEntry.SMTPAddress
    If Err.Number <> 0 Or strResult = "" Then
        strResult = objEntry.Address
        Err.Clear
    End If
    GetSMTP = strResult
End Function

' Dateinamen bereinigen und mit Zeitstempel versehen
Private Function CleanFileName(strSubject As String, dtReceived As Date) As String
    Dim s As String
    s = strSubject
    s = Replace(s, ":", "")
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, Chr(34), "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "|", "")
    s = Left(Trim(s), 50)
    CleanFileName = Format(dtReceived, "yyyymmdd_hhnn") & "_" & s & ".msg"
End Function

' Dateinamen einfach bereinigen (kein Zeitstempel, kein .msg)
Private Function CleanFileNameSimple(strInput As String) As String
    Dim s As String
    s = strInput
    s = Replace(s, ":", "")
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, Chr(34), "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "|", "")
    CleanFileNameSimple = Left(Trim(s), 80)
End Function

' Logging Helpers mit automatischem Zaehler und Fehlererklaerung
Private Sub Log_OK(strMsg As String, ByRef lngOK As Long)
    Debug.Print "  [ OK   ] " & strMsg
    lngOK = lngOK + 1
End Sub

Private Sub Log_FAIL(strMsg As String, ByRef lngFail As Long)
    Dim strHint As String
    strHint = ""
    If InStr(strMsg, "287 -") > 0 Then
        strHint = " --> OOM Security Guard (Programmatic Access blocked)"
    ElseIf InStr(strMsg, "429 -") > 0 Then
        strHint = " --> ActiveX-Komponente nicht registriert / nicht installiert"
    ElseIf InStr(strMsg, "-2147467259") > 0 Then
        strHint = " --> OOM Security Guard (HRESULT E_FAIL)"
    ElseIf InStr(strMsg, "462 -") > 0 Then
        strHint = " --> Remoteprozedur - Outlook evtl. nicht geoeffnet"
    ElseIf InStr(strMsg, "438 -") > 0 Then
        strHint = " --> Objekt unterstuetzt diese Eigenschaft/Methode nicht"
    End If
    Debug.Print "  [ FAIL ] " & strMsg & strHint
    lngFail = lngFail + 1
End Sub
