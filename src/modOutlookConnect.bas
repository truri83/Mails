Attribute VB_Name = "modOutlookConnect"
Option Compare Database
Option Explicit

' ===========================================================================
' modOutlookConnect - Outlook/Redemption Verbindungsverwaltung
' ===========================================================================
' Stellt sicher, dass Outlook und Redemption korrekt initialisiert sind.
' Bietet SMTP-Adress-Aufloesung fuer interne Exchange-Adressen.
'
' Funktionen:
'   ConnectOutlook()        -> Outlook.Application initialisieren
'   ConnectRDO()            -> Redemption.RDOSession initialisieren + Logon
'   DisconnectAll()         -> Alle Verbindungen trennen
'   ErstelleSafeMail()      -> SafeMailItem aus OOM-MailItem erstellen
'   GetSMTPFromEntry()      -> SMTP aus AddressEntry/Sender aufloesen
'   GetSMTPFromRecipient()  -> SMTP aus Recipient-Objekt aufloesen
'   GetAbsenderSMTP()       -> Absender-SMTP einer RDO-Mail ermitteln
'   OeffneOrdner()          -> Outlook-Ordner per Pfad oeffnen
' ===========================================================================


' ---------------------------------------------------------------------------
' OUTLOOK.APPLICATION INITIALISIEREN (OOM)
' ---------------------------------------------------------------------------
Public Function ConnectOutlook() As Boolean
    On Error GoTo ErrHandler

    If Not g_objOutlook Is Nothing Then
        ' Pruefen ob Verbindung noch aktiv
        Dim strVer As String
        On Error Resume Next
        strVer = g_objOutlook.Version
        If Err.Number = 0 Then
            ConnectOutlook = True
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        Set g_objOutlook = Nothing
        On Error GoTo ErrHandler
    End If

    Set g_objOutlook = CreateObject("Outlook.Application")
    If g_objOutlook Is Nothing Then
        LogError "Outlook.Application konnte nicht erstellt werden", "CONNECT"
        ConnectOutlook = False
    Else
        LogInfo "Outlook " & g_objOutlook.Version & " verbunden", "CONNECT"
        ConnectOutlook = True
    End If
    Exit Function

ErrHandler:
    LogVBAError "ConnectOutlook"
    ConnectOutlook = False
End Function


' ---------------------------------------------------------------------------
' REDEMPTION RDOSESSION INITIALISIEREN
' ---------------------------------------------------------------------------
Public Function ConnectRDO() As Boolean
    On Error GoTo ErrHandler

    If Not g_objRDO Is Nothing Then
        ' Pruefen ob Session noch gueltig
        On Error Resume Next
        Dim strVer As String
        strVer = g_objRDO.Version
        If Err.Number = 0 Then
            ConnectRDO = True
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        Set g_objRDO = Nothing
        On Error GoTo ErrHandler
    End If

    Set g_objRDO = CreateObject("Redemption.RDOSession")
    If g_objRDO Is Nothing Then
        LogError "Redemption.RDOSession nicht verfuegbar. " & _
                 "Registrierung: regsvr32 """ & RDO_DLL_64 & """", "CONNECT"
        ConnectRDO = False
        Exit Function
    End If

    g_objRDO.Logon
    LogInfo "Redemption " & g_objRDO.Version & " verbunden", "CONNECT"
    ConnectRDO = True
    Exit Function

ErrHandler:
    LogVBAError "ConnectRDO"
    ConnectRDO = False
End Function


' ---------------------------------------------------------------------------
' ALLE VERBINDUNGEN TRENNEN
' ---------------------------------------------------------------------------
Public Sub DisconnectAll()
    On Error Resume Next

    If Not g_objRDO Is Nothing Then
        g_objRDO.Logoff
        Set g_objRDO = Nothing
        LogDebug "RDO Logoff", "CONNECT"
    End If

    Set g_objOutlook = Nothing
    LogDebug "Outlook freigegeben", "CONNECT"

    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' REDEMPTION SAFEMAILITEM ERSTELLEN (OOM-Wrapper)
' ---------------------------------------------------------------------------
Public Function ErstelleSafeMail(objOOMMail As Object) As Object
    On Error GoTo ErrHandler

    If objOOMMail Is Nothing Then
        Set ErstelleSafeMail = Nothing
        Exit Function
    End If

    Dim objSafe As Object
    Set objSafe = CreateObject("Redemption.SafeMailItem")
    objSafe.Item = objOOMMail
    Set ErstelleSafeMail = objSafe
    Exit Function

ErrHandler:
    LogVBAError "ErstelleSafeMail"
    Set ErstelleSafeMail = Nothing
End Function


' ---------------------------------------------------------------------------
' SMTP-ADRESSE AUS ADDRESSENTRY / SENDER-OBJEKT HOLEN
' Loest interne Exchange-Adressen (/O=BWL/...) korrekt auf
' ---------------------------------------------------------------------------
Public Function GetSMTPFromEntry(objEntry As Object) As String
    On Error Resume Next
    Dim strResult As String

    If objEntry Is Nothing Then
        GetSMTPFromEntry = ""
        Exit Function
    End If

    ' Versuch 1: Direkt SMTPAddress
    strResult = objEntry.SMTPAddress
    If Err.Number <> 0 Then Err.Clear: strResult = ""

    ' Versuch 2: GetExchangeUser (fuer Exchange-Eintraege)
    If strResult = "" Or Left(strResult, 3) = "/O=" Then
        Dim objExUser As Object
        Set objExUser = objEntry.GetExchangeUser
        If Not objExUser Is Nothing Then
            strResult = objExUser.PrimarySmtpAddress
            If Err.Number <> 0 Then Err.Clear: strResult = ""
        End If
    End If

    ' Versuch 3: Fallback auf .Address
    If strResult = "" Or Left(strResult, 3) = "/O=" Then
        strResult = objEntry.Address
        If Err.Number <> 0 Then Err.Clear: strResult = ""
    End If

    On Error GoTo 0
    GetSMTPFromEntry = strResult
End Function


' ---------------------------------------------------------------------------
' SMTP-ADRESSE AUS REDEMPTION RECIPIENT HOLEN
' ---------------------------------------------------------------------------
Public Function GetSMTPFromRecipient(objRecipient As Object) As String
    On Error Resume Next
    Dim strResult As String
    Dim objAE As Object

    Set objAE = objRecipient.AddressEntry
    If Not objAE Is Nothing Then
        strResult = objAE.SMTPAddress
        If Err.Number <> 0 Or strResult = "" Then
            Err.Clear
            strResult = objAE.Address
        End If
    End If

    If strResult = "" Then strResult = objRecipient.Address
    If Err.Number <> 0 Then Err.Clear

    On Error GoTo 0
    GetSMTPFromRecipient = Nz(strResult, "")
End Function


' ---------------------------------------------------------------------------
' ABSENDER-SMTP EINER RDO-MAIL ERMITTELN
' Behandelt sowohl SMTP- als auch EX-Absender korrekt
' ---------------------------------------------------------------------------
Public Function GetAbsenderSMTP(objRDOMail As Object) As String
    On Error Resume Next
    Dim strResult As String

    ' Pruefen ob Exchange-Adresse
    If objRDOMail.SenderEmailType = "EX" Then
        ' Ueber Sender-Objekt aufloesen
        strResult = GetSMTPFromEntry(objRDOMail.Sender)
    Else
        strResult = objRDOMail.SenderEmailAddress
    End If

    If Err.Number <> 0 Then Err.Clear: strResult = ""

    ' Fallback: MAPI Property PR_SMTP_ADDRESS
    If strResult = "" Or Left(strResult, 3) = "/O=" Then
        strResult = objRDOMail.Fields(PR_SMTP_ADDRESS)
        If Err.Number <> 0 Then Err.Clear: strResult = ""
    End If

    On Error GoTo 0
    If strResult = "" Then strResult = DEFAULT_EMAIL
    GetAbsenderSMTP = strResult
End Function


' ---------------------------------------------------------------------------
' OUTLOOK-ORDNER ANHAND PFAD OEFFNEN (via RDO)
' Format: "Postfachname\Ordner\Unterordner"
' ---------------------------------------------------------------------------
Public Function OeffneOrdner(ByVal strPfad As String) As Object
    On Error GoTo ErrHandler

    If Not ConnectRDO() Then
        Set OeffneOrdner = Nothing
        Exit Function
    End If

    ' Pfad bereinigen
    If Left(strPfad, 2) = "\\" Then strPfad = Mid(strPfad, 3)
    strPfad = Replace(strPfad, "\\", "\")

    Dim arrTeile() As String
    arrTeile = Split(strPfad, "\")

    If UBound(arrTeile) < 0 Then
        Set OeffneOrdner = Nothing
        Exit Function
    End If

    ' Store anhand des ersten Pfadteils finden
    Dim objStore As Object
    Dim objFolder As Object
    Dim i As Long

    For Each objStore In g_objRDO.Stores
        If objStore.DisplayName = arrTeile(0) Then
            Set objFolder = objStore.RootFolder
            Exit For
        End If
    Next objStore

    If objFolder Is Nothing Then
        LogWarn "Store '" & arrTeile(0) & "' nicht gefunden", "CONNECT"
        Set OeffneOrdner = Nothing
        Exit Function
    End If

    ' Durch Unterordner navigieren
    For i = 1 To UBound(arrTeile)
        If Trim(arrTeile(i)) <> "" Then
            On Error Resume Next
            Set objFolder = objFolder.Folders(arrTeile(i))
            On Error GoTo ErrHandler
            If objFolder Is Nothing Then
                LogWarn "Ordner '" & arrTeile(i) & "' nicht gefunden in: " & strPfad, "CONNECT"
                Set OeffneOrdner = Nothing
                Exit Function
            End If
        End If
    Next i

    Set OeffneOrdner = objFolder
    Exit Function

ErrHandler:
    LogVBAError "OeffneOrdner(" & strPfad & ")"
    Set OeffneOrdner = Nothing
End Function
