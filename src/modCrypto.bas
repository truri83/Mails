Attribute VB_Name = "modCrypto"
Option Compare Database
Option Explicit

' ===========================================================================
' modCrypto - SHA256-Hash via Windows CryptoAPI
' ===========================================================================
' Berechnet SHA256-Hashes fuer Duplikat-Erkennung bei Mail-Import.
' Nutzt ausschliesslich die Windows CryptoAPI (advapi32.dll).
' Kein externer Verweis noetig.
'
' Funktionen:
'   SHA256_Hash(strInput)                          -> 64 Zeichen Hex-String
'   GeneriereMailHash(Betreff, Absender, Empf, Dt) -> Hash fuer Duplikat-Check
' ===========================================================================

' ---------------------------------------------------------------------------
' API-DEKLARATIONEN (32/64-Bit kompatibel)
' ---------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" _
        Alias "CryptAcquireContextA" ( _
        ByRef phProv As LongPtr, ByVal pszContainer As String, _
        ByVal pszProvider As String, ByVal dwProvType As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, _
        ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long

    Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, ByRef pbData As Byte, ByVal dwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Byte, _
        ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" ( _
        ByVal hHash As LongPtr) As Long

    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
        ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function CryptAcquireContext Lib "advapi32.dll" _
        Alias "CryptAcquireContextA" ( _
        ByRef phProv As Long, ByVal pszContainer As String, _
        ByVal pszProvider As String, ByVal dwProvType As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
        ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, _
        ByVal dwFlags As Long, ByRef phHash As Long) As Long

    Private Declare Function CryptHashData Lib "advapi32.dll" ( _
        ByVal hHash As Long, ByRef pbData As Byte, ByVal dwDataLen As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
        ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Byte, _
        ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long

    Private Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
        ByVal hHash As Long) As Long

    Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
        ByVal hProv As Long, ByVal dwFlags As Long) As Long
#End If

' CryptoAPI-Konstanten
Private Const PROV_RSA_AES      As Long = 24
Private Const CALG_SHA_256      As Long = 32780    ' &H800C
Private Const HP_HASHVAL        As Long = 2
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000


' ---------------------------------------------------------------------------
' SHA256-Hash aus String berechnen
' Rueckgabe: 64 Zeichen Hex-String (lowercase), oder "" bei Fehler
' ---------------------------------------------------------------------------
Public Function SHA256_Hash(ByVal strInput As String) As String
    On Error GoTo ErrHandler

    #If VBA7 Then
        Dim hProv As LongPtr, hHash As LongPtr
    #Else
        Dim hProv As Long, hHash As Long
    #End If

    Dim abData()    As Byte
    Dim abHash()    As Byte
    Dim lngHashLen  As Long
    Dim i           As Long
    Dim strResult   As String

    If Len(strInput) = 0 Then
        SHA256_Hash = ""
        Exit Function
    End If

    ' String in Byte-Array konvertieren (ANSI)
    abData = StrConv(strInput, vbFromUnicode)

    ' Kryptografie-Provider oeffnen
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) = 0 Then
        SHA256_Hash = "": Exit Function
    End If

    ' Hash-Objekt erstellen
    If CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash) = 0 Then
        CryptReleaseContext hProv, 0
        SHA256_Hash = "": Exit Function
    End If

    ' Daten hashen
    If CryptHashData(hHash, abData(0), UBound(abData) + 1, 0) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        SHA256_Hash = "": Exit Function
    End If

    ' Hash-Laenge ermitteln und Ergebnis lesen
    CryptGetHashParam hHash, HP_HASHVAL, ByVal 0&, lngHashLen, 0
    ReDim abHash(lngHashLen - 1)
    CryptGetHashParam hHash, HP_HASHVAL, abHash(0), lngHashLen, 0

    ' In Hex-String umwandeln (mit fuehrender Null pro Byte!)
    strResult = ""
    For i = LBound(abHash) To UBound(abHash)
        strResult = strResult & Right("0" & Hex(abHash(i)), 2)
    Next i

    ' Ressourcen freigeben
    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0

    SHA256_Hash = LCase(strResult)
    Exit Function

ErrHandler:
    On Error Resume Next
    If hHash <> 0 Then CryptDestroyHash hHash
    If hProv <> 0 Then CryptReleaseContext hProv, 0
    SHA256_Hash = ""
End Function


' ---------------------------------------------------------------------------
' Mail-Hash generieren (fuer Duplikat-Erkennung)
' Eingabe: Betreff + Absender-Email + Empfaenger(To) + Empfangsdatum
' Rueckgabe: SHA256-Hex-String (64 Zeichen)
' ---------------------------------------------------------------------------
Public Function GeneriereMailHash(ByVal strBetreff As String, _
                                  ByVal strAbsender As String, _
                                  ByVal strEmpfaenger As String, _
                                  ByVal dtEmpfangen As Date) As String
    Dim strRaw As String
    strRaw = Nz(strBetreff, "") & "|" & _
             Nz(strAbsender, "") & "|" & _
             Nz(strEmpfaenger, "") & "|" & _
             Format(dtEmpfangen, "yyyymmddhhnnss")
    GeneriereMailHash = SHA256_Hash(strRaw)
End Function
