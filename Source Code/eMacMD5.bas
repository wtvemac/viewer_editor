Attribute VB_Name = "eMacMD5"
'''''''''''''''''''''''''''''''''
' WebTV IPE (In-place Edit) 4.0 '
'                               '
' By: Eric MacDonald            '
' Date: April 24, 2005          '
'                               '
' This is a patcher tool        '
' for any SuperViewer template  '
'''''''''''''''''''''''''''''''''

Option Explicit



Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long

Private Const SERVICE_PROVIDER As String = "Microsoft Enhanced Cryptographic Provider v1.0" & vbNullChar
Private Const KEY_CONTAINER As String = "GCN SSL Container" & vbNullChar
Private Const HP_HASHVAL As Long = 2
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT = &HF0000000
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const CALG_MD2 As Long = 32769
Private Const CALG_MD4 As Long = 32770
Private Const CALG_MD5 As Long = 32771
Private Const CALG_SHA1 As Long = 32772

Public Function HashIt(lngAlg As Long, ByVal strData As String, Optional blnHexOutput As Boolean = False) As String
    Dim TheAnswer As Long
    Dim lngReturnValue As Long
    Dim strHash As String
    Dim hCryptProv As Long
    Dim hHash As Long
    Dim lngHashLen As Long
    Dim bytActiveChar As Byte
    Dim strTmp As String
    Dim i As Integer
    lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET)
    If lngReturnValue = 0 Then
        lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0)
        If lngReturnValue = 0 Then Exit Function
    End If
    lngReturnValue = CryptCreateHash(hCryptProv, lngAlg, 0, 0, hHash)
    lngReturnValue = CryptHashData(hHash, strData, Len(strData), 0)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0)
    strHash = String(lngHashLen, vbNullChar)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0)
    If hHash <> 0 Then CryptDestroyHash hHash
    If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0
    If blnHexOutput = True Then
        For i = 1 To Len(strHash)
            bytActiveChar = Asc(Mid$(strHash, i, 1))
            strTmp = strTmp & IIf(bytActiveChar > 15, Hex(bytActiveChar), "0" & Hex(bytActiveChar))
        Next i
        HashIt = strTmp
    Else
        HashIt = strHash
    End If
End Function

Public Function MD2(ByVal strData As String, Optional blnHexOutput As Boolean = False) As String
    MD2 = HashIt(CALG_MD2, strData, blnHexOutput)
End Function

Public Function MD4(ByVal strData As String, Optional blnHexOutput As Boolean = False) As String
    MD4 = HashIt(CALG_MD4, strData, blnHexOutput)
End Function

Public Function MD5(ByVal strData As String, Optional blnHexOutput As Boolean = False) As String
    MD5 = HashIt(CALG_MD5, strData, blnHexOutput)
End Function

Public Function SHA1(ByVal strData As String, Optional blnHexOutput As Boolean = False) As String
    SHA1 = HashIt(CALG_SHA1, strData, blnHexOutput)
End Function
