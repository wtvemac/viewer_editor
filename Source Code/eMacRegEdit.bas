Attribute VB_Name = "eMacRegEdit"
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

Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&
Private Const MAX_PATH = 256&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = &H20000
Private Const STANDARD_RIGHTS_WRITE = &H20000
Private Const STANDARD_RIGHTS_EXECUTE = &H20000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SEDataValue = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Public SubKColl As Collection
Public ValColl As Collection
Public ValTypeColl As Collection
Public HasSubKeys() As Boolean
Public theRE As New RegExp
Public REMatches As MatchCollection
Public REMatch As Match


Public Sub CreateKey(hKey As Long, strPath As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function DeleteKey(ByVal hKey As Long, ByVal strPath As String) As Boolean
    Dim lRegResult As Long
    lRegResult = RegDeleteKey(hKey, strPath)
    If lRegResult = 0 Then
        DeleteKey = True
    Else
        DeleteKey = False
    End If
End Function
Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegDeleteValue(hCurKey, strValue)
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    'Upgraded to read REG_EXPAND_SZ
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    If Not IsEmpty(Default) Then
        GetSettingString = Default
    Else
        GetSettingString = ""
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Or REG_EXPAND_SZ Then
            strBuffer = String(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingString = Left$(strBuffer, intZeroPos - 1)
            Else
                GetSettingString = strBuffer
            End If
            If lValueType = REG_EXPAND_SZ Then GetSettingString = StripTerminator(ExpandEnvStr(GetSettingString))
        End If
    Else
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long
    Dim lRegResult As Long
    Dim lValueType As Long
    Dim lBuffer As Long
    Dim lDataBufferSize As Long
    Dim hCurKey As Long
    If Not IsEmpty(Default) Then
        GetSettingLong = Default
    Else
        GetSettingLong = 0
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lDataBufferSize = 4
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetSettingLong = lBuffer
        End If
    Else
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant
    Dim lValueType As Long
    Dim byBuffer() As Byte
    Dim lDataBufferSize As Long
    Dim lRegResult As Long
    Dim hCurKey As Long
    ReDim byBuffer(0 To 1) As Byte
    byBuffer(0) = 0
    If Not IsEmpty(Default) Then
        If VarType(Default) = vbArray + vbByte Then
            GetSettingByte = Default
        Else
            GetSettingByte = 0
        End If
    Else
        GetSettingByte = 0
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_BINARY Then
            If lDataBufferSize = 0 Then
                ReDim byBuffer(0) As Byte
                byBuffer(0) = 0
                GetSettingByte = byBuffer
            Else
                ReDim byBuffer(lDataBufferSize - 1) As Byte
                lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
                GetSettingByte = byBuffer
            End If
        End If
    Else
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, byData() As Byte)
    Dim lRegResult As Long
    Dim hCurKey As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Sub SaveSettingEmptyByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String)
    Dim lRegResult As Long
    Dim hCurKey As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, 0&, 0&)
    lRegResult = RegCloseKey(hCurKey)
End Sub


Public Function GetAllKeys(hKey As Long, strPath As String) As Variant
    'Modified by Ian Northwood - thanks Ian
    Dim lRegResult As Long
    Dim lCounter As Long
    Dim hCurKey As Long
    Dim hCurKey2 As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim strnames() As String
    Dim temp As String
    Dim intZeroPos As Integer
    Dim strDummy As String
    
    If Len(strPath) > 0 Then temp = "\"
    lCounter = 0
    ReDim strnames(lCounter) As String
    strnames(lCounter) = "  "
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    Do
        DoEvents
        lDataBufferSize = 255
        strBuffer = String(lDataBufferSize, " ")
        lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)
        If lRegResult = ERROR_SUCCESS Then
            ReDim Preserve strnames(lCounter) As String
            ReDim Preserve HasSubKeys(lCounter) As Boolean
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                strnames(UBound(strnames)) = Left$(strBuffer, intZeroPos - 1)
            Else
                strnames(UBound(strnames)) = strBuffer
            End If
            
            If Right$(strPath, 1) = "\" Then
                strDummy = strPath + strnames(lCounter)
            Else
                strDummy = strPath + temp + strnames(lCounter)
            End If

            lRegResult = RegOpenKey(hKey, strDummy, hCurKey2)
            lDataBufferSize = 255
            strBuffer = String(lDataBufferSize, " ")
            lRegResult = RegEnumKey(hCurKey2, 0, strBuffer, lDataBufferSize)
            If lRegResult = ERROR_SUCCESS Then
                HasSubKeys(UBound(HasSubKeys)) = True
            Else
                HasSubKeys(UBound(HasSubKeys)) = False
            End If
            lCounter = lCounter + 1
        Else
            Exit Do
        End If
    Loop
    GetAllKeys = strnames
End Function


Public Function CountAllKeys(hKey As Long) As Boolean
    Dim lRegResult As Long
    Dim lCounter As Long
    Dim hCurKey As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim strnames() As String
    Dim intZeroPos As Integer
    lCounter = 0
    lRegResult = RegOpenKey(hKey, "", hCurKey)
    Do
    lDataBufferSize = 255
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        ReDim Preserve strnames(lCounter) As String
        intZeroPos = InStr(strBuffer, Chr$(0))
        If intZeroPos > 0 Then
            strnames(UBound(strnames)) = Left$(strBuffer, intZeroPos - 1)
            lCounter = lCounter + 1
        Else
            strnames(UBound(strnames)) = strBuffer
            lCounter = lCounter + 1
        End If
    Else
        Exit Do
    End If
    If lCounter > 0 Then
    CountAllKeys = True
    Exit Do
    End If
Loop
End Function





Private Function ExpandEnvStr(sData As String) As String
    'This is cool - borrowed this
    Dim c As Long, s As String
    s = ""
    c = ExpandEnvironmentStrings(sData, s, c)
    s = String$(c - 1, 0)
    c = ExpandEnvironmentStrings(sData, s, c)
    ExpandEnvStr = s
End Function

Public Function checkRE(chkString As String, chkRE As String)

    theRE.Pattern = chkRE
    Set REMatches = theRE.Execute(chkString)
   
    If REMatches.count > 0 Then
        Set REMatch = REMatches.Item(0)
        checkRE = 1
    Else
        Set REMatch = Nothing
        checkRE = 0
    End If

End Function



