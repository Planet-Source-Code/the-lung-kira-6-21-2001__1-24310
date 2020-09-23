Attribute VB_Name = "winreg"
Option Explicit


Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As Any, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByRef lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByVal lpReserved As Long, ByRef lpcSubKeys As Long, ByRef lpcbMaxSubKeyLen As Long, ByRef lpcbMaxClassLen As Long, ByRef lpcValues As Long, ByRef lpcbMaxValueNameLen As Long, ByRef lpcbMaxValueLen As Long, ByRef lpcbSecurityDescriptor As Long, ByRef lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006


Public Function CreateKey(ByVal hKey As Long, ByVal strPath As String, Optional ByRef retDisposition As Long)
    Dim retKey As Long
    
    apiError = RegCreateKeyEx(hKey, strPath & Chr$(0), 0, Chr$(0), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal &H0, retKey, retDisposition): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegCreateKeyEx"
    apiError = RegCloseKey(retKey): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegCloseKey"
End Function

Public Function DeleteKey(ByVal hKey As Long, ByVal strPath As String)
    apiError = RegDeleteKey(hKey, strPath): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegDeleteKey"
End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String)
    Dim hCurKey As Long
    
    apiError = RegOpenKeyEx(hKey, strPath, &H0, KEY_SET_VALUE, hCurKey): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegOpenKeyEx"
    apiError = RegDeleteValue(hCurKey, strValueName): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegDeleteValue"
    apiError = RegCloseKey(hCurKey): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegCloseKey"
End Function

Public Function EnumValue(ByVal hKey As Long, ByVal strPath As String, ByRef strValueName() As String, ByRef strData() As String, ByRef lngDataType() As Long, ByRef lngCount As Long)
    Dim hCurKey As Long
    Dim lenMaxValueName As Long
    Dim FILETIME As FILETIME
    
    Dim lngValueType As Long
    Dim lngValue As Long
    Dim lenData As Long
    
    Do
        apiError = RegOpenKeyEx(hKey, strPath, &H0, KEY_READ, hCurKey): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegOpenKeyEx"
        apiError = RegQueryInfoKey(hCurKey, &H0, &H0, &H0, &H0, &H0, &H0, &H0, lenMaxValueName, &H0, &H0, FILETIME): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegQueryValueEx"
        
        ReDim Preserve strValueName(lngCount)
        strValueName(lngCount) = String$(lenMaxValueName + 1, &H0)
        lngValue = Len(strValueName(lngCount))
        
        ReDim Preserve strData(lngCount)
        strData(lngCount) = String$(4096, &H0)
        ReDim Preserve lngDataType(lngCount)
        lenData = 4096
        
        apiError = RegEnumValue(hCurKey, lngCount, strValueName(lngCount), lngValue, &H0, lngDataType(lngCount), ByVal strData(lngCount), lenData)
        If apiError <> 234 Then
        If apiError > 0 Then
            Exit Do
        End If
        End If
        
        If lngValue > 0 Then
            strValueName(lngCount) = Fix_NullTermStr(strValueName(lngCount))
        End If
        If lenData > 0 Then
            strData(lngCount) = Fix_NullTermStr(Left$(strData(lngCount), lenData))
        End If
        
        lngCount = lngCount + 1
        RegCloseKey hCurKey
    Loop While Not apiError = ERROR_NO_MORE_ITEMS

    If apiError <> ERROR_NO_MORE_ITEMS Then
        Errors apiError, "RegEnumValue"
    End If
End Function

Public Function GetRegSetting(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String) As Variant
    Dim hCurKey As Long
    Dim lngDataBufferSize As Long
    Dim lngValueType As Long
    
    apiError = RegOpenKeyEx(hKey, strPath, 0, KEY_QUERY_VALUE, hCurKey): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegOpenKeyEx"
    apiError = RegQueryValueEx(hCurKey, strValue, &H0, lngValueType, ByVal &H0, lngDataBufferSize): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegQueryValueEx"
    
    Select Case lngValueType
        Case REG_BINARY
            Dim strBuffer As String
            strBuffer = String$(lngDataBufferSize, &H0)
            
            apiError = RegQueryValueEx(hCurKey, strValue, &H0, lngValueType, ByVal strBuffer, lngDataBufferSize): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegQueryValueEx"
            
            GetRegSetting = strBuffer
        Case REG_DWORD
            Dim lngBuffer As Long
            lngDataBufferSize = 4
            
            apiError = RegQueryValueEx(hCurKey, strValue, &H0, lngValueType, lngBuffer, lngDataBufferSize): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegQueryValueEx"
            GetRegSetting = CDbl(lngBuffer)
        Case REG_SZ
            strBuffer = String$(lngDataBufferSize, &H0)
            
            apiError = RegQueryValueEx(hCurKey, strValue, &H0, lngValueType, ByVal strBuffer, lngDataBufferSize): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegQueryValueEx"
            GetRegSetting = Fix_NullTermStr(strBuffer)
    End Select
    
    apiError = RegCloseKey(hCurKey): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegCloseKey"
End Function

Public Sub SaveRegSetting(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal varData As Variant, ByVal datType As Long, Optional ByRef retDisposition As Long)
    Dim retKey As Long
    
    apiError = RegCreateKeyEx(hKey, strPath, 0, Chr$(0), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal &H0, retKey, retDisposition): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegCreateKeyEx"
    
    Select Case datType
        Case REG_BINARY
            apiError = RegSetValueEx(retKey, strValue, &H0, datType, ByVal CStr(varData), Len(varData)): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegSetValueEx"
        Case REG_DWORD
            apiError = RegSetValueEx(retKey, strValue, &H0, datType, CLng(varData), 4): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegSetValueEx"
        Case REG_SZ
            apiError = RegSetValueEx(retKey, strValue, &H0, datType, ByVal CStr(varData), Len(varData)): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegSetValueEx"
    End Select
    
    apiError = RegCloseKey(retKey): If apiError <> ERROR_SUCCESS Then Errors apiError, "RegCloseKey"
End Sub
