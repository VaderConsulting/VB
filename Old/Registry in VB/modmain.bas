Attribute VB_Name = "modMain"

Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" _
      (ByVal lpMachineName As String, _
       ByVal hKey As Long, phkResult As Long) _
       As Long
       
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
      (ByVal hKey As Long, _
       ByVal lpValueName As String, _
       ByVal lpReserved As Long, _
       lpType As Long, _
       lpData As Any, _
       lpcbData As Long) _
       As Long
      
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
      (ByVal hKey As Long, _
       ByVal lpSubKey As String, _
       ByVal dwReserved As Long, _
       ByVal samDesired As Long, _
       phkResult As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' Registry manipulation API's (32-bit)
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_NO_MORE_ITEMS = 259&

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const KEY_ALL_ACCESS As Long = &HF0063

