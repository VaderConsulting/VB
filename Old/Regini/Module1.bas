Attribute VB_Name = "Module1"
'Basic Module Code
Option Explicit
' module: modRegistry
' Description : This module Implements routines for manipulating the registry.

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngHKey As Long) As Long

Private Declare Function RegOpenKeyEx _
Lib "advapi32.dll" _
Alias "RegOpenKeyExA" _
(ByVal lngHKey As Long, _
ByVal lpSubKey As String, _
ByVal ulOptions As Long, _
ByVal samDesired As Long, _
phkResult As Long) _
As Long

Private Declare Function RegDeleteKey _
Lib "advapi32.dll" _
Alias "RegDeleteKeyA" _
(ByVal lngHKey As Long, _
ByVal lpSubKey As String) _
As Long

Private Declare Function RegDeleteValue _
Lib "advapi32.dll" _
Alias "RegDeleteValueA" _
(ByVal lngHKey As Long, _
ByVal lpValueName As String) _
As Long

Public Enum EnumRegistryRootKeys
    rrkHKeyClassesRoot = &H80000000
    rrkHKeyCurrentUser = &H80000001
    rrkHKeyLocalMachine = &H80000002
    rrkHKeyUsers = &H80000003
End Enum

Public Enum EnumRegistryValueType
    rrkRegSZ = 1
    rrkregBinary = 3
    rrkRegDWord = 4
End Enum

Private Const mcregOptionNonVolatile = 0
Private Const mcregErrorNone = 0
Private Const mcregErrorBadDB = 1
Private Const mcregErrorBadKey = 2
Private Const mcregErrorCantOpen = 3
Private Const mcregErrorCantRead = 4
Private Const mcregErrorCantWrite = 5
Private Const mcregErrorOutOfMemory = 6
Private Const mcregErrorInvalidParameter = 7
Private Const mcregErrorAccessDenied = 8
Private Const mcregErrorInvalidParameterS = 87
Private Const mcregErrorNoMoreItems = 259

Public Const mcregSynchronize = &H100000

Public Const mcregKeyQueryValue = &H1
Public Const mcregKeySetValue = &H2
Public Const mcregKeyCreateSubKey = &H4
Public Const mcregKeyEnumerateSubKeys = &H8
Public Const mcregKeyCreateLink = &H20
Public Const mcregKeyNotify = &H10
Public Const mcregReadControl = &H20000
Public Const mcregStandardRightsAll = &H1F0000
Public Const mcregStandardRightsRead = (mcregReadControl)
Public Const mcregStandardRightsWrite = (mcregReadControl)

Public Const mcregKeyAllAccess = ((mcregStandardRightsAll Or mcregKeyQueryValue Or mcregKeySetValue Or mcregKeyCreateSubKey Or mcregKeyEnumerateSubKeys Or mcregKeyNotify Or mcregKeyCreateLink) And (Not mcregSynchronize))
Public Const mcregKeyRead = ((mcregStandardRightsRead Or mcregKeyQueryValue Or mcregKeyEnumerateSubKeys Or mcregKeyNotify) And (Not mcregSynchronize))
Public Const mcregKeyWrite = ((mcregStandardRightsWrite Or mcregKeySetValue Or mcregKeyCreateSubKey) And (Not mcregSynchronize))

Public Sub RegistryDeleteKey(eRootKey As EnumRegistryRootKeys, strKeyName As String)
    ' Comments  : Deletes a key from the system registry
    ' Parameters: eRootKey - The root key
      '             strKeyName - The name of the key to delete
      ' Returns   : Nothing
    Dim lngRetVal As Long
    On Error GoTo PROC_ERR
    ' Delete the key
    lngRetVal = RegDeleteKey(eRootKey, strKeyName)
        
PROC_EXIT:
    Exit Sub
PROC_ERR:
    'MsgBox "Error: " & Err.Number & ". " & Err.Description, , "RegistryDeleteKey"
    Resume PROC_EXIT
End Sub

Public Sub RegistryDeleteValue(eRootKey As EnumRegistryRootKeys, strKeyName As String, strValueName As String)
    ' Comments  : Deletes a value from the system registry
    ' Parameters: eRootKey - The root key
      '             strKeyName - The name of the key to delete
      '             strValueName - The name of the value to delete
      ' Returns   : Nothing
    Dim lngRetVal As Long
    Dim lngHKey As Long
    
    On Error GoTo PROC_ERR
    ' Open the key
    lngRetVal = RegOpenKeyEx(eRootKey, strKeyName, 0, mcregKeyWrite, lngHKey)
    ' If the key was opened successfully, then delete it
    If lngRetVal = mcregErrorNone Then
        lngRetVal = RegDeleteValue(lngHKey, strValueName)
    End If
PROC_EXIT:
    Exit Sub
PROC_ERR:
    'MsgBox "Error: " & Err.Number & ". " & Err.Description, , "RegistryDeleteValue"
    Resume PROC_EXIT
End Sub


