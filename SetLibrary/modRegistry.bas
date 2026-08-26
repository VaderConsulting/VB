Attribute VB_Name = "modRegistry"
Option Explicit

Enum RegHive ':( Missing Scope
    HKEY_CLASSES_ROOT = &H80000000
    HKCR = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKCU = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKLM = &H80000002
    HKEY_USERS = &H80000003
    HKUS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    HKCC = &H80000005
    HKEY_DYN_DATA = &H80000006
    HKDD = &H80000006
End Enum

Public Enum RegType
    REG_SZ = 1       'Unicode nul terminated string
    REG_BINARY = 3   'Free form binary
    REG_DWORD = 4    '32-bit number
End Enum

Public Const ERROR_SUCCESS = 0&
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : CopyRegByte
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function CopyRegByte(ByVal From_hKey As RegHive, ByVal From_strPath As String, ByVal From_strKeyName As String, ByVal To_strPath As String, Optional ByVal To_hKey As RegHive, Optional ByVal To_strKeyName As String) ':( As Variant ?

  Dim intLength As Integer
  Dim intLoop As Integer
  Dim lngResult As Long

    On Error GoTo CopyRegByte_Error

    If To_hKey = 0 Then
        To_hKey = From_hKey
      Else 'NOT TO_HKEY...
        To_hKey = To_hKey
    End If
    If To_strKeyName = "" Then
        To_strKeyName = From_strKeyName
      Else 'NOT TO_STRKEYNAME...
        To_strKeyName = To_strKeyName
    End If

  Dim mybytes As Variant ':( Move line to top of current Function
    mybytes = GetRegByte(From_hKey, From_strPath, From_strKeyName)
    intLength = UBound(mybytes)
  Dim x() As Byte ':( Move line to top of current Function
    ReDim x(intLength)
    For intLoop = 0 To UBound(mybytes)
        x(intLoop) = mybytes(intLoop)
    Next intLoop
    lngResult = SaveRegByte(To_hKey, To_strPath, To_strKeyName, x)

    On Error GoTo 0

Exit Function

CopyRegByte_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure CopyRegByte of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : CopyRegLong
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function CopyRegLong(ByVal hKey As RegHive, ByVal From_strPath As String, ByVal From_strKeyName As String, ByVal To_strPath As String, Optional ByVal To_hKey As RegHive, Optional ByVal From_hKey As RegHive, Optional ByVal To_strKeyName As String) ':( As Variant ?

  Dim lngResult As Long
  Dim mylong As Long
    On Error GoTo CopyRegLong_Error

    If To_hKey = 0 Then
        To_hKey = From_hKey
      Else 'NOT TO_HKEY...
        To_hKey = To_hKey
    End If
    If To_strKeyName = "" Then
        To_strKeyName = From_strKeyName
      Else 'NOT TO_STRKEYNAME...
        To_strKeyName = To_strKeyName
    End If

    mylong = GetRegLong(From_hKey, From_strPath, From_strKeyName)
    lngResult = SaveRegLong(To_hKey, To_strPath, To_strKeyName, mylong)

    On Error GoTo 0

Exit Function

CopyRegLong_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure CopyRegLong of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : CopyRegString
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function CopyRegString(ByVal From_hKey As RegHive, ByVal From_strPath As String, ByVal From_strKeyName As String, ByVal To_strPath As String, Optional ByVal To_hKey As RegHive, Optional ByVal To_strKeyName As String) ':( As Variant ?

  Dim lngResult As Long
  Dim mystring As String
    On Error GoTo CopyRegString_Error

    If To_hKey = 0 Then
        To_hKey = From_hKey
      Else 'NOT TO_HKEY...
        To_hKey = To_hKey
    End If
    If To_strKeyName = "" Then
        To_strKeyName = From_strKeyName
      Else 'NOT TO_STRKEYNAME...
        To_strKeyName = To_strKeyName
    End If

    mystring = GetRegString(From_hKey, From_strPath, From_strKeyName)
    lngResult = SaveRegString(To_hKey, To_strPath, To_strKeyName, mystring)

    On Error GoTo 0

Exit Function

CopyRegString_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure CopyRegString of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateRegKey
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function CreateRegKey(hKey As RegHive, strPath As String) ':( As Variant ?

  Dim hCurKey As Long
  Dim lRegResult As Long

    On Error GoTo CreateRegKey_Error

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If
    lRegResult = RegCloseKey(hCurKey)

    On Error GoTo 0

Exit Function

CreateRegKey_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateRegKey of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : DelRegKey
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function DelRegKey(ByVal hKey As RegHive, ByVal strPath As String) As Long

  Dim lRegResult As Long

    On Error GoTo DelRegKey_Error

    lRegResult = RegDeleteKey(hKey, strPath)
    DelRegKey = lRegResult

    On Error GoTo 0

Exit Function

DelRegKey_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure DelRegKey of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : DelRegValue
' DateTime  : 07-04-2003 08:08
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function DelRegValue(ByVal hKey As RegHive, ByVal strPath As String, ByVal strValue As String) ':( As Variant ?

  Dim hCurKey As Long
  Dim lRegResult As Long

    On Error GoTo DelRegValue_Error

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegDeleteValue(hCurKey, strValue)
    lRegResult = RegCloseKey(hCurKey)

    On Error GoTo 0

Exit Function

DelRegValue_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure DelRegValue of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRegByte
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetRegByte(ByVal hKey As RegHive, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant

  Dim lValueType As Long
  Dim byBuffer() As Byte
  Dim lDataBufferSize As Long
  Dim lRegResult As Long
  Dim hCurKey As Long

    On Error GoTo GetRegByte_Error

    If Not IsEmpty(Default) Then
        If VarType(Default) = vbArray + vbByte Then
            GetRegByte = Default
          Else 'NOT VARTYPE(DEFAULT)...
            GetRegByte = 0
        End If
      Else 'NOT NOT...
        GetRegByte = 0
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_BINARY Then
            ReDim byBuffer(lDataBufferSize - 1) As Byte
            lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
            GetRegByte = byBuffer
        End If
      Else 'NOT LREGRESULT...
        'there is a problem
    End If
    lRegResult = RegCloseKey(hCurKey)

    On Error GoTo 0

Exit Function

GetRegByte_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure GetRegByte of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRegLong
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetRegLong(ByVal hKey As RegHive, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long

  Dim lRegResult As Long
  Dim lValueType As Long
  Dim lBuffer As Long
  Dim lDataBufferSize As Long
  Dim hCurKey As Long

    On Error GoTo GetRegLong_Error

    'Set up default value
    If Not IsEmpty(Default) Then
        GetRegLong = Default
      Else 'NOT NOT...
        GetRegLong = 0
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lDataBufferSize = 4 '4 bytes = 32 bits = long
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetRegLong = lBuffer
        End If
      Else 'NOT LREGRESULT...
        'there is a problem
    End If
    lRegResult = RegCloseKey(hCurKey)

    On Error GoTo 0

Exit Function

GetRegLong_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure GetRegLong of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRegString
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetRegString(hKey As RegHive, strPath As String, strValue As String, Optional Default As String) As String

  Dim hCurKey As Long
  Dim lResult As Long
  Dim lValueType As Long
  Dim strBuffer As String
  Dim lDataBufferSize As Long
  Dim intZeroPos As Integer
  Dim lRegResult As Long

    On Error GoTo GetRegString_Error

    'Set up default value
    If Not IsEmpty(Default) Then
        GetRegString = Default
      Else 'NOT NOT...
        GetRegString = ""
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuffer = String$(lDataBufferSize, " ")
            lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetRegString = Left$(strBuffer, intZeroPos - 1)
              Else 'NOT INTZEROPOS...
                GetRegString = strBuffer
            End If
        End If
      Else 'NOT LREGRESULT...
        'there is a problem
    End If
    lRegResult = RegCloseKey(hCurKey)

    On Error GoTo 0

Exit Function

GetRegString_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure GetRegString of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRegSubKeyList
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetRegSubKeyList(ByVal hKey As RegHive, ByVal strPath As String) ':( As Variant ?

    On Error GoTo GetRegSubKeyList_Error

    On Error Resume Next
      Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long ':( Move line to top of current Function
      Dim sValue As String, td As Double, i As Long, Ret As Boolean, tmprst() ':( As Variant ?':( Move line to top of current Function
        Do Until Ret = True ':( Remove Pleonasm
            lResult = RegOpenKey(hKey, strPath, lKeyValue)
            sValue = Space$(2048)
            lValueLength = Len(sValue)
            lResult = RegEnumKey(lKeyValue, i, sValue, lValueLength)
            If (lResult = 0) And (Err.Number = 0) Then
                ReDim Preserve tmprst(i)
                tmprst(i) = Left$(sValue, InStr(sValue, Chr$(0)) - 1)
              Else 'NOT (LRESULT...
                Ret = True
            End If
            lResult = RegCloseKey(lKeyValue)
            i = i + 1
        Loop
        GetRegSubKeyList = tmprst

    On Error GoTo 0

Exit Function

GetRegSubKeyList_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure GetRegSubKeyList of Module modRegistry", Error

    End Function

'---------------------------------------------------------------------------------------
' Procedure : SaveRegByte
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function SaveRegByte(ByVal hKey As RegHive, ByVal strPath As String, ByVal strValueName As String, byData() As Byte)

  Dim lRegResult As Long
  Dim hCurKey As Long

    On Error GoTo SaveRegByte_Error

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)
    lRegResult = RegCloseKey(hCurKey)

    On Error GoTo 0

Exit Function

SaveRegByte_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveRegByte of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : SaveRegLong
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function SaveRegLong(ByVal hKey As RegHive, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long) ':( As Variant ?

  Dim hCurKey As Long
  Dim lRegResult As Long

    On Error GoTo SaveRegLong_Error

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)
    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If
    lRegResult = RegCloseKey(hCurKey)

    On Error GoTo 0

Exit Function

SaveRegLong_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveRegLong of Module modRegistry", Error

End Function

'---------------------------------------------------------------------------------------
' Procedure : SaveRegString
' DateTime  : 07-04-2003 08:09
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   07-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function SaveRegString(hKey As RegHive, strPath As String, strValue As String, strData As String) ':( As Variant ?

  Dim hCurKey As Long
  Dim lRegResult As Long

    On Error GoTo SaveRegString_Error

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If
    lRegResult = RegCloseKey(hCurKey)

    On Error GoTo 0

Exit Function

SaveRegString_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveRegString of Module modRegistry", Error

End Function

':) Ulli's VB Code Formatter V2.16.6 (2003-Jul-28 12:26) 32 + 511 = 543 Lines
