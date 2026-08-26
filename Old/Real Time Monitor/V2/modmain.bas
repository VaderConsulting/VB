Attribute VB_Name = "modMain"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function ReadIniFile(lpFileName As String, lpAppName As String, lpKeyName As String) As String
    '   lpfilename is similar to hive as in hkcu
    '   lpappname is similar to key in hkcu\software
    '   lpkeyname is similar to value in reg key
    '   lpreturnstring is the data being returned from the ini file
        
    Dim lpDefault$, lpReturnString$
    Dim Size%, Valid%, Path$, Succ%
    
    lpDefault$ = ""
    lpReturnString$ = Space$(128)
    Size% = Len(lpReturnString$)
    
    Valid% = GetPrivateProfileString(lpAppName$, lpKeyName$, "", lpReturnString$, Size%, lpFileName$)
    
    '* Discard the trailing spaces and null character, return as readinifile
    If Valid% > 0 Then
        ReadIniFile = Left$(lpReturnString$, Valid%)
    End If
End Function

