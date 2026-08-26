Attribute VB_Name = "modGeneral"
Public Declare Function NetApiBufferFree Lib "netapi32.dll" (BufPtr As Any) As Long
Public Declare Function NetApiBufferAllocate Lib "netapi32.dll" (ByVal ByteCount As Long, Ptr As Long) As Long
Public Declare Function lstrcpyW Lib "kernel32" (ByVal lpszDest As String, ByVal lpszSrc As Long) As Long
Public Declare Function StrLenA Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Public Declare Function StrLenW Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Public Declare Function StrCopyA Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lppEnumHwnd As Long) As Long
Public Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal pEnumHwnd As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Public Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal p_lngEnumHwnd As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)

Public Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type

Public Function PointerToString(lpszString As Long) As String
    Dim lpszStr1 As String, lpszStr2 As String, nRes As Long
    lpszStr1 = String(1000, "*")
    nRes = lstrcpyW(lpszStr1, lpszString)
    lpszStr2 = (StrConv(lpszStr1, vbFromUnicode))
    PointerToString = Left(lpszStr2, InStr(lpszStr2, Chr$(0)) - 1)
End Function

Public Function PointerToUnicodeStr(lpUnicodeStr As Long) As String
    
    On Error Resume Next ' Don't accept an error here
    
    Dim Buffer() As Byte
    Dim nLen As Long
    
    If lpUnicodeStr Then
        nLen = StrLenW(lpUnicodeStr) * 2
        If nLen Then
            ReDim Buffer(0 To (nLen - 1)) As Byte
            
            ' ------------------------------------
            ' Copy the pointer to the buffer into
            ' the type array
            ' ------------------------------------
            CopyMem Buffer(0), ByVal lpUnicodeStr, nLen
            PointerToUnicodeStr = Buffer
            
        End If
    
    End If

End Function

