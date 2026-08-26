Attribute VB_Name = "basMain"
Option Explicit

Public Declare Function SendMessageLong _
    Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LEVEL_NETWORK = 1
Public Const LEVEL_DOMAIN = 2
Public Const LEVEL_SERVER = 3
Public Const LEVEL_SHARE = 4
Public Const LEVEL_DIRECTORY = 5
Public Const LEVEL_FILE = 6

Public Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type

Public Declare Function WNetOpenEnum _
    Lib "mpr.dll" _
    Alias "WNetOpenEnumA" ( _
    ByVal dwScope As Long, _
    ByVal dwType As Long, _
    ByVal dwUsage As Long, _
    lpNetResource As Any, _
    lphEnum As Long) As Long

Public Declare Function WNetEnumResource _
    Lib "mpr.dll" _
    Alias "WNetEnumResourceA" ( _
    ByVal hEnum As Long, _
    lpcCount As Long, _
    ByVal lpBuffer As Long, _
    lpBufferSize As Long) As Long

Public Declare Function WNetCloseEnum _
    Lib "mpr.dll" ( _
    ByVal hEnum As Long) As Long

'RESOURCE ENUMERATION.
      
Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCE_REMEMBERED = &H3

Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_UNKNOWN = &HFFFF

Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
Public Const RESOURCEUSAGE_RESERVED = &H80000000

Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEDISPLAYTYPE_FILE = &H4
Public Const RESOURCEDISPLAYTYPE_GROUP = &H5
Public Const RESOURCEDISPLAYTYPE_NETWORK = &H6
Public Const RESOURCEDISPLAYTYPE_ROOT = &H7
Public Const RESOURCEDISPLAYTYPE_ADMINSHARE = &H8
Public Const RESOURCEDISPLAYTYPE_DIRECTORY = &H9

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function GlobalAlloc _
    Lib "kernel32" ( _
    ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree _
    Lib "kernel32" ( _
    ByVal hMem As Long) As Long
  
Private Declare Sub CopyMemory Lib _
    "kernel32" Alias "RtlMoveMemory" ( _
    hpvDest As Any, _
    hpvSource As Any, _
    ByVal cbCopy As Long)

Private Declare Function CopyPointer2String _
    Lib "kernel32" _
    Alias "lstrcpyA" ( _
    ByVal NewString As String, _
    ByVal OldString As Long) As Long
    
Private Declare Function CopyString2Pointer _
    Lib "kernel32" _
    Alias "lstrcpyA" ( _
    ByVal NewString As Long, _
    ByVal OldString As String) As Long
    
Private Declare Function StrToPtr _
    Lib "kernel32" _
    Alias "lstrcpyW" ( _
    ByVal Ptr As Long, _
    Source As Byte) As Long
    

Private Function PointerToString(p As Long) As String
         
'The values returned in the NETRESOURCE structures are pointers to
'ANSI strings so they need to be converted to Visual Basic Strings.
         
Dim s As String
         
    s = String(255, Chr$(0))
    CopyPointer2String s, p
    PointerToString = Left(s, InStr(s, Chr$(0)) - 1)

End Function
Private Function StringToPointer(s As String) As Long
         
'The values returned in the NETRESOURCE structures are pointers to
'ANSI strings so they need to be converted to Visual Basic Strings.
         
Dim p As Long
         
    p = GlobalAlloc(GPTR, 255)
    CopyPointer2String s, p
    StringToPointer = p
    
End Function


Public Function fnEnum( _
    strProviderAttributes As String, _
    strDisplayName() As String, _
    strAttributes() As String, _
    lngIcon() As Long) As Long

Dim typeNet As NETRESOURCE
Dim lngAttributes(8) As Long
Dim lngPos1 As Long
Dim lngPos2 As Long
Dim i As Long
Dim lngReturn As Long
Dim lngEnumHandle As Long
Dim cbBuff As Long
Dim lpBuff As Long
Dim cCount As Long
Dim p As Long ' Pointer

    typeNet.dwUsage = RESOURCEUSAGE_CONTAINER
    typeNet.lpRemoteName = 0
    
    If strProviderAttributes <> "" Then
    
'
' Should be a string with ^ as separater
'

        lngPos1 = 1
        lngPos2 = 999
        i = 1
        
        Do
            
            lngPos2 = InStr(lngPos1, strProviderAttributes, "^")
            
            If lngPos2 = 0 Then
                lngAttributes(i) = Val(Mid$( _
                    strProviderAttributes, _
                    lngPos1, _
                    Len(strProviderAttributes) - lngPos1 + 1))
                Exit Do
            End If
        
            lngAttributes(i) = Val(Mid$( _
                strProviderAttributes, _
                lngPos1, _
                lngPos2 - lngPos1))
            i = i + 1
            lngPos1 = lngPos2 + 1
            
        Loop
        
'
' Build the Resource data structure
'

        typeNet.dwDisplayType = lngAttributes(1)
        typeNet.dwScope = lngAttributes(2)
        typeNet.dwType = lngAttributes(3)
        typeNet.dwUsage = lngAttributes(4)
        typeNet.lpComment = lngAttributes(5)
        typeNet.lpLocalName = lngAttributes(6)
        typeNet.lpProvider = lngAttributes(7)
        typeNet.lpRemoteName = lngAttributes(8)

    End If
    
'
' Open a call to NetEnum
'

    lngReturn = WNetOpenEnum( _
        RESOURCE_GLOBALNET, _
        RESOURCETYPE_ANY, _
        0, _
        typeNet, _
        lngEnumHandle)


    If lngReturn = 0 Then
    
'Create a buffer large enough for the results.
'1000 bytes should be sufficient.

        cbBuff = 1000
        lpBuff = GlobalAlloc(GPTR, cbBuff)
        cCount = &HFFFFFFFF
        
        lngReturn = WNetEnumResource( _
            lngEnumHandle, _
            cCount, _
            lpBuff, _
            cbBuff)

        If lngReturn = 0 Then
            p = lpBuff
            

'WNetEnumResource fills the buffer with an array of
'NETRESOURCE structures. Walk through the list and print
'each local and remote name.
    
'
' Prepare the arrays
'

            ReDim strDisplayName(cCount)
            ReDim strAttributes(cCount)
            ReDim lngIcon(cCount)
            
            For i = 1 To cCount
                CopyMemory typeNet, ByVal p, LenB(typeNet)
                p = p + LenB(typeNet)
                
                '
                ' Remote name gives the network provider name,
                ' Domain name, Server and share etc
                '
                
                strDisplayName(i) = PointerToString(typeNet.lpRemoteName) & "^" & _
                    PointerToString(typeNet.lpComment)
                lngIcon(i) = typeNet.dwDisplayType
                strAttributes(i) = Str$(typeNet.dwDisplayType) & "^" & _
                    Str$(typeNet.dwScope) & "^" & _
                    Str$(typeNet.dwType) & "^" & _
                    Str$(typeNet.dwUsage) & "^" & _
                    Str$(typeNet.lpComment) & "^" & _
                    Str$(typeNet.lpLocalName) & "^" & _
                    Str$(typeNet.lpProvider) & "^" & _
                    Str$(typeNet.lpRemoteName)
            Next i
        Else
            fnEnum = lngReturn
        End If
    
        If lpBuff <> 0 Then GlobalFree (lpBuff)
        WNetCloseEnum (lngEnumHandle) 'Close the enumeration
    Else
        fnEnum = lngReturn
    End If
        
    End Function

Function GetIcon(lngIcon As Long) As String

    Select Case lngIcon
    Case RESOURCEDISPLAYTYPE_GENERIC
        GetIcon = "Generic"
    Case RESOURCEDISPLAYTYPE_DOMAIN
        GetIcon = "Domain"
    Case RESOURCEDISPLAYTYPE_SERVER
        GetIcon = "Server"
    Case RESOURCEDISPLAYTYPE_SHARE
        GetIcon = "Share"
    Case RESOURCEDISPLAYTYPE_FILE
        GetIcon = "File"
    Case RESOURCEDISPLAYTYPE_GROUP
        GetIcon = "Generic"
    Case RESOURCEDISPLAYTYPE_NETWORK
        GetIcon = "Network"
    Case RESOURCEDISPLAYTYPE_ROOT
        GetIcon = "Generic"
    Case RESOURCEDISPLAYTYPE_ADMINSHARE
        GetIcon = "AdminShare"
    Case RESOURCEDISPLAYTYPE_DIRECTORY
        GetIcon = "Directory"
    Case Else
        GetIcon = "Generic"
    End Select
        
End Function
Function fnGetDisplay(strInput As String, _
    strDisp As String, strComment As String) As Long
    
Dim lngPosHat As Long

    lngPosHat = InStr(strInput, "^")
    If lngPosHat > 0 Then
        strDisp = Left$(strInput, lngPosHat - 1)
        strComment = Right$(strInput, Len(strInput) - lngPosHat)
    Else
        strDisp = strInput
        strComment = ""
    End If
    

End Function
