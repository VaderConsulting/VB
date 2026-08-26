Attribute VB_Name = "MNTNetwork"
Option Explicit

Dim oApp    As New CApplication

Public Const ERROR_ENUMERATION = vbObjectError + 1000

Public oErr As New ErrMsgServer.CErrMsgNT

Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_PUBLICNET = &H2
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
Public Const RESOURCEDISPLAYTYPE_SHAREADMIN = &H8
Public Const RESOURCEDISPLAYTYPE_DIRECTORY = &H9

Public Const UNIVERSAL_NAME_INFO_LEVEL = 1
Public Const REMOTE_NAME_INFO_LEVEL = 2

Public Const WN_SUCCESS = 0

Const FILTER_TEMP_DUPLICATE_ACCOUNT = &H1
Const FILTER_NORMAL_ACCOUNT = &H2
Const FILTER_INTERDOMAIN_TRUST_ACCOUNT = &H8
Const FILTER_WORKSTATION_TRUST_ACCOUNT = &H10
Const FILTER_SERVER_TRUST_ACCOUNT = &H20

Public Const NETINFO_DLL16 = 1&
Public Const NETINFO_DISKRED = 4&   ' Provider requires disk redirections to connect
Public Const NETINFO_PRINTERRED = 8& ' Provider requires printer redirections to connect

Declare Function NetApiBufferFree Lib "netapi32" (bufptr As Any) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As Any, ByVal cbCopy As Long)

Declare Function VarPtr Lib "VB40032" (lpVoid As Any) As Long

Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" _
   (RetVal As Byte, ByVal ptr As Long) As Long
 
Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" _
   (ByVal ptr As Long, source As Byte) As Long
 
Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" _
   (RetVal As Any, ByVal ptr As Long, ByVal nCharCount As Long) As Long
 
Declare Function StrLen Lib "kernel32" Alias "lstrlenW" _
   (ByVal ptr As Long) As Long

Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" _
   (ByVal dwScope As Long, _
    ByVal dwType As Long, _
    ByVal dwUsage As Long, _
    lpNetResource As NETRESOURCE, _
    lphEnum As Long) As Long

' We need a separate declaration for the null case
Public Declare Function WNetOpenEnumRoot Lib "mpr.dll" Alias "WNetOpenEnumA" _
   (ByVal dwScope As Long, _
    ByVal dwType As Long, _
    ByVal dwUsage As Long, _
    ByVal lpNetResource As Long, _
    lphEnum As Long) As Long

Public Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" _
   (ByVal hEnum As Long, _
    lpcCount As Long, _
    lpBuffer As Byte, _
    lpBufferSize As Long) As Long

Public Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long

' Enumerate using Level 0 user structure
Declare Function NetUserEnum0 Lib "netapi32.dll" Alias "NetUserEnum" _
   (ServerName As Byte, _
    ByVal Level As Long, _
    ByVal lFilter As Long, _
    buffer As Long, _
    ByVal PrefMaxLen As Long, _
    EntriesRead As Long, _
    TotalEntries As Long, _
    ResumeHandle As Long) As Long

Declare Function NetGroupEnumUsers0 Lib "netapi32.dll" Alias "NetGroupGetUsers" _
   (ServerName As Byte, _
    GroupName As Byte, _
    ByVal Level As Long, _
    buffer As Long, _
    ByVal PrefMaxLen As Long, _
    EntriesRead As Long, _
    TotalEntries As Long, _
    ResumeHandle As Long) As Long
 
Declare Function NetGroupEnum0 Lib "netapi32.dll" Alias "NetGroupEnum" _
   (ServerName As Byte, _
    ByVal Level As Long, _
    buffer As Long, _
    ByVal PrefMaxLen As Long, _
    EntriesRead As Long, _
    TotalEntries As Long, _
    ResumeHandle As Long) As Long
 
Declare Function NetUserGetGroups0 Lib "netapi32.dll" Alias "NetUserGetGroups" _
   (ServerName As Byte, _
    UserName As Byte, _
    ByVal Level As Long, _
    buffer As Long, _
    ByVal PrefMaxLen As Long, _
    EntriesRead As Long, _
    TotalEntries As Long) As Long
 
' Important error values
Public Const ERROR_EXTENDED_ERROR = 1208&
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const ERROR_MORE_DATA = 234 '  dderror

Type MungeLong
    x As Long
    Dummy As Integer
End Type
 
Type MungeInt
    XLo As Integer
    XHi As Integer
    Dummy As Integer
End Type

Public Type NETRESOURCELONG
    dwScope         As Long
    dwType          As Long
    dwDisplayType   As Long
    dwUsage         As Long
    lpLocalName     As Long
    lpRemoteName    As Long
    lpComment       As Long
    lpProvider      As Long
End Type

Public Type NETRESOURCE
    dwScope         As Long
    dwType          As Long
    dwDisplayType   As Long
    dwUsage         As Long
    lpLocalName     As String
    lpRemoteName    As String
    lpComment       As String
    lpProvider      As String
End Type

' Holds the current copy of the netresource
Private Info As NETRESOURCE

Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" _
   (ByVal lpString As Any) As Long

Function StrFromPtr(lPtr As Long, lMaxLen As Long) As String
    Dim sWrk    As String
    Dim lRC     As Long
    sWrk = String$(lMaxLen, vbNullChar)
    lRC = lstrcpy(sWrk, lPtr)
    If lRC > 0 Then
        StrFromPtr = Left$(sWrk, InStr(sWrk, vbNullChar) - 1)
    Else
        StrFromPtr = ""
    End If
End Function



Function EnumerateGroups(ByVal sName As String, ByVal UName As String) As Collection
' FROM MSDN VBKB - RJH
'
' Enumerates GLOBAL groups only - not local groups
'
' The buffer is filled from the left with pointers to user names that
' are filled from the right side. For example:
'
'     ptr1|ptr2|...|ptrn|<garbage>|strn|...|str2|str1
'     ^-------------- BufPtr buffer ----------------^
'
' On NT, TotalEntries is the number of entries left to be read including
' the currently read entries.
'
' On LanMan and OS/2, it is the total number of entries, period. Code
' would have to be changed to reflect this if the Domain controller
' wasn't an NT machine.
'
' BufPtr gets the address of the buffer (or ptr1 - add 4 to BufPtr for
' each additional pointer)
'
Dim Result          As Long
Dim bufptr          As Long
Dim EntriesRead     As Long
Dim TotalEntries    As Long
Dim ResumeHandle    As Long
Dim BufLen          As Long
Dim SNArray()       As Byte
Dim GNArray(99)     As Byte
Dim UNArray()       As Byte
Dim GName           As String
Dim i               As Integer
Dim UNPtr           As Long
Dim TempPtr         As MungeLong
Dim TempStr         As MungeInt

Dim colGroups       As Collection
Dim oGroup          As CGroup
 
     SNArray = sName & vbNullChar       ' Move to byte array
     UNArray = UName & vbNullChar       ' Move to Byte array
     BufLen = 255                       ' Buffer size
     ResumeHandle = 0                   ' Start with the first entry
    
     Do
        If UName = "" Then
            Result = NetGroupEnum0(SNArray(0), 0, bufptr, BufLen, _
                        EntriesRead, TotalEntries, ResumeHandle)
        Else
            Result = NetUserGetGroups0(SNArray(0), UNArray(0), 0, bufptr, _
                        BufLen, EntriesRead, TotalEntries)
        End If
        If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads
                                             ' required
            Debug.Print "Error " & Result & " enumerating group " & _
                        EntriesRead & " of " & TotalEntries
            Exit Function
        End If
        
        Set colGroups = New Collection
        
        For i = 1 To EntriesRead
            ' Get pointer to string from beginning of buffer
            ' Copy 4 byte block of memory in 2 steps
            Result = PtrToInt(TempStr.XLo, bufptr + (i - 1) * 4, 2)
            Result = PtrToInt(TempStr.XHi, bufptr + (i - 1) * 4 + 2, 2)
            LSet TempPtr = TempStr ' munge 2 Integers to a Long
            
            ' Copy string to array and convert to a string
            Result = PtrToStr(GNArray(0), TempPtr.x)
            GName = Left(GNArray, StrLen(TempPtr.x))
            
            Set oGroup = New NOblets.CGroup
            If oGroup.Init(GName, sName, bGlobal:=True) Then
                colGroups.Add oGroup, Key:=oGroup.GroupName
            End If
'            Debug.Print "Group: " & GName
        Next i
    Loop Until EntriesRead = TotalEntries
    ' The above condition only valid for reading accounts on NT
    ' but not OK for OS/2 or LanMan
    
    Result = NetApiBufferFree(bufptr)         ' Don't leak memory
    Set EnumerateGroups = colGroups
 
End Function

Function IntFromByte(bytBuf() As Byte) As Integer
    Dim iRtn    As Integer
    CopyMemory iRtn, VarPtr(bytBuf(0)), Len(iRtn)
End Function

Function LongFromByte(bytBuf() As Byte) As Long
    Dim lRtn    As Long
    CopyMemory lRtn, VarPtr(bytBuf(0)), Len(lRtn)
    LongFromByte = lRtn
End Function
Private Function StringFromByteW(bytBuf() As Byte) As String
    Dim sWrk    As String
    Dim i       As Integer
    
    For i = 0 To UBound(bytBuf()) Step 2
        If bytBuf(i) = 0 Then Exit For
        sWrk = sWrk & Chr$(bytBuf(i))
    Next
    StringFromByteW = sWrk
End Function
Function pInitOn(nrParent As NETRESOURCE, sNetworkName As String) As Object
    Dim EnumerationHandle&
    Dim res&
    Dim tbuf()      As Byte
    Dim BufferSize  As Long
    Dim lLength     As Long
    Dim nr          As NETRESOURCELONG  ' Temporary structure for copying
    Dim sLocalName  As String
    Dim sRemoteName As String
    Dim sComment    As String
    Dim sProvider   As String
    Dim colObjects  As New Collection
    Dim oObject     As Object
    
    If nrParent.lpProvider = "" Then
        ' Start by opening the root
        res = WNetOpenEnumRoot(RESOURCE_PUBLICNET, RESOURCETYPE_ANY, 0, 0, EnumerationHandle)
    Else
        res = WNetOpenEnum(RESOURCE_PUBLICNET, RESOURCETYPE_ANY, 0, nrParent, EnumerationHandle)
    End If
    
    If res <> 0 Then
        If Err.LastDllError = ERROR_EXTENDED_ERROR Then
            Dim oErr As New CNetErr
            Err.Raise oErr.dwErrorCode, App.EXEName & ".EnumerateNR", oErr.szDescription
        End If
    End If
   
    ' Appleman Comment :
    ' Create a big buffer to work with. We dimension it here instead of
    ' at the declaration to make sure it's allocated off the heap and not the stack
    ReDim tbuf(16384)
    BufferSize = 16384

    Do
        res = WNetEnumResource(EnumerationHandle, 1, tbuf(0), BufferSize)
        ' Check for errors
        Select Case res
            Case 0   ' Success
                ' Copy the necessary data
                CopyMemory nr, ByVal VarPtr(tbuf(0)), Len(nr)
                
                If nr.lpRemoteName <> 0 Then
                    lLength = lstrlen(nr.lpRemoteName)
                    sRemoteName = Space$(lLength)
                    CopyMemory ByVal sRemoteName, ByVal nr.lpRemoteName, lLength
'                    sRemoteName = agGetStringFromPointer(nr.lpRemoteName)
                Else
                    sRemoteName = vbNullString
                End If
                
                If sNetworkName = sRemoteName Then
                    
                    Select Case nr.dwDisplayType
                        Case RESOURCEDISPLAYTYPE_NETWORK
                            Set oObject = New CNetwork
                        Case RESOURCEDISPLAYTYPE_DOMAIN
                            Set oObject = New CDomain
                        Case RESOURCEDISPLAYTYPE_SERVER
                            Set oObject = New CServer
                        Case RESOURCEDISPLAYTYPE_SHARE
                            Set oObject = New CShare
                    End Select
                            
                    oObject.Init Scope:=nr.dwScope, _
                        dwType:=nr.dwType, _
                        DisplayType:=nr.dwDisplayType, _
                        Usage:=nr.dwUsage, _
                        LocalName:=sLocalName, _
                        RemoteName:=sRemoteName, _
                        Comment:=sComment, _
                        Provider:=sProvider
                            
                    colObjects.Add oObject, Key:=oObject.RemoteName
                End If
                
                Exit Do
                
            Case ERROR_NO_MORE_ITEMS
            Case Else
                If Err.LastDllError = ERROR_MORE_DATA Then
                    ' A buffer too small error should be very rare, but
                    ' the case is handled just to be through. The code
                    ' will drop down and try again
                    ReDim tbuf(BufferSize + 1)
                Else
                    ' This type of error can't be handled, so exit
                    ' Err.Raise ERROR_ENUMERATION, App.EXEName & ".CNetwork", Err.Description
                End If
        End Select
    Loop While res = 0

    WNetCloseEnum EnumerationHandle

    Set pInitOn = oObject

End Function



Function EnumerateNR(nrParent As NETRESOURCE) As Collection
    Dim EnumerationHandle&
    Dim res&
    Dim tbuf()      As Byte
    Dim BufferSize  As Long
    Dim lLength     As Long
    Dim nr          As NETRESOURCELONG  ' Temporary structure for copying
    Dim sLocalName  As String
    Dim sRemoteName As String
    Dim sComment    As String
    Dim sProvider   As String
    Dim colObjects  As New Collection
    Dim oObject     As Object
    
    If nrParent.lpProvider = "" Then
        ' Start by opening the root
        res = WNetOpenEnumRoot(RESOURCE_PUBLICNET, RESOURCETYPE_ANY, 0, 0, EnumerationHandle)
    Else
        res = WNetOpenEnum(RESOURCE_PUBLICNET, RESOURCETYPE_ANY, 0, nrParent, EnumerationHandle)
    End If
    
    If res <> 0 Then
        If Err.LastDllError = ERROR_EXTENDED_ERROR Then
            Dim oErr As New CNetErr
            Err.Raise oErr.dwErrorCode, App.EXEName & ".EnumerateNR", oErr.szDescription
        End If
    End If
   
    ' Create a big buffer to work with. We dimention it here instead of
    ' at the declaration to make sure it's allocated off the heap and not the stack
    ReDim tbuf(16384)
    BufferSize = 16384

    Do
        res = WNetEnumResource(EnumerationHandle, 1, tbuf(0), BufferSize)
        ' Check for errors
        Select Case res
            Case 0   ' Success
                ' Copy the necessary data
                CopyMemory nr, ByVal VarPtr(tbuf(0)), Len(nr)
                
                If nr.lpLocalName <> 0 Then
                    lLength = lstrlen(nr.lpLocalName)
                    sLocalName = Space$(lLength)
                    CopyMemory ByVal sLocalName, ByVal nr.lpLocalName, lLength
'                    sLocalName = agGetStringFromPointer(nr.lpLocalName)
                Else
                    sLocalName = vbNullString
                End If
                
                If nr.lpRemoteName <> 0 Then
                    lLength = lstrlen(nr.lpRemoteName)
                    sRemoteName = Space$(lLength)
                    CopyMemory ByVal sRemoteName, ByVal nr.lpRemoteName, lLength
                Else
                    sRemoteName = vbNullString
                End If
                
                If nr.lpComment <> 0 Then
                    lLength = lstrlen(nr.lpComment)
                    sComment = Space$(lLength)
                    CopyMemory ByVal sComment, ByVal nr.lpComment, lLength
                Else
                    sComment = vbNullString
                End If
                
                If nr.lpProvider <> 0 Then
                    lLength = lstrlen(nr.lpProvider)
                    sProvider = Space$(lLength)
                    CopyMemory ByVal sProvider, ByVal nr.lpProvider, lLength
                Else
                    sProvider = vbNullString
                End If
                
                Select Case nr.dwDisplayType
                    Case RESOURCEDISPLAYTYPE_NETWORK
                        Set oObject = New CNetwork
                    Case RESOURCEDISPLAYTYPE_DOMAIN
                        Set oObject = New CDomain
                    Case RESOURCEDISPLAYTYPE_SERVER
                        Set oObject = New CServer
                    Case RESOURCEDISPLAYTYPE_SHARE
                        Set oObject = New CShare
                End Select
                        
                oObject.Init Scope:=nr.dwScope, _
                    dwType:=nr.dwType, _
                    DisplayType:=nr.dwDisplayType, _
                    Usage:=nr.dwUsage, _
                    LocalName:=sLocalName, _
                    RemoteName:=sRemoteName, _
                    Comment:=sComment, _
                    Provider:=sProvider
                
                colObjects.Add oObject, Key:=oObject.RemoteName
                
            Case ERROR_NO_MORE_ITEMS
            Case Else
                If Err.LastDllError = ERROR_MORE_DATA Then
                    ' A buffer too small error should be very rare, but
                    ' the case is handled just to be through. The code
                    ' will drop down and try again
                    ReDim tbuf(BufferSize + 1)
                Else
                    ' This type of error can't be handled, so exit
                    ' Err.Raise ERROR_ENUMERATION, App.EXEName & ".CNetwork", Err.Description
                End If
        End Select
    Loop While res = 0

    WNetCloseEnum EnumerationHandle

    Set EnumerateNR = colObjects
    
End Function

Function EnumerateUsers(ByVal sName As String, ByVal GName As String) As Collection

Dim Result          As Long
Dim bufptr          As Long
Dim EntriesRead     As Long
Dim TotalEntries    As Long
Dim ResumeHandle    As Long
Dim BufLen          As Long
Dim SNArray()       As Byte
Dim GNArray()       As Byte
Dim UNArray(99)     As Byte
Dim UName           As String
Dim i               As Integer
Dim UNPtr           As Long
Dim TempPtr         As MungeLong
Dim TempStr         As MungeInt
 
Dim oUser           As CUser
Dim colUsers        As Collection
 
    SNArray = sName & vbNullChar       ' Move to byte array
    GNArray = GName & vbNullChar       ' Move to Byte array
    BufLen = 255                       ' Buffer size
    ResumeHandle = 0                   ' Start with the first entry
 
    Do
        If GName = "" Then
            Result = NetUserEnum0(SNArray(0), 0, FILTER_NORMAL_ACCOUNT, _
                        bufptr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
        Else
            Result = NetGroupEnumUsers0(SNArray(0), GNArray(0), 0, bufptr, _
                        BufLen, EntriesRead, TotalEntries, ResumeHandle)
        End If
        
        If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads
                                                 ' required
            Err.Raise ERROR_ENUMERATION, App.EXEName & ".EnumerateUsers", _
                "Error " & Result & " enumerating user " & EntriesRead & " of " & TotalEntries
'            If Result = 2220 Then _
                Debug.Print "There is no **GLOBAL** group '" & GName & "'"
        End If
        
        Set colUsers = New Collection
        
        For i = 1 To EntriesRead
            
            ' Get pointer to string from beginning of buffer
            ' Copy 4-byte block of memory in 2 steps
            Result = PtrToInt(TempStr.XLo, bufptr + (i - 1) * 4, 2)
            Result = PtrToInt(TempStr.XHi, bufptr + (i - 1) * 4 + 2, 2)
            LSet TempPtr = TempStr ' munge 2 integers into a Long
            
            ' Copy string to array
            Result = PtrToStr(UNArray(0), TempPtr.x)
            UName = Left(UNArray, StrLen(TempPtr.x))
            
            Set oUser = New CUser
            If oUser.Init(UName, sName) Then
                colUsers.Add oUser, Key:=oUser.UserName
            End If
'            Debug.Print "User: " & UName
        Next i
    Loop Until EntriesRead = TotalEntries
 
    Result = NetApiBufferFree(bufptr)         ' Don't leak memory
 
    Set EnumerateUsers = colUsers
 
End Function

Sub Main()

End Sub

