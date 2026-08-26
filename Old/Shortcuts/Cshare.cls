VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CShare"
Attribute VB_Creatable = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object of type RESOURCEDISPLAYTYPE_SHARE. A resource shared by a workstation or server. Typically a disk drive or directory, or a printer."
Option Explicit

Dim Info            As NETRESOURCE
Dim typInfo         As SHARE_INFO_2_STRING

Private Declare Function NetShareGetInfo Lib "netapi32.dll" _
   (ServerName As Byte, _
    NetName As Byte, _
    ByVal Level As Long, _
    bufptr As Any) As Long
 
Private Type SHARE_INFO_1
    NetName         As Long
    ShareType       As Long
    Remark          As Long
End Type
 
Private Type SHARE_INFO_1_STRING
    NetName         As String
    ShareType       As Long
    Remark          As String
End Type
 
Private Type SHARE_INFO_2
    NetName         As Long
    ShareType       As Long
    Remark          As Long
    Permissions     As Long
    Max_Uses        As Long
    Current_Uses    As Long
    Path            As Long
    Passwd          As Long
End Type

Private Type SHARE_INFO_2_STRING
    NetName         As String
    ShareType       As Long
    Remark          As String
    Permissions     As Long
    Max_Uses        As Long
    Current_Uses    As Long
    Path            As String
    Passwd          As String
End Type

Property Get Comment() As String
Attribute Comment.VB_Description = "The comment, or description, for this share."
    Comment = Info.lpComment
End Property

Sub Enumerate()
Attribute Enumerate.VB_Description = "A stub Sub for the CShare object - there is nothing to enumerate."
    ' Sub is a stub
End Sub

Function InitOn(sNetworkName As String) As Object
Attribute InitOn.VB_Description = "A stub - does nothing."
    Set InitOn = Nothing
End Function

Property Get LocalName() As String
Attribute LocalName.VB_Description = "The LocalName property returned by the API enumeration functions."
    LocalName = Info.lpLocalName
End Property

Property Get Path() As String
Attribute Path.VB_Description = "The path local to the server for this ShareName."
    If Len(typInfo.Path) > 0 Then
        Path = typInfo.Path
    End If
End Property
Property Get Permissions() As Long
    Permissions = typInfo.Permissions
End Property

Private Function pGetShareInfo() As Boolean
    Dim lrc             As Long
    Dim lBufPtr         As Long
    Dim sWrk            As String
    Dim lLength         As Long
    Dim lBufLen         As Long
    Dim bytBuf()        As Byte
    Dim bytServer()     As Byte
    Dim bytShare()      As Byte
    Dim typShareInfo    As SHARE_INFO_2
    
    If Len(Info.lpRemoteName) = 0 Then
        Exit Function
    End If
        
    sWrk = Info.lpRemoteName
    lrc = InStr(3, sWrk, "\")
    bytServer = ""
    bytShare = sWrk
    
    lBufLen = 2048
    ReDim bytBuf(lBufLen)
    
    lrc = NetShareGetInfo(0&, bytShare(0), 1&, bytBuf(0))
    If lrc <> ERROR_SUCCESS Then
        pGetShareInfo = False
        Exit Function
    End If
    
    ' It's actually a pointer to a buffer,
    ' so copy the pointer first
    CopyMemory lBufPtr, VarPtr(bytBuf(0)), Len(lBufPtr)
    CopyMemory typShareInfo, lBufPtr, Len(typShareInfo)
    
    If typShareInfo.Remark <> 0 Then
        lLength = lstrlen(typShareInfo.Remark)
        typInfo.Remark = Space$(lLength)
        CopyMemory ByVal typInfo.Remark, ByVal typShareInfo.Remark, lLength
    Else
        typInfo.Remark = ""
    End If
    
    If typShareInfo.Path <> 0 Then
        lLength = lstrlen(typShareInfo.Path)
        typInfo.Path = Space$(lLength)
        CopyMemory ByVal typInfo.Path, ByVal typShareInfo.Path, lLength
    Else
        typInfo.Path = ""
    End If
    
    typInfo.ShareType = typShareInfo.ShareType
    typInfo.Permissions = typShareInfo.Permissions
    
    NetApiBufferFree lBufPtr
End Function
Property Get Provider() As String
Attribute Provider.VB_Description = "The network provider at the root of the hierarchy in which this share name is a member."
    Provider = Info.lpProvider
End Property

Property Get Remark() As String
Attribute Remark.VB_Description = "The Remark attached to this share name, as returned by the NetShareGetInfo API call."
    Remark = typInfo.Remark
End Property

Property Get RemoteName() As String
Attribute RemoteName.VB_Description = "The remote name of this share."
    RemoteName = Info.lpRemoteName
End Property

Sub Init(Scope As Long, dwType As Long, DisplayType As Long, Usage As Long, _
    LocalName As String, RemoteName As String, Comment As String, Provider As String)
Attribute Init.VB_Description = "Initializes the CShare object with parameters from a NETRESOURCE type."

    With Info
        .dwScope = Scope
        .dwType = dwType
        .dwDisplayType = DisplayType
        .dwUsage = Usage
        .lpLocalName = LocalName
        .lpRemoteName = RemoteName
        .lpComment = Comment
        .lpProvider = Provider
    End With
    
    Call pGetShareInfo

End Sub

Property Get ShareType() As Long
Attribute ShareType.VB_Description = "The share type. Either RESOURCETYPE_DISK (1) or RESOURCETYPE_PRINT (2)."
    ShareType = Info.dwType
End Property
Property Get ShareTypeLM() As Long
    ShareTypeLM = typInfo.ShareType
End Property


