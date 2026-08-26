VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGroup"
Attribute VB_Creatable = False
Attribute VB_Exposed = True
Attribute VB_Description = "A network group.  A global group known to the domain at large has property IsGlobal set to true.  If not global, a group can be local to a server or workstation."
Option Explicit

Dim psNetName   As String
Dim psFullName  As String
Dim psComment   As String
Dim pbGlobal    As Boolean

Private Type GROUP_INFO_1
    Name    As Long '   LPWSTR   grpi1_name;
    Comment As Long '   LPWSTR grpi1_comment
End Type

Private Declare Function NetGroupGetInfo Lib "netapi32" _
    (ServerName As Any, _
     GroupName As Any, _
     ByVal Level As Long, _
     buffer As Long) As Long
' LPWSTR servername,
' LPWSTR groupname,
'    DWORD level,
'    LPBYTE *bufptr

Property Get Comment() As String
Attribute Comment.VB_Description = "The comment that is attached to the group, whether local or global."
    Comment = psComment
End Property

Property Get FullName() As String
Attribute FullName.VB_Description = "The full, descriptive name of the group."
    FullName = psFullName
End Property
Property Get GroupName() As String
Attribute GroupName.VB_Description = "The name of the group within its domain.  This domain can be either a network domain (it is then a global group) or a server (in which case it is a local group)."
    GroupName = psNetName
End Property
Function Init(sNetGroup As String, sServer As String, bGlobal As Boolean) As Boolean
Attribute Init.VB_Description = "Initializes a CGroup object, using sNetGroup and sServer as the parameters in the API call to NetGroupGetInfo.  bGlobal should be true when the sServer is a domain controller."

Dim lrc     As Long
Dim sWrk    As String

Dim pServer()   As Byte
Dim pGroup()    As Byte
       
    psNetName = sNetGroup
       
    pGroup = sNetGroup & vbNullChar
    pServer = sServer & vbNullChar
    psComment = vbNullChar
    pbGlobal = bGlobal
    
    Dim dwLevel As Long
    dwLevel = 1
    
    Dim tmpBuffer    As GROUP_INFO_1
    Dim ptmpBuffer   As Long
    
    lrc = NetGroupGetInfo(pServer(0), pGroup(0), dwLevel, ptmpBuffer)
    'As last param is dimmed as long, the pointer to ptmpBuffer is
    'passed to dll, and the function returns a pointer to a pointer
    'to our UDT. Therefore ptmpBuffer on return holds a pointer
    'to our UDT.
    
     If lrc <> 0 Then
        Init = False
        Exit Function
     End If
    
    'Deference it!!!
    CopyMemory tmpBuffer, ptmpBuffer, LenB(tmpBuffer)
    
    Dim sByte() As Byte
    ReDim sByte(255)
    
    'Convert LPWSTR (Unicode string) to VB string.
    CopyMemory sByte(0), tmpBuffer.Name, 256
    psFullName = sByte
    psFullName = Left$(psFullName, InStr(psFullName, vbNullChar) - 1)
    
    'Convert LPWSTR (Unicode string) to VB string.
    CopyMemory sByte(0), tmpBuffer.Comment, 256
    psComment = sByte
    psComment = psComment & vbNullChar
    psComment = Left$(psComment, InStr(psComment, vbNullChar) - 1)
    
    Init = True

End Function



Property Get IsGlobal() As Boolean
Attribute IsGlobal.VB_Description = "Returns true if the group is a global group within its domain."
    IsGlobal = pbGlobal
End Property


