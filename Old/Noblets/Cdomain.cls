VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDomain"
Attribute VB_Creatable = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object of type RESOURCEDISPLAYTYPE_DOMAIN. A domain."
Option Explicit

Dim Info        As NETRESOURCE

Dim colServers  As Collection   ' Servers collection
Dim colUsers    As Collection   ' Users collection
Dim colGroups   As Collection   ' User groups collection

Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the Servers collection, setting it to Nothing and then rebuilding it."
    Set colServers = Nothing
    Set colServers = EnumerateNR(Info)
End Sub

Sub Enumerate()
Attribute Enumerate.VB_Description = "Drills down into the domain, building a collection of Servers in the domain. This can be a time consuming and very resource costly operation on a large network!"
    Set colServers = EnumerateNR(Info)
End Sub

Property Get Groups() As Collection
Attribute Groups.VB_Description = "Returns a collection of CGroup objects representing the global groups known in the Domain."
    If colGroups.Count = 0 Then _
        pGetGroups
    Set Groups = colGroups
End Property
Sub Init(Scope As Long, dwType As Long, DisplayType As Long, Usage As Long, _
    LocalName As String, RemoteName As String, Comment As String, Provider As String)
Attribute Init.VB_Description = "Initializes the CDomain object with parameters from a NETRESOURCE type."

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
End Sub
Property Get Comment() As String
Attribute Comment.VB_Description = "The Comment attached to this domain and returned with the WNetEnum API functions.  (Presently, this field is always null.)"
    Comment = Info.lpComment
End Property

Property Get LocalName() As String
Attribute LocalName.VB_Description = "The LocalName for this domain, as returned using the API enumeration functions.  Presently, this field is always null."
    LocalName = Info.lpLocalName
End Property

Private Sub pGetGroups()
    Dim oServer As CServer
    
    If colServers.Count = 0 Then _
        Set colServers = EnumerateNR(Info)
        
    For Each oServer In colServers
        If oServer.IsPDC Or oServer.IsBDC Then Exit For
    Next
    
    Set colGroups = EnumerateGroups(oServer.RemoteName, "")
End Sub

Private Sub pGetUsers()
    Dim oServer As CServer
    
    If colServers.Count = 0 Then _
        Set colServers = EnumerateNR(Info)
        
    For Each oServer In colServers
        If oServer.IsPDC Or oServer.IsBDC Then Exit For
    Next
    
    Set colUsers = EnumerateUsers(oServer.RemoteName, "")
End Sub

Property Get Provider() As String
Attribute Provider.VB_Description = "The network provider by which this domain has been enumerated (i.e., Microsoft Network)."
    Provider = Info.lpProvider
End Property

Property Get RemoteName() As String
Attribute RemoteName.VB_Description = "The remote name of this domain."
    RemoteName = Info.lpRemoteName
End Property

Property Get Server(vServer As Variant) As CServer
Attribute Server.VB_Description = "Returns the server, using vServer as the item.  The vServer, a variant, can be the index or the key in the collection.  Build the collection with the Enumerate function before calling this function."
    If colServers.Count = 0 Then _
        Set colServers = EnumerateNR(Info)
    Set Server = colServers(vServer)
End Property

Property Get Servers() As Collection
Attribute Servers.VB_Description = "Returns a collection of the CServers active in the domain (i.e., known to the enumeration browser at this time)."
    If colServers.Count = 0 Then _
        Set colServers = EnumerateNR(Info)
    Set Servers = colServers
End Property

Property Get Users() As Collection
Attribute Users.VB_Description = "Returns a collection of CUser objects, users in the domain."
    If colUsers.Count = 0 Then pGetUsers
    Set Users = colUsers
End Property

Private Sub Class_Initialize()
    Set colServers = New Collection
    Set colUsers = New Collection
    Set colGroups = New Collection
End Sub

Private Sub Class_Terminate()
    Set colServers = Nothing
    Set colUsers = Nothing
    Set colGroups = Nothing
End Sub
