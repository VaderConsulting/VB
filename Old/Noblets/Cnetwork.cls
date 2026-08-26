VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNetwork"
Attribute VB_Creatable = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object of type RESOURCEDISPLAYTYPE_NETWORK. A network."
Option Explicit

Dim Info            As NETRESOURCE

Dim colDomains  As Collection

Sub Enumerate()
Attribute Enumerate.VB_Description = "Builds a collection of CDomains known to this network, returnable as the Domains property."
    Set colDomains = EnumerateNR(Info)
End Sub

Function InitOn(sNetworkName As String) As CDomain
Attribute InitOn.VB_Description = "This function returns the CDomain object identified by sNetworkName, if it exists.  Use the Init function prior to calling this function, as the Network object must itself be valid first."
    
    Set InitOn = pInitOn(Info, sNetworkName)
            
End Function

Property Get Comment() As String
Attribute Comment.VB_Description = "The Comment for this network (presently always null)."
    Comment = Info.lpComment
End Property

Property Get Domain(vDomain As Variant) As CDomain
Attribute Domain.VB_Description = "Retuns the specific CDomain object identified by the parameter vDomain, if it is in the Networks collection of Domains.  Use the Init and Enumerate functions prior to using this function."
    Set Domain = colDomains(vDomain)
End Property

Sub Init(Scope As Long, dwType As Long, DisplayType As Long, Usage As Long, _
    LocalName As String, RemoteName As String, Comment As String, Provider As String)
Attribute Init.VB_Description = "Initializes the CNetwork object with parameters from a NETRESOURCE type."

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
Property Get LocalName() As String
Attribute LocalName.VB_Description = "The LocalName returned by the WNetEnum functions (presently always null)."
    LocalName = Info.lpLocalName
End Property

Property Get Domains() As Collection
Attribute Domains.VB_Description = "Returns the collection of CDomains known to  the network."
    If colDomains.Count = 0 Then
        Call Me.Enumerate
    End If
    Set Domains = colDomains
End Property

Property Get Provider() As String
Attribute Provider.VB_Description = "The network provider through which this network is known to the local computer. (I.e, Micorsoft Windows Network)."
    Provider = Info.lpProvider
End Property

Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the Domains collection, setting it to Nothing and then rebuilding it."
    Set colDomains = Nothing
    Set colDomains = EnumerateNR(Info)
End Sub

Property Get RemoteName() As String
Attribute RemoteName.VB_Description = "The remote name of this network."
    RemoteName = Info.lpRemoteName
End Property

Private Sub Class_Initialize()
    Set colDomains = New Collection
End Sub

Private Sub Class_Terminate()
    Set colDomains = Nothing
End Sub


