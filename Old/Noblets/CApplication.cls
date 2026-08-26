VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CApplication"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Attribute VB_Description = "The NT Network Enumeration OLE Library - an encapsulation of WNetEnumResource in OLE."
Option Explicit

Private colNetworks     As Collection

Function WorkstationLinks() As Collection
Attribute WorkstationLinks.VB_Description = "Returns a collection enumerating the shared resources connected by the workstation."
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
    Dim oShare      As CShare
    Dim oErr        As CNetErr
    
    res = WNetOpenEnumRoot(RESOURCE_CONNECTED, RESOURCETYPE_ANY, 0, 0, EnumerationHandle)
    
    If res <> 0 Then
        If Err.LastDllError = ERROR_EXTENDED_ERROR Then
            Set oErr = New CNetErr
            Err.Raise oErr.dwErrorCode, App.EXEName & ".WkstaLinks", oErr.szDescription
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
                
        
                Set oShare = New CShare
                        
                oShare.Init Scope:=nr.dwScope, _
                    dwType:=nr.dwType, _
                    DisplayType:=nr.dwDisplayType, _
                    Usage:=nr.dwUsage, _
                    LocalName:=sLocalName, _
                    RemoteName:=sRemoteName, _
                    Comment:=sComment, _
                    Provider:=sProvider
        
                On Local Error Resume Next  ' Can be multiple links to a share
                colObjects.Add oShare, Key:=oShare.RemoteName
                On Local Error GoTo 0
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

    Set WorkstationLinks = colObjects
    
End Function



Function Init() As Boolean
Attribute Init.VB_Description = "Initilizes the CApplication object by enumerating the networks at the root."
    Dim nr  As NETRESOURCE
    
    nr.lpProvider = ""
    Set colNetworks = EnumerateNR(nr)
End Function
Function InitOn(sNetworkName As String) As CNetwork
Attribute InitOn.VB_Description = "Initializes on a particular network (i.e., Microsoft Windows Network) and builds a collection of CDomain objects representing the domains known to that network."
    Dim nr  As NETRESOURCE
    
    nr.lpProvider = ""
    Set InitOn = pInitOn(nr, sNetworkName)
            
End Function


Property Get Network(vNetwork As Variant) As CNetwork
Attribute Network.VB_Description = "Finds and returns the CNetwork object indicated by vNetwork.  vNetwork can be either the index (spinning through) or the key (give me THIS one) of the CNetwork object.  Returns Nothing if not found."
    Set Network = colNetworks(vNetwork)
End Property

Property Get Networks() As Collection
Attribute Networks.VB_Description = "A pointer to the CNetwork objects collection owned by this Application instance. "
    Set Networks = colNetworks
End Property

Private Sub Class_Initialize()
    Set colNetworks = New Collection
End Sub

