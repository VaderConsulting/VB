VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Peer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim WithEvents oServer As CSocket
Attribute oServer.VB_VarHelpID = -1

Public Function StartServer() As Boolean
    Set oServer = New CSocket
    'oServer.Protocol = sckTCPProtocol
    oServer.LocalPort = 1001
    oServer.Listen
    StartServer = True
End Function

Public Function CloseServer() As Boolean
    oServer.CloseSocket
    Set oServer = Nothing
End Function

Private Sub Class_Initialize()
    Debug.Print Now & " TCPPeer class initialised."
End Sub

Private Sub Class_Terminate()
    Debug.Print "TCPPeer class terminating."
End Sub

Private Sub oServer_OnConnect()
    Debug.Print "Connection request."
End Sub

Private Sub oServer_OnConnectionRequest(ByVal requestID As Long)
    Debug.Print "Connect request received - " & requestID
End Sub
