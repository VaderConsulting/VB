VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemoteProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private xTaskID As Integer
Private xProcessID As Long
Private xAgentID As Integer
Private xCMD As String
Private xTCPID As Integer
Private ProcState As Integer
Private LastUpdate As Date

Public Property Get TCPID() As Integer
  TCPID = xTCPID
End Property

Public Property Get AgentID() As Integer
  AgentID = xAgentID
End Property

Public Property Get TaskID() As Integer
  TaskID = xTaskID
End Property

Public Sub Create(ByVal iAgentID As Integer, _
                  ByVal iTaskID As Integer, _
                  ByVal sCmdExec As String)
  Dim i As Integer
  Dim oTCP As clsTCP
  
  xAgentID = iAgentID
  xTaskID = iTaskID
  xCMD = sCmdExec
  For i = 1 To TCPConns.Count
    Set oTCP = TCPConns.Item(i)
    If (oTCP.AgentID = iAgentID) Then
      xTCPID = i
      oTCP.SendString "S106:" & TaskID & "," & sCmdExec
      ProcState = psPending
    End If
  Next
  Set oTCP = Nothing
  'Sanity Check
  If (ProcState <> psPending) Then ProcState = psError
  LastUpdate = Now
End Sub

Private Sub Class_Initialize()
  xAgentID = 0
  xTaskID = 0
  xCMD = ""
  xProcessID = 0
  xTCPID = 0
  ProcState = psIdle
  LastUpdate = Now
End Sub
