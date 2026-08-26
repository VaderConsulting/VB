VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xID As Integer
Private DataStr As String
Private ConnState As Integer
Private LastSentString As String
' csNotAuthorised
' csAuthorised
' csConfigured
' csReady

'Agent Info
Private xAgentID As Integer
Private xAgentHostname As String
Private MaxProcesses As Integer
Private ExclusiveMode As Boolean
Private xAgentUsername As String


'Encryption Information
Private EncKey As String

Property Get isReceivingTasks() As Boolean
  isReceivingTasks = (ConnState = csReady)
End Property

Property Get Options(ByVal OptName As String) As Variant
  OptName = UCase(OptName)
  Select Case OptName
    Case "EXCLUSIVEMODE"
      Options = ExclusiveMode
    Case "MAXPROCESSES"
      Options = MaxProcesses
    Case "NUMPROCESSES"
      'Options = GetNumProcesses(xAgentID)
  End Select
End Property

Property Get ConnectionState() As Integer
  ConnectionState = ConnState
End Property

Property Get AgentID() As Integer
  AgentID = xAgentID
End Property

Property Get AgentHostname() As String
  AgentHostname = xAgentHostname
End Property

Property Get AgentUsername() As String
  AgentUsername = xAgentUsername
End Property

Property Get isConfigured() As Boolean
  isConfigured = (ConnState = csConfigured) Or (ConnState = csReady)
End Property

Private Sub ProcessPacket(ByVal sData As String)
  Dim Cmd As String
  Dim Parm As String
  ' *******************************
  sData = Decrypt(sData, EncKey)
  ' *******************************
  If (Len(sData) >= 5) Then
    Cmd = Left(sData, 5)
    Parm = Right(sData, Len(sData) - 5)
    
    AddToHistory Cmd + Parm
    
    ' Doesn't matter if client is Authorised or not
    Select Case Cmd
      Case "R1FD:" ' General broadcast from client
        SendString "S1FD:OK" + vbCrLf
        Exit Sub
    End Select
    
    If isAuthorised Then ' Incoming message is from an authorised client
      Select Case Cmd
        Case "R101:" ' Requesting Configuration
          ' Process Agent Information
          xAgentHostname = UCase(Parm)
          SendString "S101:OK" & vbCrLf
          ConnState = csConfigured
          frmMain.RefreshAgentsList
          Exit Sub
        Case "R102:" ' Agent is Ready to Receive Alerts
          If (ConnState <> csConfigured) Then Exit Sub
          SendString "S102:OK" + vbCrLf
          ConnState = csReady
          SendString "S104:OK" + vbCrLf
          Exit Sub
        Case "R105:" 'Agent is stopped
          If (ConnState <> csReady) Then Exit Sub
          If (Parm <> "OK") Then Exit Sub
          ConnState = csConfigured
          Exit Sub
        Case "R103:" ' Agent is sending username
          SendString "S103:OK" + vbCrLf
          xAgentUsername = Parm
          frmMain.RefreshAgentsList
          Exit Sub
      End Select
    Else             ' The following messages are from a client not authorised
      Select Case Cmd
        Case "A001:" ' Requesting Authorisation
          If isAuthorised Then
            SendString "A001:ERR, Already Authorised" + vbCrLf
            Exit Sub
          End If
          If Authorise(Parm) Then
            SendString "A001:OK" + vbCrLf
            ConnState = csAuthorised
            Exit Sub
          Else
            SendString "A001:NAK" + vbCrLf
            Exit Sub
          End If
      End Select
    End If
  End If
End Sub

Public Sub DataArrival(ByVal BytesTotal As Integer)
  Dim s As String
  Dim i As Integer
  Dim j As Integer
  Dim t As String
  
  frmMain.tcpMain(xID).GetData s, vbString
  DataStr = DataStr + s
  
  'Check for CRLF Chars
  Do
    i = MinNZ(0, InStr(DataStr, Chr(13)))
    i = MinNZ(i, InStr(DataStr, Chr(10)))
    If (i <> 0) Then
      t = Left(DataStr, i - 1)
      If DataStr <> "" Then
        DataStr = Right(DataStr, Len(DataStr) - i)
      End If
      ProcessPacket t
    End If
    DoEvents
  Loop Until (i = 0)
  DoEvents
End Sub

Private Function Authorise(ByVal Pwd As String)
  Authorise = False
  If isAuthorised Then
    Authorise = True
    Exit Function
  End If
  If (Pwd = "ALERTER") Then ConnState = csAuthorised 'Change this to a shifting password
  Authorise = (ConnState = csAuthorised)
End Function

Property Get ID() As Integer
  ID = xID
End Property

Public Sub SendBroadcast(ByVal Data As String)
  If ConnState = csReady Then SendString Data
End Sub

Public Sub SendString(ByVal Data As String)
  On Error GoTo SendStringError
  
  'Debug.Print "Sending " & Data & " to Client " & xID
  LastSentString = Data
  
  If Right(Data, 2) = vbCrLf Then Data = Left(Data, Len(Data) - 2)
  Data = Encrypt(Data, EncKey) + vbCrLf
  
  'Debug.Print "Data Encrypted as " & Data
  frmMain.tcpMain(xID).SendData Data
  Exit Sub
SendStringError:
  Err.Clear
  Exit Sub
End Sub

Public Sub SetupTCPInfo(ByVal tcpMainID As Integer)
  xID = tcpMainID
  If DoTCPEncrypt Then EncKey = frmMain.tcpMain(tcpMainID).RemoteHostIP + CStr(frmMain.tcpMain(tcpMainID).RemotePort)
End Sub

Property Get isAuthorised() As Boolean
  isAuthorised = (ConnState = csAuthorised) Or (ConnState = csConfigured) Or (ConnState = csReady)
End Property

Private Sub Class_Initialize()
  xID = 0
  DataStr = ""
  ConnState = csNotAuthorised
  xAgentID = 0
  EncKey = ""
End Sub

' DEBUG ONLY ******* DO NOT USE ********
'Public Property Let FakeSpawn(ByVal Value As Integer)
'  ConnState = csReady
'  xAgentID = 2 ^ Value
'  xAgentHostname = "FAKEAGENT" & Value
'  xID = Value
'
'  ExclusiveMode = False
'  MaxProcesses = 5
'End Property
