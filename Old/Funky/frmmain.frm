VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distributed Tasks Client"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "DistroTaskAgent"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstConnections 
      Height          =   2595
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock tcpMain 
      Index           =   0
      Left            =   2280
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "Idle"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtNumProcesses 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer tmrWatchDog 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3240
      Top             =   1080
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Left            =   2760
      Top             =   1080
   End
   Begin VB.Label lblTCPClients 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnStart_Click()
  StartDistroTasks
End Sub

Private Sub btnStop_Click()
  StopDistroTasks
End Sub

Private Sub Form_Load()
  Dim DoStart As Boolean
  
  'Defaults
  DoStart = False
  LogSeverity = 0
  
  WriteToLog 0, "---------------------------"
  WriteToLog 0, "DISTROCLIENT.EXE Log Opened"
   
  'Parse Command Line
  If (InStr(UCase(Command$), "/START")) Then DoStart = True
  If (InStr(UCase(Command$), "/LOG:1")) Then LogSeverity = 1
  If (InStr(UCase(Command$), "/LOG:2")) Then LogSeverity = 2
  If (InStr(UCase(Command$), "/LOG:3")) Then LogSeverity = 3
  If (InStr(UCase(Command$), "/LOG:4")) Then LogSeverity = 4
  
  If Not InitObjects Then End
  If DoStart Then StartDistroTasks
  tmrWatchDog.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  WriteToLog 0, "DISTROCLIENT.EXE Log Closed"
  CloseObjects
End Sub

Public Sub tcpMain_Close(Index As Integer)
  Dim TCPID As Integer
  Dim RIP As String
  Dim RPort As String
  
  TCPID = FindTCPConnFromID(Index)
  If (TCPID <> -1) Then TCPConns.Remove TCPID
  RIP = tcpMain(Index).RemoteHostIP
  RPort = tcpMain(Index).RemotePort
  Unload tcpMain(Index)
  Debug.Print "Closing Connection " & Index
  DoEvents
  
  WriteToLog 1, "TCP Connection Closed (" & RIP & ":" & RPort & ")"
  SetDistroTasksState
End Sub

Private Sub tcpMain_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  Dim oTCP As clsTCP
  Dim newInstanceIndex As Integer
  If (TCPConns.Count >= MaxTCPConns) Then Exit Sub

  ' Load a new instance of the control to service the connection.
  ' The variable newInstanceIndex maintains the current index of the
  ' next connection to load - this way we don't accidently use an
  ' index already created. Remember to Unload the control when the
  ' connection is closed
  newInstanceIndex = tcpMain.Count
  
  Load tcpMain(newInstanceIndex)
  ' Pass the value of the requestID parameter to the
  ' Accept method.
  tcpMain(newInstanceIndex).Accept requestID
  
  WriteToLog 1, "TCP Connection Opened (" & tcpMain(Index).RemoteHostIP & ":" & tcpMain(Index).RemotePort & ")"
  
  ' THESE LINES MUST BE LAST!!!!!
  Set oTCP = New clsTCP
  oTCP.SetupTCPInfo newInstanceIndex, tcpMain(Index).RemoteHostIP
  TCPConns.Add oTCP
  SetDistroTasksState
End Sub

Private Sub tcpMain_DataArrival(Index As Integer, ByVal BytesTotal As Long)
  Dim TCPID As Integer
  
  TCPID = FindTCPConnFromID(Index)
  If (TCPID <> -1) Then TCPConns.Item(TCPID).DataArrival BytesTotal
End Sub

Private Sub tcpMain_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  WriteToLog 0, "TCP Error (" & Number & ") " & Description
  CancelDisplay = True
End Sub

Private Sub tmrMain_Timer()
  Dim TaskID As Integer
  Dim NumProcs As Integer
  NumProcs = Processes.Count
  Do
    TaskID = FindATask(True)
    If (TaskID <> 0) Then
      If Not ExecuteTask(TaskID) Then SetStatus TaskID, -1
    End If
  Loop Until (TaskID = 0) Or ProcessListFull
  If (Processes.Count <> NumProcs) Then txtNumProcesses.Text = CStr(Processes.Count)
End Sub

Private Sub tmrWatchDog_Timer()
  Dim NumProcs As Integer
  Dim oTCP As clsTCP, CurrentConns As Integer, i As Integer
  If tmrMain.Enabled Then ProcessAgentKeepAlive ' DistroTasks is running
  DoEvents
  NumProcs = Processes.Count
  FindCompletedProcesses
  If (Processes.Count <> NumProcs) Then
    txtNumProcesses.Text = CStr(Processes.Count)
    SetDistroTasksState
  End If
  i = 1
  CurrentConns = frmMain.tcpMain().Count - 1
  'Debug.Print "Looking at connections"
  frmMain.lstConnections.Clear
  'If CurrentConns = 0 Then Exit Sub
  Do Until i > CurrentConns
    'Debug.Print i & " =" & frmMain.tcpMain(i).RemoteHostIP
    frmMain.lstConnections.AddItem frmMain.tcpMain(i).RemoteHostIP
    i = i + 1
    DoEvents
  Loop
  
End Sub
