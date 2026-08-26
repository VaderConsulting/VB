VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Marshall"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWatchdog 
      Interval        =   5000
      Left            =   3120
      Top             =   240
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtNumProcesses 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer tmrMain 
      Left            =   2640
      Top             =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Client State"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Number of Processes"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Clients"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblTCPClients 
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    'StopTasks
    cmdStop_Click
    Unload Me
    End
End Sub

Private Sub cmdPause_Click()
    DoPause = Not DoPause
    Select Case DoPause
        Case True
            cmdPause.Caption = "Resume"
        Case False
            cmdPause.Caption = "Pause"
    End Select
End Sub

Private Sub cmdStart_Click()
    cmdStart.Enabled = False
    StartTasks
    cmdStop.Enabled = True
    cmdPause.Enabled = True
End Sub

Private Sub cmdStop_Click()
    cmdStop.Enabled = False
    cmdPause.Enabled = False
    StopTasks
    cmdStart.Enabled = True
End Sub

Private Sub Form_Load()
    Dim DoStart As Boolean
    
    'Defaults
    DoStart = True
    LogSeverity = 4
    
    WriteToLog 0, "---------------------------"
    WriteToLog 0, App.EXEName & " Log Opened"
    
    'Parse Command Line
    If (InStr(UCase(Command$), "/START")) Then DoStart = True
    If (InStr(UCase(Command$), "/LOG:1")) Then LogSeverity = 1
    If (InStr(UCase(Command$), "/LOG:2")) Then LogSeverity = 2
    If (InStr(UCase(Command$), "/LOG:3")) Then LogSeverity = 3
    If (InStr(UCase(Command$), "/LOG:4")) Then LogSeverity = 4
    
    If Not InitObjects Then End
    If DoStart Then cmdStart_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'StopTasks
    cmdStop_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteToLog 0, App.EXEName & " Log Closed"
    CloseObjects
End Sub

Private Sub tmrMain_Timer()
    Dim TaskID As Integer
    Dim NumProcs As Integer
    tmrMain.Enabled = False
    NumProcs = Processes.Count
    '
    ' Find and create a new task if expected
    '

    Do Until (ProcessListFull) Or (DoStop) Or (DoPause)
        TaskID = FindATask(True)
        If (TaskID <> 0) Then
            If Not ExecuteTask(TaskID) Then SetStatus TaskID, -1
        End If
        If TaskID = 0 Then Exit Do
    Loop
    '
    ' Get the number of processes still running
    '
    FindCompletedProcesses
    SetTasksState
    If (Processes.Count <> NumProcs) Then txtNumProcesses.Text = CStr(Processes.Count)
    tmrMain.Enabled = True
End Sub
