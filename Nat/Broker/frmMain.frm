VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NAT Broker Process"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddDomain 
      Caption         =   "Add"
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox cmbDomainName 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer tmrWatchdog 
      Interval        =   60000
      Left            =   600
      Top             =   1680
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Stopped"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtNumProcesses 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer tmrMain 
      Left            =   120
      Top             =   1680
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   3600
      X2              =   120
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      X1              =   3600
      X2              =   120
      Y1              =   2180
      Y2              =   2180
   End
   Begin VB.Label lblDomainName 
      Caption         =   "Domain Name"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "NAT State"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Number of running processes"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Clients"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblTCPClients 
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbDomainName_Change()
    If cmbDomainName.ListCount = 1 Then
        cmbDomainName.Text = cmbDomainName.List(0)
    End If
End Sub

Private Sub cmdAddDomain_Click()
    Dim NewDomain As String
    NewDomain = InputBox("Enter a valid Domain to audit", "Enter Domain name", Environ$("USERDNSDOMAIN"))
    If Trim(NewDomain) <> "" Then
        cmbDomainName.AddItem NewDomain
        cmdStart.Enabled = True
    End If
End Sub

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
            lblStatus.Caption = "Paused"
            cmdPause.Caption = "Resume"
        Case False
            lblStatus.Caption = "Working"
            cmdPause.Caption = "Pause"
    End Select
End Sub

Private Sub cmdStart_Click()
    Dim d As Date
    
    If cmbDomainName.Text = "" Then
        MsgBox "Please select a Domain to audit", vbExclamation, "NAT Error"
        Exit Sub
    End If
    
    cmdStart.Enabled = False
    
    If Dir(App.Path & "\collect.exe") <> "" Then
        lblStatus.Caption = "Spawning NAT Collect"
        Me.Refresh
        Shell App.Path & "\collect.exe " & cmbDomainName.Text
        d = Now
        ' Give Collect.exe time to do its stuff
        lblStatus.Caption = "Allowing Collect to initialise"
        Me.Refresh
        Do Until Now > DateAdd("s", 15, d)
            DoEvents
        Loop
        
        If Dir(App.Path & "\ping.exe") <> "" Then
            lblStatus.Caption = "Spawning NAT Ping"
            Me.Refresh
            Shell App.Path & "\ping.exe " & cmbDomainName.Text
            d = Now
            ' Give Ping.exe time to do its stuff
            lblStatus.Caption = "Allowing Ping to initialise"
            Me.Refresh
            Do Until Now > DateAdd("s", 15, d)
                DoEvents
            Loop
        End If
    End If
    
    StartTasks
    cmdStop.Enabled = True
    cmdPause.Enabled = True
End Sub

Private Sub cmdStop_Click()
    Dim s As String, Computername As String
    cmdStop.Enabled = False
    cmdPause.Enabled = False
    StopTasks
    
    On Error Resume Next
    s = Dir("c:\NAT\" & cmbDomainName.Text & "\*.na3", vbNormal)
    
    Do Until s = ""
        Computername = Left(s, InStr(1, s, ".") - 1)
        Name "c:\NAT\" & cmbDomainName.Text & "\" & s As "c:\NAT\" & cmbDomainName.Text & "\" & Computername & ".na2"
        s = Dir
    Loop
    
    lblStatus.Caption = "Stopped"
    cmdStart.Enabled = True
End Sub

Private Sub Form_Load()
    Dim DoStart As Boolean
    Dim d As String
    
    lblStatus.Caption = "Idle"
    Me.Show
    Me.Refresh
    
    'Defaults
    DoStart = False
    LogSeverity = 1
    
    WriteToLog 0, "---------------------------"
    WriteToLog 0, App.EXEName & " Log Opened"
    
    'Parse Command Line
    If (InStr(UCase(Command$), "/START")) Then DoStart = True
    If (InStr(UCase(Command$), "/LOG:1")) Then LogSeverity = 1
    If (InStr(UCase(Command$), "/LOG:2")) Then LogSeverity = 2
    If (InStr(UCase(Command$), "/LOG:3")) Then LogSeverity = 3
    If (InStr(UCase(Command$), "/LOG:4")) Then LogSeverity = 4
    
    lblStatus.Caption = "Retrieving Domains"
    Me.Refresh
    d = Dir(App.Path & "\*.*", vbDirectory)
    
    Do Until d = ""
        If d <> "." And d <> ".." And LCase(Right(d, 3)) <> "exe" And LCase(Right(d, 3)) <> "mdb" And LCase(Right(d, 3)) <> "xml" Then
            cmbDomainName.AddItem UCase(d)
            cmdStart.Enabled = True
        End If
        d = Dir
    Loop
    
    lblStatus.Caption = "Initialising"
    Me.Refresh
    If Not InitObjects Then End
    
    lblStatus.Caption = "Ready"
    Me.Refresh
    
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
    Dim d As Date
    tmrMain.Enabled = False
    NumProcs = Processes.Count
    '
    ' Find and create a new task if expected
    '
    lblStatus.Caption = "Checking for processes"
    Me.Refresh
    
    Do Until (ProcessListFull) Or (DoStop) Or (DoPause) Or Dir(App.Path & "\" & cmbDomainName.Text & "\*.na2") = ""
        TaskID = 1
        'TaskID = FindATask(True)
        If (TaskID <> 0) Then
            lblStatus.Caption = "Creating a Process"
            Me.Refresh
            If Not ExecuteTask(TaskID) Then SetStatus TaskID, -1
        End If
        If TaskID = 0 Then Exit Do
        ' Give each task 2 seconds to startup
        d = Now
        Do Until Now > DateAdd("s", 2, d)
            DoEvents
        Loop
    Loop
    '
    ' Get the number of processes still running
    '
    lblStatus.Caption = "Counting active processes"
    Me.Refresh
    FindCompletedProcesses
    SetTasksState
    If (Processes.Count <> NumProcs) Then txtNumProcesses.Text = CStr(Processes.Count)
    tmrMain.Enabled = True
    lblStatus.Caption = "Working"
    Me.Refresh
End Sub

Private Sub tmrWatchdog_Timer()
    If cmbDomainName.Text = "" Then Exit Sub
    If Dir(App.Path & "\ping.exe") <> "" Then
        lblStatus.Caption = "Spawning NAT Ping"
        Me.Refresh
        Shell App.Path & "\ping.exe " & cmbDomainName.Text
    End If
'    If Dir(App.Path & "\reader.exe") <> "" Then
'        lblStatus.Caption = "Spawning NAT Reader"
'        Me.Refresh
'        Shell App.Path & "\reader.exe " & cmbDomainName.Text
'    End If
End Sub
