VERSION 5.00
Object = "{BC865B94-984A-11D3-9875-DD95F105434B}#2.0#0"; "tsksch.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Schedule (1.0B)"
   ClientHeight    =   6870
   ClientLeft      =   2160
   ClientTop       =   2250
   ClientWidth     =   6375
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   6375
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5640
      TabIndex        =   19
      Top             =   7200
      Width           =   735
   End
   Begin axSchedule.Schedule Schedule1 
      Left            =   120
      Top             =   6480
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   6270
   End
   Begin VB.Frame Frame1 
      Height          =   3930
      Left            =   30
      TabIndex        =   1
      Top             =   2415
      Width           =   6270
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run Now!"
         Height          =   495
         Left            =   135
         TabIndex        =   18
         Top             =   3300
         Width           =   960
      End
      Begin VB.TextBox txtSchedule 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1305
         Left            =   1230
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   2445
         Width           =   4665
      End
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   1230
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   705
         Width           =   4665
      End
      Begin VB.Label lbLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule"
         Height          =   195
         Index           =   7
         Left            =   495
         TabIndex        =   16
         Top             =   2430
         Width           =   675
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   5
         Left            =   1230
         TabIndex        =   15
         Top             =   2205
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   4
         Left            =   1230
         TabIndex        =   14
         Top             =   1965
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   3
         Left            =   1230
         TabIndex        =   13
         Top             =   1710
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   2
         Left            =   1230
         TabIndex        =   12
         Top             =   1455
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   1230
         TabIndex        =   11
         Top             =   465
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   1230
         TabIndex        =   10
         Top             =   225
         Width           =   45
      End
      Begin VB.Label lbLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Creator:"
         Height          =   195
         Index           =   6
         Left            =   630
         TabIndex        =   9
         Top             =   2205
         Width           =   555
      End
      Begin VB.Label lbLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Next Runtime:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   8
         Top             =   1965
         Width           =   1005
      End
      Begin VB.Label lbLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Runtime:"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   6
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label lbLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flags:"
         Height          =   195
         Index           =   3
         Left            =   765
         TabIndex        =   5
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label lbLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         Height          =   195
         Index           =   2
         Left            =   405
         TabIndex        =   4
         Top             =   735
         Width           =   780
      End
      Begin VB.Label lbLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Command Line:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   465
         Width           =   1095
      End
      Begin VB.Label lbLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Program File:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   7800
      X2              =   0
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6480
      Width           =   6135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TASK_IDLE = 267008
Const TASK_RUNNING = 267009

Private Sub cmdRun_Click()

    Schedule1(List1.ListIndex + 1).Run
    
End Sub


Private Sub Command1_Click()
    'Dim j As Job
    'Set j.ApplicationName = "c:\winnt\system32\calc.exe"
    'Set j.Creator = "POLICE\PD71044A"
    'j.Name = "Calculator"
    'Schedule1.CreateTask "Calculator"
End Sub

Private Sub List1_Click()

    lblInfo(0).Caption = Schedule1(List1.ListIndex + 1).ApplicationName
    lblInfo(1).Caption = Schedule1(List1.ListIndex + 1).CommandLine
    txtComments.Text = Schedule1(List1.ListIndex + 1).Comment
    lblInfo(2).Caption = Flags(Schedule1(List1.ListIndex + 1).Flags)
    lblInfo(3).Caption = Format(Schedule1(List1.ListIndex + 1).LastRunTime, "HH:MM:SS DD-MMM-YYYY")
    lblInfo(4).Caption = Format(Schedule1(List1.ListIndex + 1).NextRunTime, "HH:MM:SS DD-MMM-YYYY")
    lblInfo(5).Caption = Schedule1(List1.ListIndex + 1).Creator
    
    Select Case Schedule1(List1.ListIndex + 1).Status
        Case TASK_RUNNING
            Debug.Print "This task is running"
        Case TASK_IDLE
            Debug.Print "This task is idle"
    End Select
    
    Dim Tgr As Trigger
    
    txtSchedule.Text = ""
    
    For Each Tgr In Schedule1(List1.ListIndex + 1).Triggers
        txtSchedule.Text = txtSchedule.Text & Tgr.Text & vbNewLine
    Next
    
End Sub

Private Function Flags(ByVal F As Long) As String
Dim T As String
    
    If (F And saDeleteWhenDone) Then T = T & "Delete When Done - "
    If (F And saDisabled) Then T = T & "Disabled - "
    If (F And saDontStartIfOnBatteries) Then T = T & "Don't Start If On Batteries - "
    If (F And saHidden) Then T = T & "Hidden - "
    If (F And saInteractive) Then T = T & "Interactive - "
    If (F And saKillIfGoingOnBatteries) Then T = T & "Kill If Going On Batteries - "
    If (F And saKillOnIdleEnd) Then T = T & "Kill On Idle End - "
    If (F And saStartOnlyIfIdle) Then T = T & "Start Only If Idle - "
    
    Flags = T
    
End Function

Private Sub mnuConnect_Click()
    Dim Host As String
    Me.Show
    Me.Refresh
    Host = InputBox("Enter a computername to connect to", "Enter", Host)
    If Left(Host, 2) <> "\\" Then Host = "\\" & Host
    lblStatus = "Contacting " & Host
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If Host <> "" Then
        Schedule1.TargetComputer = Host
        Schedule1.Refresh
        Dim Tsk As Job
        
        lblStatus = "Enumerating tasks on " & Host
        lblStatus.Refresh
        For Each Tsk In Schedule1
            List1.AddItem Tsk.Name
            'Tsk.ShowProperties Me.hWnd
        Next
        lblStatus = "Done"
    Else
        lblStatus = "'" & Host & "' is not a valid name"
    End If
    lblStatus.Refresh
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuExit_Click()
    End
End Sub
