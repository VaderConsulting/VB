VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TaskScheduler Sample"
   ClientHeight    =   6390
   ClientLeft      =   3765
   ClientTop       =   2565
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   6345
   Begin MSComDlg.CommonDialog cmdlgOpenfile 
      Left            =   2925
      Top             =   2955
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   5000
      Left            =   5490
      Top             =   1950
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run Now!"
      Height          =   375
      Left            =   5190
      TabIndex        =   21
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cdmDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5190
      TabIndex        =   20
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton cmdProps 
      Caption         =   "&Properties"
      Height          =   375
      Left            =   5190
      TabIndex        =   19
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Add New"
      Height          =   375
      Left            =   5190
      TabIndex        =   18
      Top             =   60
      Width           =   1095
   End
   Begin VB.ListBox lstTasks 
      Height          =   2400
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   5100
   End
   Begin VB.Frame Frame1 
      Height          =   3930
      Left            =   30
      TabIndex        =   1
      Top             =   2415
      Width           =   6270
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Schedule As New TaskScheduler.Schedule
Sub RefreshList()
Dim Tsk As Job

   lstTasks.Clear
   
   Schedule.Refresh
   
   For Each Tsk In Schedule
      lstTasks.AddItem Tsk.Name
   Next
   
End Sub

Private Sub cdmDelete_Click()
   
   If lstTasks.ListIndex > -1 Then
   
      If MsgBox("The operation can't be undone. Are you sure you want to delete """ & lstTasks.List(lstTasks.ListIndex) & """?", vbYesNo) = vbYes Then
         Schedule.Delete lstTasks.List(lstTasks.ListIndex)
         RefreshList
      End If
   
   End If
   
End Sub

Private Sub cmdNew_Click()
Dim sTaskName As String
Dim oJob As Job

   cmdlgOpenfile.Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
   
   On Error Resume Next
   
   cmdlgOpenfile.ShowOpen
   
   If Err = 0 Then
   
      sTaskName = InputBox("Type the task name", , "My New Task")
   
      On Error GoTo ErrHandler
      
      Set oJob = Schedule.CreateTask(sTaskName)
      
      With oJob
               
         ' Set the exe path
         .ApplicationName = cmdlgOpenfile.FileName
      
         ' Set other properties
         .Creator = "Task Scheduler control sample"
      
         ' Add a schedule
         With .Triggers
      
            With .Add
         
               .TriggerType = ttDaily     ' Daily
               .BeginDay = Now            ' Starting today
               .StartTime = Now           ' At 10:00 AM
               
            End With
            
         End With
                  
         .Flags = tfDisabled
         
         ' Save the task
         .Save
         
         .ShowProperties
      
      End With
   
      RefreshList
      
   End If
   
ErrHandler:

End Sub

Private Sub cmdProps_Click()

   If lstTasks.ListIndex > -1 Then
   
      Schedule.Job(lstTasks.List(lstTasks.ListIndex)).ShowProperties Me.hWnd
      
      RefreshList
   
   End If
   
End Sub


Private Sub cmdRun_Click()

    Schedule(lstTasks.ListIndex + 1).Run
    
End Sub


Private Sub Form_Load()

   RefreshList
    
End Sub


Private Sub lstTasks_Click()

    lblInfo(0).Caption = Schedule(lstTasks.ListIndex + 1).ApplicationName
    lblInfo(1).Caption = Schedule(lstTasks.ListIndex + 1).CommandLine
    txtComments.Text = Schedule(lstTasks.ListIndex + 1).Comment
    lblInfo(2).Caption = Flags(Schedule(lstTasks.ListIndex + 1).Flags)
    lblInfo(3).Caption = Format(Schedule(lstTasks.ListIndex + 1).LastRunTime, "HH:MM:SS DD-MMM-YYYY")
    lblInfo(4).Caption = Format(Schedule(lstTasks.ListIndex + 1).NextRunTime, "HH:MM:SS DD-MMM-YYYY")
    lblInfo(5).Caption = Schedule(lstTasks.ListIndex + 1).Creator
    
    Dim Tgr As Trigger
    
    txtSchedule.Text = ""
    
    For Each Tgr In Schedule(lstTasks.ListIndex + 1).Triggers
        txtSchedule.Text = txtSchedule.Text & Tgr.Text & vbNewLine
    Next
    
End Sub

Private Function Flags(ByVal F As Long) As String
Dim T As String
    
    If (F And tfDeleteWhenDone) Then T = T & "Delete When Done - "
    If (F And tfDisabled) Then T = T & "Disabled - "
    If (F And tfDontStartIfOnBatteries) Then T = T & "Don't Start If On Batteries - "
    If (F And tfHidden) Then T = T & "Hidden - "
    If (F And tfInteractive) Then T = T & "Interactive - "
    If (F And tfKillIfGoingOnBatteries) Then T = T & "Kill If Going On Batteries - "
    If (F And tfKillOnIdleEnd) Then T = T & "Kill On Idle End - "
    If (F And tfStartOnlyIfIdle) Then T = T & "Start Only If Idle - "
    
    Flags = T
    
End Function

Private Sub tmrRefresh_Timer()

   RefreshList
   
End Sub


