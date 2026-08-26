VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ServiceMan"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5490
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optManualRefresh 
      Caption         =   "Manual Refresh"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1680
      Width           =   1575
   End
   Begin VB.OptionButton optRefresh 
      Caption         =   "Auto Refresh"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdControlService 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdControlService 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdControlService 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdControlService 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.ListBox lstService 
      Height          =   2010
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtComputer 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   2640
   End
   Begin VB.Label lblInfo 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   5295
   End
   Begin VB.Label lblState 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblService 
      Caption         =   "Service"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblComputer 
      Caption         =   "Computer"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()

  Dim objWMIService As Object
  Dim colServiceList As Object
  Dim objService As Variant

    If Trim$(txtComputer) = "" Then Exit Sub ':( Expand Structure or consider reversing Condition

    lstService.Clear
    lblInfo = "Connecting..."
    frmMain.Refresh

    Screen.MousePointer = vbHourglass
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & txtComputer & "\root\cimv2")
    Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service")

    lblInfo = "Enumerating Services..."
    For Each objService In colServiceList
        lstService.AddItem objService.Name
        frmMain.Refresh
    Next ':( Repeat For-Variable: OBJSERVICE

    lblInfo = "Idle"
    Screen.MousePointer = vbDefault
    Set objService = Nothing
    Set colServiceList = Nothing
    Set objWMIService = Nothing

End Sub

Private Sub cmdControlService_Click(Index As Integer)

  Dim objWMIService As Object
  Dim colServiceList As Object
  Dim objService As Variant

    lblInfo = "Connecting..."
    frmMain.Refresh

    Screen.MousePointer = vbHourglass

    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & txtComputer & "\root\cimv2")
    Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service Where Name='" & lstService.List(lstService.ListIndex) & "'")

    For Each objService In colServiceList
        Select Case Index
          Case 0 ' Start
            objService.startservice
          Case 1 ' Stop
            objService.stopservice
          Case 2 ' Pause
            objService.pauseservice
          Case 3 ' Continue
            objService.resumeservice
        End Select
    Next ':( Repeat For-Variable: OBJSERVICE

    Set objWMIService = Nothing
    Set colServiceList = Nothing
    Set objService = Nothing

    Screen.MousePointer = vbDefault

End Sub

Private Sub lstService_Click()

    If lstService.List(lstService.ListIndex) = "" Then
        tmrRefresh.Enabled = False
        Exit Sub '>---> Bottom
      Else 'NOT LSTSERVICE.LIST(LSTSERVICE.LISTINDEX)...
        tmrRefresh.Enabled = True
        lblState.Caption = ""
    End If

End Sub

Private Sub optManualRefresh_Click()

    tmrRefresh.Enabled = False

End Sub

Private Sub optRefresh_Click()

    If lstService.List(lstService.ListIndex) <> "" Then
        tmrRefresh.Enabled = True
      Else 'NOT LSTSERVICE.LIST(LSTSERVICE.LISTINDEX)...
        tmrRefresh.Enabled = False
    End If

End Sub

Private Sub tmrRefresh_Timer()

  Dim objWMIService As Object
  Dim colServiceList As Object
  Dim objService As Variant

    tmrRefresh.Enabled = False

    lblInfo = "Connecting..."
    frmMain.Refresh

    Screen.MousePointer = vbHourglass

    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & txtComputer & "\root\cimv2")
    Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service Where Name='" & lstService.List(lstService.ListIndex) & "'")

    lblInfo = "Refreshing..."
    frmMain.Refresh

    For Each objService In colServiceList
        lblState.Caption = CStr(objService.State)
        Select Case lblState.Caption
          Case "Running"
            If objService.acceptstop Then
                cmdControlService(0).Enabled = False
                cmdControlService(1).Enabled = True
                cmdControlService(3).Enabled = False
            End If
            If objService.acceptpause Then
                cmdControlService(2).Enabled = True
              Else 'OBJSERVICE.ACCEPTPAUSE = FALSE/0
                cmdControlService(2).Enabled = False
            End If
          Case "Stopped"
            cmdControlService(0).Enabled = True
            cmdControlService(1).Enabled = False
            cmdControlService(2).Enabled = False
            cmdControlService(3).Enabled = False
          Case "Paused"
            cmdControlService(0).Enabled = False
            cmdControlService(1).Enabled = True
            cmdControlService(2).Enabled = False
            cmdControlService(3).Enabled = True
        End Select
    Next ':( Repeat For-Variable: OBJSERVICE

    lblInfo = "Idle"

    tmrRefresh.Enabled = True
    frmMain.Refresh
    Screen.MousePointer = vbDefault

    Set objService = Nothing
    Set colServiceList = Nothing
    Set objWMIService = Nothing

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jul-28 12:24) 1 + 153 = 154 Lines
