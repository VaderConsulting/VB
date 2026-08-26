VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Time Check"
   ClientHeight    =   6525
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5700
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   3000
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".cfg"
   End
   Begin VB.CheckBox chkDelta 
      Caption         =   "Retrieve Delta times"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "^"
      Default         =   -1  'True
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox txtServerName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   1815
   End
   Begin VB.ListBox lstTimes 
      Height          =   5715
      Left            =   2520
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   3135
   End
   Begin VB.ListBox lstServers 
      Height          =   5715
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   7560
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   7560
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Config"
         Begin VB.Menu mnuLoad 
            Caption         =   "Load"
         End
         Begin VB.Menu mnuSave 
            Caption         =   "Save"
            Enabled         =   0   'False
         End
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
Private Sub cmdAdd_Click()
  If txtServerName = "" Then Exit Sub
  mnuStart.Enabled = True
  mnuSave.Enabled = True
  lstServers.AddItem UCase(txtServerName)
  txtServerName = ""
End Sub

Private Sub mnuExit_Click()
  Unload Me
  End
End Sub

Private Sub mnuExport_Click()
  Open "c:\temp\Server Times.txt" For Output As #1
    For lp = 0 To lstTimes.ListCount - 1
      Print #1, lstTimes.List(lp)
    Next lp
  Close 1
  MsgBox "Results saved to c:\temp\Server Times.txt"
End Sub

Private Sub mnuLoad_Click()
  On Error GoTo OpenError
  With cdlFiles
    .DefaultExt = ".cfg"
    .Filter = "Config files (*.cfg)|*.cfg"
    .DialogTitle = "Load config file"
    .CancelError = True
    .InitDir = "c:\temp"
    .ShowOpen
  End With
  If cdlFiles.FileName <> "" Then
    lstServers.Clear
    Open cdlFiles.FileName For Input As #1
      Do Until EOF(1)
        Line Input #1, TimeCheckInfo
        Servername = Trim(TimeCheckInfo)
        If Servername <> "" Then lstServers.AddItem Servername
      Loop
    Close 1
  End If
  frmMain.mnuStart.Enabled = True
  Exit Sub
OpenError:
  If Err.Number <> 32755 Then
    Debug.Print Err.Number
    MsgBox "Unknown error " & Err.Number
  Else
    ' Not a problem
  End If
End Sub

Private Sub mnuSave_Click()
  On Error GoTo OpenError
  With cdlFiles
    .DefaultExt = ".cfg"
    .Filter = "Config files (*.cfg)|*.cfg"
    .DialogTitle = "Save config file"
    .CancelError = True
    .InitDir = "c:\temp"
    .ShowOpen
  End With
  If cdlFiles.FileName <> "" Then
    'lstServers.Clear
    Open cdlFiles.FileName For Output As #1
      For lp = 0 To lstServers.ListCount - 1
        Print #1, lstServers.List(lp)
      Next lp
    Close 1
  End If
  Exit Sub
OpenError:
  If Err.Number <> 32755 Then
    Debug.Print Err.Number
    MsgBox "Unknown error " & Err.Number
  Else
    ' Not a problem
  End If
End Sub

Private Sub mnuStart_Click()
  Dim t As New NetHostTime.HostTime
  lstTimes.Clear
  For lp = 0 To lstServers.ListCount - 1
    If chkDelta.Value = vbChecked Then
      lstTimes.AddItem "Local " & Format(Time, "HH:MM:SS") & "," & lstServers.List(lp) & " " & Format(t.GetTime(lstServers.List(lp), True), "HH:MM:SS")
    Else
      lstTimes.AddItem "Local " & Format(Time, "HH:MM:SS") & "," & lstServers.List(lp) & " " & Format(t.GetTime(lstServers.List(lp)), "HH:MM:SS")
    End If
    lstTimes.Refresh
  Next lp
  mnuExport.Enabled = True
End Sub
