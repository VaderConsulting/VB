VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Filter..."
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton optAgent 
      Caption         =   "Use Microsoft Speech API"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   24
      Top             =   240
      Width           =   2295
   End
   Begin VB.OptionButton optAgent 
      Caption         =   "Use Microsoft Agent"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   23
      Top             =   600
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton cmdRevert 
      Caption         =   "Revert"
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5520
      TabIndex        =   21
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame fmeSounds 
      Caption         =   "Sounds"
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton cmdPlay 
         Height          =   315
         Index           =   2
         Left            =   3360
         Picture         =   "frmOptions.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdPlay 
         Height          =   315
         Index           =   1
         Left            =   3360
         Picture         =   "frmOptions.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdPlay 
         Height          =   315
         Index           =   0
         Left            =   3360
         Picture         =   "frmOptions.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Index           =   2
         Left            =   1560
         TabIndex        =   20
         Text            =   "cmbSound"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Text            =   "cmbSound"
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   18
         Text            =   "cmbSound"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblSound 
         Caption         =   "Critical sound"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblSound 
         Caption         =   "Warning sound"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblSound 
         Caption         =   "Information sound"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fmeGeneral 
      Caption         =   "General"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   3855
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Ouput Summary only"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Quiet between 7:00 pm and 6:00 am"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame fmeAutoMute 
      Caption         =   "Dont alert me for:"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   3135
      Begin TabDlg.SSTab tabAlert 
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2566
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Severity"
         TabPicture(0)   =   "frmOptions.frx":0820
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstCriteria"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Groups"
         TabPicture(1)   =   "frmOptions.frx":083C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstGroups"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Systems"
         TabPicture(2)   =   "frmOptions.frx":0858
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstSystems"
         Tab(2).ControlCount=   1
         Begin VB.ListBox lstCriteria 
            Height          =   840
            ItemData        =   "frmOptions.frx":0874
            Left            =   240
            List            =   "frmOptions.frx":0876
            TabIndex        =   28
            Top             =   480
            Width           =   2415
         End
         Begin VB.ListBox lstGroups 
            Height          =   840
            ItemData        =   "frmOptions.frx":0878
            Left            =   -74760
            List            =   "frmOptions.frx":087A
            TabIndex        =   27
            Top             =   480
            Width           =   2415
         End
         Begin VB.ListBox lstSystems 
            Height          =   840
            ItemData        =   "frmOptions.frx":087C
            Left            =   -74760
            List            =   "frmOptions.frx":087E
            TabIndex        =   26
            Top             =   480
            Width           =   2415
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame fmeCharacter 
      Height          =   2295
      Left            =   4080
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      Begin VB.CheckBox chkOptions 
         Caption         =   "Verbose speech output"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Size balloon to text"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Autopace"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Autohide balloon"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Stop between actions"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Word Balloon"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Sounds"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApply_Click()
  SaveSettings
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFilter_Click()
  frmFilter.Show vbModal
End Sub

Private Sub cmdOK_Click()
  SaveSettings
  Unload Me
End Sub

Private Sub cmdPlay_Click(Index As Integer)
  PlayWav App.Path & "\" & cmbSound(Index).Text
End Sub

Private Sub cmdRevert_Click()
  GetSettings
End Sub

Private Sub Form_Load()
  GetSettings
  'Me.Show
  Me.Refresh
End Sub

Sub GetSettings()
  Dim SQL As String, Criteria As String
  Dim adoRS As New ADODB.Recordset
  Dim isOK As Boolean
  
  SQL = "SELECT DISTINCT (System) FROM tblAlerts ORDER BY System"
  adoRS.Open SQL, DSN
  lstSystems.Clear
  Do Until adoRS.EOF
    lstSystems.AddItem adoRS("System")
    adoRS.MoveNext
  Loop
  adoRS.Close
  SQL = "SELECT DISTINCT ([Group]) FROM tblAlerts ORDER BY [Group]"
  adoRS.Open SQL, DSN
  lstGroups.Clear
  Do Until adoRS.EOF
    lstGroups.AddItem adoRS("Group")
    adoRS.MoveNext
  Loop
  adoRS.Close
  SQL = "SELECT DISTINCT (Severity) FROM tblAlerts ORDER BY Severity"
  adoRS.Open SQL, DSN
  lstCriteria.Clear
  Do Until adoRS.EOF
    Select Case adoRS("Severity")
      Case "1"
        Criteria = "Information"
      Case "2"
        Criteria = "Warning"
      Case "3"
        Criteria = "Critical Alert"
      Case Else
        Criteria = "Unknown severity"
    End Select
      lstCriteria.AddItem Criteria
    adoRS.MoveNext
  Loop
  adoRS.Close
  
  ' Populate combo boxes
  cmbSound(0).Clear
  cmbSound(1).Clear
  cmbSound(2).Clear
  cmbSound(0).AddItem "more work.wav"
  cmbSound(0).AddItem "1.wav"
  cmbSound(1).AddItem "im not listening.wav"
  cmbSound(1).AddItem "2.wav"
  cmbSound(2).AddItem "3.wav"
  cmbSound(2).AddItem "4.wav"
  
  '***** Get setting for Agent option
  Misc = GetValue(Host, "Software\Alerter", "Agent", isOK, REG_SZ)
  If Misc = "1" Then
    optAgent(0).Value = True
  Else
    optAgent(1).Value = True
  End If
  
  '***** Get setting for Sounds option
  Misc = GetValue(Host, "Software\Alerter", "Sounds", isOK, REG_SZ)
  chkOptions(1) = Misc
  
  '***** Get setting for Word Balloon option
  Misc = GetValue(Host, "Software\Alerter", "Word Balloon", isOK, REG_SZ)
  chkOptions(2) = Misc
  
  '***** Get setting for Stop Between Actions option
  Misc = GetValue(Host, "Software\Alerter", "Stop Between Actions", isOK, REG_SZ)
  chkOptions(3) = Misc
  
  '***** Get setting for Auto Hide Balloon option
  Misc = GetValue(Host, "Software\Alerter", "Auto Hide Balloon", isOK, REG_SZ)
  chkOptions(4) = Misc
  
  '***** Get setting for Auto Pace option
  Misc = GetValue(Host, "Software\Alerter", "Auto Pace", isOK, REG_SZ)
  chkOptions(5) = Misc
  
  '***** Get setting for Size Balloon to Text option
  Misc = GetValue(Host, "Software\Alerter", "Size Balloon to Text", isOK, REG_SZ)
  chkOptions(6) = Misc
  
  '***** Get setting for Verbose option
  Misc = GetValue(Host, "Software\Alerter", "Verbose", isOK, REG_SZ)
  chkOptions(7) = Misc
  
  '***** Get setting for Quiet option
  Misc = GetValue(Host, "Software\Alerter", "Quiet", isOK, REG_SZ)
  chkGeneral(0) = Misc
  
  '***** Get setting for Output Summary Only option
  Misc = GetValue(Host, "Software\Alerter", "Output Summary Only", isOK, REG_SZ)
  chkGeneral(1) = Misc
  
  '***** Get setting for Information Sound option
  Misc = GetValue(Host, "Software\Alerter", "Information Sound", isOK, REG_SZ)
  cmbSound(0).Text = Misc
  
  '***** Get setting for Warning Sound option
  Misc = GetValue(Host, "Software\Alerter", "Warning Sound", isOK, REG_SZ)
  cmbSound(1).Text = Misc
  
  '***** Get setting for Critical Alert Sound option
  Misc = GetValue(Host, "Software\Alerter", "Critical Alert Sound", isOK, REG_SZ)
  cmbSound(2).Text = Misc
  
End Sub

Sub SaveSettings()
  Dim Host As String
  Host = ""
  If optAgent(0).Value = True Then
    boolAgent = True
    SetValueString Host, "Software\Alerter", "Agent", "1", REG_SZ
  Else
    boolAgent = False
    SetValueString Host, "Software\Alerter", "Agent", "0", REG_SZ
  End If
  
  SetValueString Host, "Software\Alerter", "Sounds", chkOptions(1).Value, REG_SZ
  SetValueString Host, "Software\Alerter", "Word Balloon", chkOptions(2).Value, REG_SZ
  SetValueString Host, "Software\Alerter", "Stop Between Actions", chkOptions(3).Value, REG_SZ
  SetValueString Host, "Software\Alerter", "Auto Hide Balloon", chkOptions(4).Value, REG_SZ
  SetValueString Host, "Software\Alerter", "Auto Pace", chkOptions(5).Value, REG_SZ
  SetValueString Host, "Software\Alerter", "Size Balloon to Text", chkOptions(6).Value, REG_SZ
  SetValueString Host, "Software\Alerter", "Verbose", chkOptions(7).Value, REG_SZ
  SetValueString Host, "Software\Alerter", "Quiet", chkGeneral(0).Value, REG_SZ
  SetValueString Host, "Software\Alerter", "Output Summary Only", chkGeneral(1).Value, REG_SZ
  boolSounds = CBool(chkOptions(1).Value)
  boolWordBalloon = CBool(chkOptions(2).Value)
  boolStopBetweenActions = CBool(chkOptions(3).Value)
  boolAutoHideBalloon = CBool(chkOptions(4).Value)
  boolAutoPace = CBool(chkOptions(5).Value)
  boolSizeBalloontoText = CBool(chkOptions(6).Value)
  boolVerbose = CBool(chkOptions(7).Value)
  boolQuiet = CBool(chkGeneral(0).Value)
  boolOutputSummaryOnly = CBool(chkGeneral(1).Value)
  strInformationSound = cmbSound(0).Text
  strWarningSound = cmbSound(1).Text
  strCriticalSound = cmbSound(2).Text
  SetValueString Host, "Software\Alerter", "Information Sound", strInformationSound, REG_SZ
  SetValueString Host, "Software\Alerter", "Warning Sound", strWarningSound, REG_SZ
  SetValueString Host, "Software\Alerter", "Critical Alert Sound", strCriticalSound, REG_SZ
End Sub

Private Sub optAgent_Click(Index As Integer)
  If optAgent(0).Value = True Then
    fmeCharacter.Enabled = True
    CharLoaded = False
    frmMain.GetAgent strSelectedAgent
    frmMain.Character.Show
    frmMain.mnuLoad.Enabled = True
    For lp = 1 To 7
      chkOptions(lp).Enabled = True
    Next lp
  Else
    fmeCharacter.Enabled = False
    If frmMain.CharLoaded Then
      frmMain.Character.Hide
      frmMain.mnuLoad.Enabled = False
    End If
    For lp = 1 To 7
      chkOptions(lp).Enabled = False
    Next lp
  End If
End Sub
