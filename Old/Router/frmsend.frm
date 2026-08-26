VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSend 
   Caption         =   "Router"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSecurity 
      Height          =   285
      Left            =   15000
      TabIndex        =   39
      Text            =   "1"
      Top             =   10800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtLog 
      Height          =   285
      Left            =   13800
      TabIndex        =   38
      Top             =   9120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdTreeview 
      Caption         =   "Show treeview"
      Height          =   255
      Left            =   13800
      TabIndex        =   37
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Frame fmeGroupnames 
      Caption         =   "Groups"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   36
      Top             =   8040
      Width           =   3375
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   9
         Left            =   1800
         TabIndex        =   212
         ToolTipText     =   "Right click to modify group names"
         Top             =   1680
         Width           =   1485
      End
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   211
         ToolTipText     =   "Right click to modify group names"
         Top             =   1320
         Width           =   1485
      End
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   210
         ToolTipText     =   "Right click to modify group names"
         Top             =   960
         Width           =   1485
      End
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   209
         ToolTipText     =   "Right click to modify group names"
         Top             =   600
         Width           =   1485
      End
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   208
         ToolTipText     =   "Right click to modify group names"
         Top             =   240
         Width           =   1485
      End
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Right click to modify group names"
         Top             =   1680
         Width           =   1485
      End
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Right click to modify group names"
         Top             =   1320
         Width           =   1485
      End
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Right click to modify group names"
         Top             =   960
         Width           =   1485
      End
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Right click to modify group names"
         Top             =   600
         Width           =   1485
      End
      Begin VB.CheckBox chkGroup 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Right click to modify group names"
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "Select &none"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Clear all individuals"
      Top             =   10560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Select &all"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Select all individuals"
      Top             =   10200
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Show &help"
      Height          =   255
      Left            =   13800
      TabIndex        =   26
      ToolTipText     =   "Show Welcome.txt"
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear pre&view"
      Height          =   255
      Left            =   13800
      TabIndex        =   25
      ToolTipText     =   "Remove file from preview window"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame fmeFiles 
      Caption         =   "Signals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3480
      TabIndex        =   30
      Top             =   8040
      Width           =   2655
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Change directory"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Change file directory"
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox chkArchive 
         Caption         =   "Archive automatically"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Send distributed files to archive directory"
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "De&lete"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         ToolTipText     =   "Delete file"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdRename 
         Caption         =   "&Rename"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Rename file"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.FileListBox filPath 
         Height          =   1260
         Left            =   120
         TabIndex        =   0
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtExtension 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         ToolTipText     =   "Select the type of files to display"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblExtension 
         Alignment       =   2  'Center
         Caption         =   "Default e&xtension"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "Don&e"
      Height          =   855
      Left            =   14040
      Picture         =   "frmSend.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "End program"
      Top             =   10200
      Width           =   855
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Distribute"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      Picture         =   "frmSend.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Distribute file"
      Top             =   10200
      Width           =   855
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&Info"
      Height          =   855
      Left            =   2520
      Picture         =   "frmSend.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "About this program"
      Top             =   10200
      Width           =   855
   End
   Begin VB.Frame fmeGroups 
      Caption         =   "Group membership"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6240
      TabIndex        =   33
      Top             =   8040
      Width           =   7455
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         ToolTipText     =   "Save group membership details"
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "<"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   17
         ToolTipText     =   "Remove from Group"
         Top             =   1920
         Width           =   495
      End
      Begin TabDlg.SSTab tabGroups 
         Height          =   2445
         Left            =   3960
         TabIndex        =   19
         Top             =   480
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   4313
         _Version        =   327681
         Tabs            =   10
         TabsPerRow      =   5
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "1"
         TabPicture(0)   =   "frmSend.frx":1108
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstGroup(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "2"
         TabPicture(1)   =   "frmSend.frx":1124
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstGroup(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "3"
         TabPicture(2)   =   "frmSend.frx":1140
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstGroup(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "4"
         TabPicture(3)   =   "frmSend.frx":115C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lstGroup(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "5"
         TabPicture(4)   =   "frmSend.frx":1178
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lstGroup(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "6"
         TabPicture(5)   =   "frmSend.frx":1194
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "lstGroup(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "7"
         TabPicture(6)   =   "frmSend.frx":11B0
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "lstGroup(6)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "8"
         TabPicture(7)   =   "frmSend.frx":11CC
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "lstGroup(7)"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "9"
         TabPicture(8)   =   "frmSend.frx":11E8
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "lstGroup(8)"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "10"
         TabPicture(9)   =   "frmSend.frx":1204
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "lstGroup(9)"
         Tab(9).ControlCount=   1
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   9
            Left            =   -74760
            TabIndex        =   67
            Top             =   720
            Width           =   2775
         End
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   8
            Left            =   -74760
            TabIndex        =   66
            Top             =   720
            Width           =   2775
         End
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   7
            Left            =   -74760
            TabIndex        =   65
            Top             =   720
            Width           =   2775
         End
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   6
            Left            =   -74760
            TabIndex        =   64
            Top             =   720
            Width           =   2775
         End
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   5
            Left            =   -74760
            TabIndex        =   63
            Top             =   720
            Width           =   2775
         End
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   4
            Left            =   -74760
            TabIndex        =   24
            Top             =   720
            Width           =   2775
         End
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   3
            Left            =   -74760
            TabIndex        =   23
            Top             =   720
            Width           =   2775
         End
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   2
            Left            =   -74760
            TabIndex        =   22
            Top             =   720
            Width           =   2775
         End
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   1
            Left            =   -74760
            TabIndex        =   21
            Top             =   720
            Width           =   2775
         End
         Begin VB.ListBox lstGroup 
            Height          =   1620
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   2775
         End
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   ">"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   16
         ToolTipText     =   "Add to Group"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ListBox lstIndividuals 
         Height          =   2400
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Individuals to choose from"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblGroups 
         Alignment       =   2  'Center
         Caption         =   "Groups"
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblIndividuals 
         Alignment       =   2  'Center
         Caption         =   "Individuals"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fmeRecipients 
      Caption         =   "Recipients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   6735
      Begin TabDlg.SSTab tabRecipients 
         Height          =   7575
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   13361
         _Version        =   327681
         Tabs            =   8
         TabsPerRow      =   8
         TabHeight       =   520
         TabCaption(0)   =   "1-20"
         TabPicture(0)   =   "frmSend.frx":1220
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkPersonnel(19)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkPersonnel(18)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkPersonnel(17)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkPersonnel(16)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkPersonnel(15)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkPersonnel(14)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkPersonnel(13)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkPersonnel(12)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkPersonnel(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "chkPersonnel(1)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkPersonnel(2)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkPersonnel(3)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "chkPersonnel(4)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkPersonnel(6)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chkPersonnel(7)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "chkPersonnel(8)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "chkPersonnel(9)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "chkPersonnel(5)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "chkPersonnel(11)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "chkPersonnel(10)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).ControlCount=   20
         TabCaption(1)   =   "21-40"
         TabPicture(1)   =   "frmSend.frx":123C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkPersonnel(39)"
         Tab(1).Control(1)=   "chkPersonnel(38)"
         Tab(1).Control(2)=   "chkPersonnel(37)"
         Tab(1).Control(3)=   "chkPersonnel(36)"
         Tab(1).Control(4)=   "chkPersonnel(35)"
         Tab(1).Control(5)=   "chkPersonnel(34)"
         Tab(1).Control(6)=   "chkPersonnel(33)"
         Tab(1).Control(7)=   "chkPersonnel(32)"
         Tab(1).Control(8)=   "chkPersonnel(31)"
         Tab(1).Control(9)=   "chkPersonnel(30)"
         Tab(1).Control(10)=   "chkPersonnel(29)"
         Tab(1).Control(11)=   "chkPersonnel(28)"
         Tab(1).Control(12)=   "chkPersonnel(27)"
         Tab(1).Control(13)=   "chkPersonnel(26)"
         Tab(1).Control(14)=   "chkPersonnel(25)"
         Tab(1).Control(15)=   "chkPersonnel(24)"
         Tab(1).Control(16)=   "chkPersonnel(23)"
         Tab(1).Control(17)=   "chkPersonnel(22)"
         Tab(1).Control(18)=   "chkPersonnel(21)"
         Tab(1).Control(19)=   "chkPersonnel(20)"
         Tab(1).ControlCount=   20
         TabCaption(2)   =   "41-60"
         TabPicture(2)   =   "frmSend.frx":1258
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkPersonnel(40)"
         Tab(2).Control(1)=   "chkPersonnel(41)"
         Tab(2).Control(2)=   "chkPersonnel(42)"
         Tab(2).Control(3)=   "chkPersonnel(43)"
         Tab(2).Control(4)=   "chkPersonnel(44)"
         Tab(2).Control(5)=   "chkPersonnel(45)"
         Tab(2).Control(6)=   "chkPersonnel(46)"
         Tab(2).Control(7)=   "chkPersonnel(47)"
         Tab(2).Control(8)=   "chkPersonnel(48)"
         Tab(2).Control(9)=   "chkPersonnel(49)"
         Tab(2).Control(10)=   "chkPersonnel(50)"
         Tab(2).Control(11)=   "chkPersonnel(51)"
         Tab(2).Control(12)=   "chkPersonnel(52)"
         Tab(2).Control(13)=   "chkPersonnel(53)"
         Tab(2).Control(14)=   "chkPersonnel(54)"
         Tab(2).Control(15)=   "chkPersonnel(55)"
         Tab(2).Control(16)=   "chkPersonnel(56)"
         Tab(2).Control(17)=   "chkPersonnel(57)"
         Tab(2).Control(18)=   "chkPersonnel(58)"
         Tab(2).Control(19)=   "chkPersonnel(59)"
         Tab(2).ControlCount=   20
         TabCaption(3)   =   "61-80"
         TabPicture(3)   =   "frmSend.frx":1274
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chkPersonnel(60)"
         Tab(3).Control(1)=   "chkPersonnel(61)"
         Tab(3).Control(2)=   "chkPersonnel(62)"
         Tab(3).Control(3)=   "chkPersonnel(63)"
         Tab(3).Control(4)=   "chkPersonnel(64)"
         Tab(3).Control(5)=   "chkPersonnel(65)"
         Tab(3).Control(6)=   "chkPersonnel(66)"
         Tab(3).Control(7)=   "chkPersonnel(67)"
         Tab(3).Control(8)=   "chkPersonnel(68)"
         Tab(3).Control(9)=   "chkPersonnel(69)"
         Tab(3).Control(10)=   "chkPersonnel(70)"
         Tab(3).Control(11)=   "chkPersonnel(71)"
         Tab(3).Control(12)=   "chkPersonnel(72)"
         Tab(3).Control(13)=   "chkPersonnel(73)"
         Tab(3).Control(14)=   "chkPersonnel(74)"
         Tab(3).Control(15)=   "chkPersonnel(75)"
         Tab(3).Control(16)=   "chkPersonnel(76)"
         Tab(3).Control(17)=   "chkPersonnel(77)"
         Tab(3).Control(18)=   "chkPersonnel(78)"
         Tab(3).Control(19)=   "chkPersonnel(79)"
         Tab(3).ControlCount=   20
         TabCaption(4)   =   "81-100"
         TabPicture(4)   =   "frmSend.frx":1290
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "chkPersonnel(80)"
         Tab(4).Control(1)=   "chkPersonnel(81)"
         Tab(4).Control(2)=   "chkPersonnel(82)"
         Tab(4).Control(3)=   "chkPersonnel(83)"
         Tab(4).Control(4)=   "chkPersonnel(84)"
         Tab(4).Control(5)=   "chkPersonnel(85)"
         Tab(4).Control(6)=   "chkPersonnel(86)"
         Tab(4).Control(7)=   "chkPersonnel(87)"
         Tab(4).Control(8)=   "chkPersonnel(88)"
         Tab(4).Control(9)=   "chkPersonnel(89)"
         Tab(4).Control(10)=   "chkPersonnel(90)"
         Tab(4).Control(11)=   "chkPersonnel(91)"
         Tab(4).Control(12)=   "chkPersonnel(92)"
         Tab(4).Control(13)=   "chkPersonnel(93)"
         Tab(4).Control(14)=   "chkPersonnel(94)"
         Tab(4).Control(15)=   "chkPersonnel(95)"
         Tab(4).Control(16)=   "chkPersonnel(96)"
         Tab(4).Control(17)=   "chkPersonnel(97)"
         Tab(4).Control(18)=   "chkPersonnel(98)"
         Tab(4).Control(19)=   "chkPersonnel(99)"
         Tab(4).ControlCount=   20
         TabCaption(5)   =   "101-120"
         TabPicture(5)   =   "frmSend.frx":12AC
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "chkPersonnel(100)"
         Tab(5).Control(1)=   "chkPersonnel(101)"
         Tab(5).Control(2)=   "chkPersonnel(102)"
         Tab(5).Control(3)=   "chkPersonnel(103)"
         Tab(5).Control(4)=   "chkPersonnel(104)"
         Tab(5).Control(5)=   "chkPersonnel(105)"
         Tab(5).Control(6)=   "chkPersonnel(106)"
         Tab(5).Control(7)=   "chkPersonnel(107)"
         Tab(5).Control(8)=   "chkPersonnel(108)"
         Tab(5).Control(9)=   "chkPersonnel(109)"
         Tab(5).Control(10)=   "chkPersonnel(110)"
         Tab(5).Control(11)=   "chkPersonnel(111)"
         Tab(5).Control(12)=   "chkPersonnel(112)"
         Tab(5).Control(13)=   "chkPersonnel(113)"
         Tab(5).Control(14)=   "chkPersonnel(114)"
         Tab(5).Control(15)=   "chkPersonnel(115)"
         Tab(5).Control(16)=   "chkPersonnel(116)"
         Tab(5).Control(17)=   "chkPersonnel(117)"
         Tab(5).Control(18)=   "chkPersonnel(118)"
         Tab(5).Control(19)=   "chkPersonnel(119)"
         Tab(5).ControlCount=   20
         TabCaption(6)   =   "121-140"
         TabPicture(6)   =   "frmSend.frx":12C8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "chkPersonnel(120)"
         Tab(6).Control(1)=   "chkPersonnel(121)"
         Tab(6).Control(2)=   "chkPersonnel(122)"
         Tab(6).Control(3)=   "chkPersonnel(123)"
         Tab(6).Control(4)=   "chkPersonnel(124)"
         Tab(6).Control(5)=   "chkPersonnel(125)"
         Tab(6).Control(6)=   "chkPersonnel(126)"
         Tab(6).Control(7)=   "chkPersonnel(127)"
         Tab(6).Control(8)=   "chkPersonnel(128)"
         Tab(6).Control(9)=   "chkPersonnel(129)"
         Tab(6).Control(10)=   "chkPersonnel(130)"
         Tab(6).Control(11)=   "chkPersonnel(131)"
         Tab(6).Control(12)=   "chkPersonnel(132)"
         Tab(6).Control(13)=   "chkPersonnel(133)"
         Tab(6).Control(14)=   "chkPersonnel(134)"
         Tab(6).Control(15)=   "chkPersonnel(135)"
         Tab(6).Control(16)=   "chkPersonnel(136)"
         Tab(6).Control(17)=   "chkPersonnel(137)"
         Tab(6).Control(18)=   "chkPersonnel(138)"
         Tab(6).Control(19)=   "chkPersonnel(139)"
         Tab(6).ControlCount=   20
         TabCaption(7)   =   "141-160"
         TabPicture(7)   =   "frmSend.frx":12E4
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "chkPersonnel(140)"
         Tab(7).Control(1)=   "chkPersonnel(141)"
         Tab(7).Control(2)=   "chkPersonnel(142)"
         Tab(7).Control(3)=   "chkPersonnel(143)"
         Tab(7).Control(4)=   "chkPersonnel(144)"
         Tab(7).Control(5)=   "chkPersonnel(145)"
         Tab(7).Control(6)=   "chkPersonnel(146)"
         Tab(7).Control(7)=   "chkPersonnel(147)"
         Tab(7).Control(8)=   "chkPersonnel(148)"
         Tab(7).Control(9)=   "chkPersonnel(149)"
         Tab(7).Control(10)=   "chkPersonnel(150)"
         Tab(7).Control(11)=   "chkPersonnel(151)"
         Tab(7).Control(12)=   "chkPersonnel(152)"
         Tab(7).Control(13)=   "chkPersonnel(153)"
         Tab(7).Control(14)=   "chkPersonnel(154)"
         Tab(7).Control(15)=   "chkPersonnel(155)"
         Tab(7).Control(16)=   "chkPersonnel(156)"
         Tab(7).Control(17)=   "chkPersonnel(157)"
         Tab(7).Control(18)=   "chkPersonnel(158)"
         Tab(7).Control(19)=   "chkPersonnel(159)"
         Tab(7).ControlCount=   20
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   159
            Left            =   -74880
            TabIndex        =   207
            Top             =   7200
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   158
            Left            =   -74880
            TabIndex        =   206
            Top             =   6840
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   157
            Left            =   -74880
            TabIndex        =   205
            Top             =   6480
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   156
            Left            =   -74880
            TabIndex        =   204
            Top             =   6120
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   155
            Left            =   -74880
            TabIndex        =   203
            Top             =   5760
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   154
            Left            =   -74880
            TabIndex        =   202
            Top             =   5400
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   153
            Left            =   -74880
            TabIndex        =   201
            Top             =   5040
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   152
            Left            =   -74880
            TabIndex        =   200
            Top             =   4680
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   151
            Left            =   -74880
            TabIndex        =   199
            Top             =   4320
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   150
            Left            =   -74880
            TabIndex        =   198
            Top             =   3960
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   149
            Left            =   -74880
            TabIndex        =   197
            Top             =   3600
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   148
            Left            =   -74880
            TabIndex        =   196
            Top             =   3240
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   147
            Left            =   -74880
            TabIndex        =   195
            Top             =   2880
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   146
            Left            =   -74880
            TabIndex        =   194
            Top             =   2520
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   145
            Left            =   -74880
            TabIndex        =   193
            Top             =   2160
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   144
            Left            =   -74880
            TabIndex        =   192
            Top             =   1800
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   143
            Left            =   -74880
            TabIndex        =   191
            Top             =   1440
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   142
            Left            =   -74880
            TabIndex        =   190
            Top             =   1080
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   141
            Left            =   -74880
            TabIndex        =   189
            Top             =   720
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   140
            Left            =   -74880
            TabIndex        =   188
            Top             =   360
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   139
            Left            =   -74880
            TabIndex        =   187
            Top             =   7200
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   138
            Left            =   -74880
            TabIndex        =   186
            Top             =   6840
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   137
            Left            =   -74880
            TabIndex        =   185
            Top             =   6480
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   136
            Left            =   -74880
            TabIndex        =   184
            Top             =   6120
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   135
            Left            =   -74880
            TabIndex        =   183
            Top             =   5760
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   134
            Left            =   -74880
            TabIndex        =   182
            Top             =   5400
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   133
            Left            =   -74880
            TabIndex        =   181
            Top             =   5040
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   132
            Left            =   -74880
            TabIndex        =   180
            Top             =   4680
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   131
            Left            =   -74880
            TabIndex        =   179
            Top             =   4320
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   130
            Left            =   -74880
            TabIndex        =   178
            Top             =   3960
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   129
            Left            =   -74880
            TabIndex        =   177
            Top             =   3600
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   128
            Left            =   -74880
            TabIndex        =   176
            Top             =   3240
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   127
            Left            =   -74880
            TabIndex        =   175
            Top             =   2880
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   126
            Left            =   -74880
            TabIndex        =   174
            Top             =   2520
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   125
            Left            =   -74880
            TabIndex        =   173
            Top             =   2160
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   124
            Left            =   -74880
            TabIndex        =   172
            Top             =   1800
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   123
            Left            =   -74880
            TabIndex        =   171
            Top             =   1440
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   122
            Left            =   -74880
            TabIndex        =   170
            Top             =   1080
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   121
            Left            =   -74880
            TabIndex        =   169
            Top             =   720
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   120
            Left            =   -74880
            TabIndex        =   168
            Top             =   360
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   119
            Left            =   -74880
            TabIndex        =   167
            Top             =   7200
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   118
            Left            =   -74880
            TabIndex        =   166
            Top             =   6840
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   117
            Left            =   -74880
            TabIndex        =   165
            Top             =   6480
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   116
            Left            =   -74880
            TabIndex        =   164
            Top             =   6120
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   115
            Left            =   -74880
            TabIndex        =   163
            Top             =   5760
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   114
            Left            =   -74880
            TabIndex        =   162
            Top             =   5400
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   113
            Left            =   -74880
            TabIndex        =   161
            Top             =   5040
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   112
            Left            =   -74880
            TabIndex        =   160
            Top             =   4680
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   111
            Left            =   -74880
            TabIndex        =   159
            Top             =   4320
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   110
            Left            =   -74880
            TabIndex        =   158
            Top             =   3960
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   109
            Left            =   -74880
            TabIndex        =   157
            Top             =   3600
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   108
            Left            =   -74880
            TabIndex        =   156
            Top             =   3240
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   107
            Left            =   -74880
            TabIndex        =   155
            Top             =   2880
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   106
            Left            =   -74880
            TabIndex        =   154
            Top             =   2520
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   105
            Left            =   -74880
            TabIndex        =   153
            Top             =   2160
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   104
            Left            =   -74880
            TabIndex        =   152
            Top             =   1800
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   103
            Left            =   -74880
            TabIndex        =   151
            Top             =   1440
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   102
            Left            =   -74880
            TabIndex        =   150
            Top             =   1080
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   101
            Left            =   -74880
            TabIndex        =   149
            Top             =   720
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   100
            Left            =   -74880
            TabIndex        =   148
            Top             =   360
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   99
            Left            =   -74880
            TabIndex        =   147
            Top             =   7200
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   98
            Left            =   -74880
            TabIndex        =   146
            Top             =   6840
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   97
            Left            =   -74880
            TabIndex        =   145
            Top             =   6480
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   96
            Left            =   -74880
            TabIndex        =   144
            Top             =   6120
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   95
            Left            =   -74880
            TabIndex        =   143
            Top             =   5760
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   94
            Left            =   -74880
            TabIndex        =   142
            Top             =   5400
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   93
            Left            =   -74880
            TabIndex        =   141
            Top             =   5040
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   92
            Left            =   -74880
            TabIndex        =   140
            Top             =   4680
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   91
            Left            =   -74880
            TabIndex        =   139
            Top             =   4320
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   90
            Left            =   -74880
            TabIndex        =   138
            Top             =   3960
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   89
            Left            =   -74880
            TabIndex        =   137
            Top             =   3600
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   88
            Left            =   -74880
            TabIndex        =   136
            Top             =   3240
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   87
            Left            =   -74880
            TabIndex        =   135
            Top             =   2880
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   86
            Left            =   -74880
            TabIndex        =   134
            Top             =   2520
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   85
            Left            =   -74880
            TabIndex        =   133
            Top             =   2160
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   84
            Left            =   -74880
            TabIndex        =   132
            Top             =   1800
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   83
            Left            =   -74880
            TabIndex        =   131
            Top             =   1440
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   82
            Left            =   -74880
            TabIndex        =   130
            Top             =   1080
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   81
            Left            =   -74880
            TabIndex        =   129
            Top             =   720
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   80
            Left            =   -74880
            TabIndex        =   128
            Top             =   360
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   79
            Left            =   -74880
            TabIndex        =   127
            Top             =   7200
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   78
            Left            =   -74880
            TabIndex        =   126
            Top             =   6840
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   77
            Left            =   -74880
            TabIndex        =   125
            Top             =   6480
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   76
            Left            =   -74880
            TabIndex        =   124
            Top             =   6120
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   75
            Left            =   -74880
            TabIndex        =   123
            Top             =   5760
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   74
            Left            =   -74880
            TabIndex        =   122
            Top             =   5400
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   73
            Left            =   -74880
            TabIndex        =   121
            Top             =   5040
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   72
            Left            =   -74880
            TabIndex        =   120
            Top             =   4680
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   71
            Left            =   -74880
            TabIndex        =   119
            Top             =   4320
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   70
            Left            =   -74880
            TabIndex        =   118
            Top             =   3960
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   69
            Left            =   -74880
            TabIndex        =   117
            Top             =   3600
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   68
            Left            =   -74880
            TabIndex        =   116
            Top             =   3240
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   67
            Left            =   -74880
            TabIndex        =   115
            Top             =   2880
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   66
            Left            =   -74880
            TabIndex        =   114
            Top             =   2520
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   65
            Left            =   -74880
            TabIndex        =   113
            Top             =   2160
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   64
            Left            =   -74880
            TabIndex        =   112
            Top             =   1800
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   63
            Left            =   -74880
            TabIndex        =   111
            Top             =   1440
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   62
            Left            =   -74880
            TabIndex        =   110
            Top             =   1080
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   61
            Left            =   -74880
            TabIndex        =   109
            Top             =   720
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   60
            Left            =   -74880
            TabIndex        =   108
            Top             =   360
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   59
            Left            =   -74880
            TabIndex        =   107
            Top             =   7200
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   58
            Left            =   -74880
            TabIndex        =   106
            Top             =   6840
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   57
            Left            =   -74880
            TabIndex        =   105
            Top             =   6480
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   56
            Left            =   -74880
            TabIndex        =   104
            Top             =   6120
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   55
            Left            =   -74880
            TabIndex        =   103
            Top             =   5760
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   54
            Left            =   -74880
            TabIndex        =   102
            Top             =   5400
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   53
            Left            =   -74880
            TabIndex        =   101
            Top             =   5040
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   52
            Left            =   -74880
            TabIndex        =   100
            Top             =   4680
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   51
            Left            =   -74880
            TabIndex        =   99
            Top             =   4320
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   50
            Left            =   -74880
            TabIndex        =   98
            Top             =   3960
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   49
            Left            =   -74880
            TabIndex        =   97
            Top             =   3600
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   48
            Left            =   -74880
            TabIndex        =   96
            Top             =   3240
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   47
            Left            =   -74880
            TabIndex        =   95
            Top             =   2880
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   46
            Left            =   -74880
            TabIndex        =   94
            Top             =   2520
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   45
            Left            =   -74880
            TabIndex        =   93
            Top             =   2160
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   44
            Left            =   -74880
            TabIndex        =   92
            Top             =   1800
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   43
            Left            =   -74880
            TabIndex        =   91
            Top             =   1440
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   42
            Left            =   -74880
            TabIndex        =   90
            Top             =   1080
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   41
            Left            =   -74880
            TabIndex        =   89
            Top             =   720
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   40
            Left            =   -74880
            TabIndex        =   88
            Top             =   360
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   39
            Left            =   -74880
            TabIndex        =   87
            Top             =   7200
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   38
            Left            =   -74880
            TabIndex        =   86
            Top             =   6840
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   37
            Left            =   -74880
            TabIndex        =   85
            Top             =   6480
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   36
            Left            =   -74880
            TabIndex        =   84
            Top             =   6120
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   35
            Left            =   -74880
            TabIndex        =   83
            Top             =   5760
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   34
            Left            =   -74880
            TabIndex        =   82
            Top             =   5400
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   33
            Left            =   -74880
            TabIndex        =   81
            Top             =   5040
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   32
            Left            =   -74880
            TabIndex        =   80
            Top             =   4680
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   31
            Left            =   -74880
            TabIndex        =   79
            Top             =   4320
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   30
            Left            =   -74880
            TabIndex        =   78
            Top             =   3960
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   29
            Left            =   -74880
            TabIndex        =   77
            Top             =   3600
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   28
            Left            =   -74880
            TabIndex        =   76
            Top             =   3240
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   27
            Left            =   -74880
            TabIndex        =   75
            Top             =   2880
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   26
            Left            =   -74880
            TabIndex        =   74
            Top             =   2520
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   25
            Left            =   -74880
            TabIndex        =   73
            Top             =   2160
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   24
            Left            =   -74880
            TabIndex        =   72
            Top             =   1800
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   23
            Left            =   -74880
            TabIndex        =   71
            Top             =   1440
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   22
            Left            =   -74880
            TabIndex        =   70
            Top             =   1080
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   21
            Left            =   -74880
            TabIndex        =   69
            Top             =   720
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   20
            Left            =   -74880
            TabIndex        =   68
            Top             =   360
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   62
            Top             =   3960
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   61
            Top             =   4320
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   60
            Top             =   2160
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   59
            Top             =   3600
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   58
            Top             =   3240
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   57
            Top             =   2880
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   56
            Top             =   2520
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   55
            Top             =   1800
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   54
            Top             =   1440
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   53
            Top             =   1080
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   720
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   50
            Top             =   4680
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   49
            Top             =   5040
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   48
            Top             =   5400
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   47
            Top             =   5760
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   46
            Top             =   6120
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   45
            Top             =   6480
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   44
            Top             =   6840
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox chkPersonnel 
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   43
            Top             =   7200
            Visible         =   0   'False
            Width           =   6135
         End
      End
   End
   Begin RichTextLib.RichTextBox rtbPreview 
      Height          =   7695
      Left            =   6840
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Preview of file"
      Top             =   240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   13573
      _Version        =   327681
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmSend.frx":1300
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSecurity 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   13800
      TabIndex        =   41
      Top             =   9840
      Width           =   1335
   End
   Begin VB.Label lblSecuritylevel 
      Alignment       =   2  'Center
      Caption         =   "Program security level"
      Height          =   375
      Left            =   13800
      TabIndex        =   40
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label lblPreview 
      Alignment       =   2  'Center
      Caption         =   "File preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   28
      Top             =   0
      Width           =   8385
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkGroup_Click(Index As Integer)
    Rem this code is used to ensure that either an individual or a group
    Rem is selected, as well as a file before enabling the distribute button
    Dim chk As CheckBox
    a = 0
    enable = 0
    For Each CheckBox In chkPersonnel()
        If chkPersonnel(a).Value = 1 And filPath.ListIndex <> -1 Then enable = 1
        a = a + 1
    Next
    Rem enable or disable the distribute button as appropriate
    If enable = 1 Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
    Rem select the individuals if they are in the selected group
    If chkGroup(Index).Value = 1 Then
        If lstGroup(Index).ListCount > 0 Then
            For a = 0 To lstGroup(Index).ListCount - 1
                chkPersonnel(lstGroup(Index).ItemData(a)).Value = 1
            Next a
        End If
    Else
        If lstGroup(Index).ListCount > 0 Then
            For a = 0 To lstGroup(Index).ListCount - 1
                chkPersonnel(lstGroup(Index).ItemData(a)).Value = 0
            Next a
        End If
    End If
End Sub

Private Sub chkGroup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Rem select the appropriate group automatically as the mouse moves over its name
    tabGroups.Tab = Index
End Sub

Private Sub chkGroup_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Rem this allows the user to change the name of the clicked group
    Rem with a right mouse button click
    If Button = vbRightButton Then
        response = InputBox("Enter new name", "New group name", chkGroup(Index).Caption)
        If response <> "" Then chkGroup(Index).Caption = response
    End If
End Sub

Private Sub chkPersonnel_Click(Index As Integer)
    Rem this code is used to ensure that either an individual or a group
    Rem is selected, as well as a file before enabling the distribute button
    Dim chk As CheckBox
    a = 0
    For Each chk In chkPersonnel()
        If chkPersonnel(a).Value = 1 And filPath.ListIndex <> -1 Then enable = 1
        a = a + 1
    Next
    a = 0
    For Each chk In chkGroup()
        If chkGroup(a).Value = 1 And filPath.ListIndex <> -1 Then enable = 1
        a = a + 1
    Next
    If enable = 1 Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub

Private Sub cmdAbout_Click()
    Rem show about window
    frmAbout.Show vbModal
End Sub

Private Sub cmdAll_Click()
    Rem select all individual chkboxes, but only if it has a name
    For a = chkPersonnel().LBound To chkPersonnel().UBound
        If chkPersonnel(a).Caption <> "" Then chkPersonnel(a).Value = 1
    Next a
End Sub

Private Sub cmdBrowse_Click()
    Rem change file directory to a new user selected one
    response = InputBox("Enter new directory", "Change directory", filPath.Path)
    If response <> "" Then
        filPath.Path = response
        fmeFiles.ToolTipText = filPath.Path
        filPath.Refresh
    End If
End Sub

Private Sub cmdClear_Click()
    Rem remove file preview
    rtbPreview.filename = ""
    filPath.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
    If filPath.ListIndex <> -1 Then
        Rem ask if it is ok to delete the selected file
        response = MsgBox("Delete file: " & filPath.List(filPath.ListIndex) & " ?", vbOKCancel)
        If response = vbOK Then
            Rem delete selected file
            Kill (filPath.Path & "\" & filPath.List(filPath.ListIndex))
            txtLog = Date & " " & Time & " Deleted " & (filPath.Path & "\" & filPath.List(filPath.ListIndex))
            filPath.Refresh
            rtbPreview.filename = ""
        End If
    Else
        Rem no file selected
        MsgBox "Please select a file to delete"
    End If
End Sub

Private Sub cmdEnd_Click()
    Rem end program?
    response = MsgBox("End Router?", vbQuestion + vbYesNo + vbDefaultButton2, "End")
    If response = vbYes Then
        End
    End If
End Sub

Private Sub cmdHelp_Click()
    Rem show helpful file in preview window
    rtbPreview.filename = "c:\welcome.txt"
    rtbPreview.Refresh
End Sub

Private Sub cmdModify_Click()
    Rem enable frame around group info
    fmeGroups.Enabled = True
End Sub

Private Sub cmdNone_Click()
    Rem clear all selected individuals
    For a = chkPersonnel().LBound To chkPersonnel().UBound
        chkPersonnel(a).Value = 0
    Next a
End Sub

Private Sub cmdRename_Click()
    If filPath.ListIndex <> -1 Then
        Rem rename the selected file
        response = InputBox("Rename " & filPath.List(filPath.ListIndex) & " to: ", "Rename file", filPath.List(filPath.ListIndex))
        If response <> "" Then
            Name filPath.Path & "\" & filPath.List(filPath.ListIndex) As filPath.Path & "\" & response
            txtLog = Date & " " & Time & " Renamed " & filPath.Path & "\" & filPath.List(filPath.ListIndex) & " to " & filPath.Path & "\" & response
            filPath.Refresh
            rtbPreview.filename = ""
        End If
    Else
        Rem no file selected
        MsgBox "Please select a file to rename"
    End If
End Sub

Private Sub cmdSave_Click()
    Rem delete groups file
    Kill gGroups
    Open gGroups For Output As #1
        Rem group names
        For a = 1 To 10
            Write #1, a
            Print #1, chkGroup(a - 1).Caption
        Next a
        Rem group members
        For a = lstGroup().LBound To lstGroup().UBound
            Output = a + 1 & "," & lstGroup(a).ListCount
            For b = 0 To lstGroup(a).ListCount - 1
                 Output = Output & "," & lstGroup(a).ItemData(b)
            Next b
            Print #1, Output
        Next a
    Close 1
End Sub

Private Sub cmdSend_Click()
    Dim no_individuals, no_groups, a As Integer
    Rem do the following only if a file has been selected
    If filPath.ListIndex > -1 Then
        Rem check if any individuals have been chosen
        For a = chkPersonnel().LBound To chkPersonnel().UBound
            If chkPersonnel(a).Value = 1 Then no_individuals = no_individuals + 1
        Next a
        Rem check if any groups have been chosen
        For a = chkGroup().LBound To chkGroup().UBound
            If chkGroup(a).Value = 1 Then no_groups = no_groups + 1
        Next a
        Rem one or more individuals have been chosen
        If no_individuals <> 0 Then
            For a = chkPersonnel().LBound To chkPersonnel().UBound
                Rem distribute file to individuals
                If chkPersonnel(a).Value = 1 Then
                    Rem check if file already exists in destination directory
                    If Dir(gDir(a) & "\" & filPath.List(filPath.ListIndex)) <> filPath.List(filPath.ListIndex) Then
                        FileCopy filPath.Path & "\" & filPath.List(filPath.ListIndex), gDir(a) & "\" & filPath.List(filPath.ListIndex)
                        txtLog = Date & " " & Time & " Distributed " & filPath.Path & "\" & filPath.List(filPath.ListIndex) & " to " & gDir(a) & "\" & filPath.List(filPath.ListIndex)
                    Else
                        Rem inform user that the file already exists
                        MsgBox filPath.List(filPath.ListIndex) & " already exists in " & gDir(a)
                    End If
                End If
                Rem clear selected individual
                chkPersonnel(a).Value = 0
            Next a
        End If
        Rem clear selected groups
        For a = chkGroup().LBound To chkGroup().UBound
            chkGroup(a).Value = 0
        Next a
        Rem archive selected file
        If no_individuals <> 0 Or no_groups <> 0 Then
            rtbPreview.filename = ""
            If chkArchive = 1 Then
                If Dir(gArchive & "\" & filPath.List(filPath.ListIndex)) <> filPath.List(filPath.ListIndex) Then
                    FileCopy filPath.Path & "\" & filPath.List(filPath.ListIndex), gArchive & "\" & filPath.List(filPath.ListIndex)
                Else
                    Rem file already exists in archive directory
                End If
            End If
            filPath.ListIndex = -1
        End If
        cmdSend.Enabled = False
    End If
End Sub

Private Sub cmdTransfer_Click(Index As Integer)
    Select Case Index
        Rem transfer to group
        Case 0
            If lstIndividuals.ListIndex > -1 Then
                For a = 0 To lstGroup(tabGroups.Tab).ListCount
                    Rem test if the selected individual is already in the selected group
                    If lstGroup(tabGroups.Tab).List(a) = lstIndividuals.List(lstIndividuals.ListIndex) Then
                        present = 1
                    End If
                Next a
                Rem if the individual was not already present, then add them
                Rem to the current group
                If present = 0 Then
                    Rem add individual
                    lstGroup(tabGroups.Tab).AddItem lstIndividuals.List(lstIndividuals.ListIndex)
                    Rem give the individual in the group a number representing
                    Rem its original position in the individuals list
                    lstGroup(tabGroups.Tab).ItemData(lstGroup(tabGroups.Tab).ListCount - 1) = lstIndividuals.ListIndex
                    Rem MsgBox lstIndividuals.ListIndex
                End If
            End If
        Rem remove selected individual from selected group
        Case 1
            If lstGroup(tabGroups.Tab).ListIndex > -1 Then
                lstGroup(tabGroups.Tab).RemoveItem (lstGroup(tabGroups.Tab).ListIndex)
            End If
    End Select
End Sub

Private Sub cmdTreeview_Click()
    Rem display the view existing files form
    frmTreeview.Show vbModal
End Sub

Private Sub filPath_Click()
    Rem make the preview box filename the selected file
    Rem first clear the preview screen
    rtbPreview.filename = ""
    If filPath.ListIndex > -1 Then
        Open filPath.Path & "\" & filPath.List(filPath.ListIndex) For Input As #3
        Rem test selected file for security classification
        frmWait.Show
        Rem reset classifications
        gClassification = "Unknown"
        gSecurity = 0
        For a = 1 To gNo_Of_Classes
            gClass(a) = 0
        Next a
        a = 0
        Do Until EOF(3)
            a = a + 1
            Line Input #3, current_line
            current_line = UCase$(current_line)
            If current_line <> "" Then
                Rem get title
                If InStr(1, current_line, "SUBJ") > 0 Then
                    subject = Right$(current_line, Len(current_line) - 6)
                End If
                If InStr(1, current_line, "UNCLAS") > 0 Then
                    gClass(1) = 1
                    Rem Exit Do
                End If
                If InStr(1, current_line, "R E S T R I C T E D") > 0 Then
                    gClass(2) = 1
                    Rem Exit Do
                End If
                If InStr(1, current_line, "C O N F I D E N T I A L") > 0 Then
                    gClass(3) = 1
                    Rem Exit Do
                End If
                If InStr(1, current_line, "S E C R E T") > 0 Then
                    gClass(4) = 1
                    Rem Exit Do
                End If
                If InStr(1, current_line, "T O P S E C R E T") > 0 Then
                    gClass(5) = 1
                    Rem Exit Do
                End If
                If InStr(1, current_line, "IN CONFIDENCE") > 0 Then
                    gClass(6) = 1
                    Rem Exit Do
                End If
            End If
            DoEvents
            Rem search only the first 100 lines of the file
            If a >= 100 Then
                Exit Do
            End If
        Loop
        Close 3
        Rem six levels of security
        For a = 1 To gNo_Of_Classes
            If gClass(a) = 1 Then
                gClassification = gType(a)
                gSecurity = a
            End If
        Next a
        Unload frmWait
        frmSend.Show
        If txtSecurity < gSecurity Then
            msg = "Subject: " & subject & vbCrLf
            msg = msg & "Message classification: " & gClassification & vbCrLf
            msg = msg & "This signal is classified higher than that allowed for viewing for this program:" & vbCrLf
            msg = msg & gType(txtSecurity)
            MsgBox msg, vbCritical + vbOKOnly
        Else
            If Right$(filPath.Path, 1) <> "\" Then
                rtbPreview.filename = filPath.Path & "\" & filPath.List(filPath.ListIndex)
            Else
                rtbPreview.filename = filPath.Path & filPath.List(filPath.ListIndex)
            End If
        End If
    End If
    Dim chk As CheckBox
    a = 0
    Rem cycle through individual check boxes, checking if at least one is selected
    For Each chk In chkPersonnel()
        If chkPersonnel(a).Value = 1 And filPath.ListIndex <> -1 Then enable = 1
        a = a + 1
    Next
    a = 0
    Rem cycle through group check boxes, checking if at least one is selected
    For Each chk In chkGroup()
        If chkGroup(a).Value = 1 And filPath.ListIndex <> -1 Then enable = 1
        a = a + 1
    Next
    Rem if either a group or an individual is selected, then enable the distribute button
    Rem but if neither a group or individual is selected, then disable the button
    If enable = 1 Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub

Private Sub filPath_PathChange()
    Rem clear the preview
    rtbPreview.filename = ""
End Sub

Private Sub Form_Activate()
    Rem set the group names from the ones found in the text file
    For a = 0 To 9
        chkGroup(a).Caption = gGroup(a)
        chkGroup(a).Refresh
    Next a
    Rem the if statement prevents the Extension text box from being reset after
    Rem every Message box
    If txtExtension = "" Then txtExtension = gExtension
    Rem display program security level
    lblSecurity = gType(txtSecurity)
End Sub

Private Sub Form_Load()
    Rem set up form stuff
    Dim a As Integer
    filPath.Pattern = "." & gExtension
    fmeFiles.ToolTipText = filPath.Path
    filPath.Refresh
End Sub

Private Sub txtExtension_Change()
    Rem update the files listbox after a modification to the extension textbox
    filPath.Pattern = "*." & UCase$(txtExtension)
    filPath.Refresh
    rtbPreview.filename = ""
End Sub

Private Sub txtExtension_Click()
    txtExtension.SelStart = 0
    txtExtension.SelLength = Len(txtExtension)
End Sub

Private Sub txtLog_Change()
    Rem this writes the state of the txtLog control to a log file
    Open gLog For Append As #2
        Print #2, txtLog
    Close 2
End Sub
