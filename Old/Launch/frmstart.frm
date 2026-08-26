VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStart 
   BorderStyle     =   0  'None
   Caption         =   "Launcher"
   ClientHeight    =   2730
   ClientLeft      =   8985
   ClientTop       =   10305
   ClientWidth     =   7980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   7980
   Begin VB.Timer timTZUpdate 
      Interval        =   30000
      Left            =   1080
      Top             =   2040
   End
   Begin VB.CommandButton cmdPrinters 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7400
      TabIndex        =   109
      ToolTipText     =   "Change default printer"
      Top             =   10
      Width           =   200
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   45
      ToolTipText     =   "Display more LAN specific information"
      Top             =   10
      Width           =   200
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7600
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Send Launcher to Tooltray"
      Top             =   10
      Width           =   200
   End
   Begin VB.Timer timClock 
      Interval        =   1000
      Left            =   120
      Top             =   2040
   End
   Begin ComctlLib.StatusBar sbrLauncher 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   46
      Top             =   2475
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10054
            MinWidth        =   10054
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3933
            MinWidth        =   3933
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "End Launcher"
      Top             =   10
      Width           =   200
   End
   Begin VB.PictureBox trayLaunch 
      Height          =   480
      Left            =   8400
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   110
      Top             =   2280
      Width           =   1200
   End
   Begin TabDlg.SSTab tabLaunch 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   3113
      TabCaption(0)   =   "Applications"
      TabPicture(0)   =   "frmStart.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdButton(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdButton(14)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdButton(15)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdButton(16)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdButton(17)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdButton(18)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdButton(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdButton(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdButton(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdButton(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdButton(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdButton(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdButton(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdButton(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdButton(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdButton(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdButton(19)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdButton(10)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdButton(11)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdButton(12)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdButton(20)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdButton(21)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdButton(22)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdButton(23)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdButton(24)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdButton(25)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   " Files"
      TabPicture(1)   =   "frmStart.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdButton(51)"
      Tab(1).Control(1)=   "cmdButton(50)"
      Tab(1).Control(2)=   "cmdButton(49)"
      Tab(1).Control(3)=   "cmdButton(48)"
      Tab(1).Control(4)=   "cmdButton(47)"
      Tab(1).Control(5)=   "cmdButton(46)"
      Tab(1).Control(6)=   "cmdButton(45)"
      Tab(1).Control(7)=   "cmdButton(44)"
      Tab(1).Control(8)=   "cmdButton(43)"
      Tab(1).Control(9)=   "cmdButton(42)"
      Tab(1).Control(10)=   "cmdButton(41)"
      Tab(1).Control(11)=   "cmdButton(40)"
      Tab(1).Control(12)=   "cmdButton(38)"
      Tab(1).Control(13)=   "cmdButton(37)"
      Tab(1).Control(14)=   "cmdButton(36)"
      Tab(1).Control(15)=   "cmdButton(35)"
      Tab(1).Control(16)=   "cmdButton(34)"
      Tab(1).Control(17)=   "cmdButton(33)"
      Tab(1).Control(18)=   "cmdButton(32)"
      Tab(1).Control(19)=   "cmdButton(31)"
      Tab(1).Control(20)=   "cmdButton(30)"
      Tab(1).Control(21)=   "cmdButton(29)"
      Tab(1).Control(22)=   "cmdButton(28)"
      Tab(1).Control(23)=   "cmdButton(27)"
      Tab(1).Control(24)=   "cmdButton(26)"
      Tab(1).Control(25)=   "cmdButton(39)"
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmStart.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdButton(52)"
      Tab(2).Control(1)=   "cmdButton(53)"
      Tab(2).Control(2)=   "cmdButton(54)"
      Tab(2).Control(3)=   "cmdButton(55)"
      Tab(2).Control(4)=   "cmdButton(56)"
      Tab(2).Control(5)=   "cmdButton(57)"
      Tab(2).Control(6)=   "cmdButton(58)"
      Tab(2).Control(7)=   "cmdButton(59)"
      Tab(2).Control(8)=   "cmdButton(60)"
      Tab(2).Control(9)=   "cmdButton(61)"
      Tab(2).Control(10)=   "cmdButton(62)"
      Tab(2).Control(11)=   "cmdButton(63)"
      Tab(2).Control(12)=   "cmdButton(64)"
      Tab(2).Control(13)=   "cmdButton(65)"
      Tab(2).Control(14)=   "cmdButton(66)"
      Tab(2).Control(15)=   "cmdButton(67)"
      Tab(2).Control(16)=   "cmdButton(68)"
      Tab(2).Control(17)=   "cmdButton(69)"
      Tab(2).Control(18)=   "cmdButton(70)"
      Tab(2).Control(19)=   "cmdButton(71)"
      Tab(2).Control(20)=   "cmdButton(72)"
      Tab(2).Control(21)=   "cmdButton(73)"
      Tab(2).Control(22)=   "cmdButton(74)"
      Tab(2).Control(23)=   "cmdButton(75)"
      Tab(2).Control(24)=   "cmdButton(76)"
      Tab(2).Control(25)=   "cmdButton(77)"
      Tab(2).ControlCount=   26
      TabCaption(3)   =   "  Personal"
      TabPicture(3)   =   "frmStart.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdButton(78)"
      Tab(3).Control(1)=   "cmdButton(79)"
      Tab(3).Control(2)=   "cmdButton(80)"
      Tab(3).Control(3)=   "cmdButton(81)"
      Tab(3).Control(4)=   "cmdButton(82)"
      Tab(3).Control(5)=   "cmdButton(83)"
      Tab(3).Control(6)=   "cmdButton(84)"
      Tab(3).Control(7)=   "cmdButton(85)"
      Tab(3).Control(8)=   "cmdButton(86)"
      Tab(3).Control(9)=   "cmdButton(87)"
      Tab(3).Control(10)=   "cmdButton(88)"
      Tab(3).Control(11)=   "cmdButton(89)"
      Tab(3).Control(12)=   "cmdButton(90)"
      Tab(3).Control(13)=   "cmdButton(91)"
      Tab(3).Control(14)=   "cmdButton(92)"
      Tab(3).Control(15)=   "cmdButton(93)"
      Tab(3).Control(16)=   "cmdButton(94)"
      Tab(3).Control(17)=   "cmdButton(95)"
      Tab(3).Control(18)=   "cmdButton(96)"
      Tab(3).Control(19)=   "cmdButton(97)"
      Tab(3).Control(20)=   "cmdButton(98)"
      Tab(3).Control(21)=   "cmdButton(99)"
      Tab(3).Control(22)=   "cmdButton(100)"
      Tab(3).Control(23)=   "cmdButton(101)"
      Tab(3).Control(24)=   "cmdButton(102)"
      Tab(3).Control(25)=   "cmdButton(103)"
      Tab(3).ControlCount=   26
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   103
         Left            =   -67680
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   102
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   101
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   100
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   99
         Left            =   -70080
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   98
         Left            =   -70680
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   97
         Left            =   -71280
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   96
         Left            =   -71880
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   95
         Left            =   -72480
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   94
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   93
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   92
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   91
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   90
         Left            =   -67680
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   89
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   88
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   87
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   86
         Left            =   -70080
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   85
         Left            =   -70680
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   84
         Left            =   -71280
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   83
         Left            =   -71880
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   82
         Left            =   -72480
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   81
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   80
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   79
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   78
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   77
         Left            =   -67680
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   76
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   75
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   74
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   73
         Left            =   -70080
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   72
         Left            =   -70680
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   71
         Left            =   -71280
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   70
         Left            =   -71880
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   69
         Left            =   -72480
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   68
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   67
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   66
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   65
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   64
         Left            =   -67680
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   63
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   62
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   61
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   60
         Left            =   -70080
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   59
         Left            =   -70680
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   58
         Left            =   -71280
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   57
         Left            =   -71880
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   56
         Left            =   -72480
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   55
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   54
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   53
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   52
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   51
         Left            =   -67680
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   50
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   49
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   48
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   47
         Left            =   -70080
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   46
         Left            =   -70680
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   45
         Left            =   -71280
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   44
         Left            =   -71880
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   43
         Left            =   -72480
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   42
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   41
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   40
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   38
         Left            =   -67680
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   37
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   36
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   35
         Left            =   -69480
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   34
         Left            =   -70080
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   33
         Left            =   -70680
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   32
         Left            =   -71280
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   31
         Left            =   -71880
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   30
         Left            =   -72480
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   25
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   24
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   23
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   22
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   21
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   20
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   12
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   11
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   10
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   29
         Left            =   -73080
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   28
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   27
         Left            =   -74280
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   26
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   39
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   19
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   9
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   8
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   7
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   6
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   5
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   4
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   3
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   2
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   18
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   17
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   16
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   15
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   14
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdButton 
         Height          =   615
         Index           =   13
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   2520
      Picture         =   "frmStart.frx":04B2
      Top             =   2040
      Width           =   480
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Wrong
    If Button = vbLeftButton Then
        If gProgram(Index) <> "" Then
            ChDir (gDir(Index))
            Shell gProgram(Index) & " " & gFile(Index), vbNormalFocus
            'frmStart.Hide
            'trayLaunch.InTray = True
        End If
    End If
    If Button = vbRightButton Then
        With frmRightClick
        .Show
        .lblProgram = gProgram(Index)
        .lblFile = gFile(Index)
        .lblDirectory = gDir(Index)
        .lblIcon = gIcon(Index)
        .lblTitle = gTip(Index) & " properties"
        .tmrTime.Interval = 8000
        End With
    End If
    GoTo correct
Wrong:
    msg = "Error " & Err & " (" & Err.Description & ")" & vbCrLf & vbCrLf
    msg = msg & "Try editing your personal program shortcuts (h:\personal.txt)" & vbCrLf
    msg = msg & "and configuring it to run the program directly." & vbCrLf & vbCrLf
    msg = msg & "If this does not work, ensure the program exists" & vbCrLf
    msg = msg & "and you have read and execute rights to it."
    MsgBox msg, vbCritical + vbOKOnly, "Critical Error"
correct:
End Sub

Private Sub cmdExit_Click()
    response = MsgBox("Quit?", vbYesNo + vbQuestion + vbDefaultButton2)
    If response = vbYes Then
        Unload Me
        Shell_NotifyIcon NIM_DELETE, nid
        Set frmStart = Nothing
        End
    End If
End Sub

Private Sub cmdHelp_Click()
    frmStart.Hide
    frmOther.Show
End Sub

Private Sub cmdMinimize_Click()
    frmStart.Hide
    'trayLaunch.InTray = True
End Sub

Private Sub cmdPrinters_Click()
    frmPrinters.Show
    frmStart.Hide
End Sub

Private Sub Form_Activate()
    'Get the local time zone from the server
    UpdateTime
    mytime = Format(Time, "Long Time")
    mytime = DateAdd("h", gTZ, mytime)
    sbrLauncher.Panels(2).Text = Format(mytime, "hh:mm:ss AMPM") & " Local"
    Left = Screen.Width - Width
    Height = 1935 + 255 ' + 255 for status bar
    Top = Screen.Height - Height - 410
    Rem make command buttons with associated programs visible
    On Error Resume Next
    For a = 0 To cmdButton().UBound
        If gProgram(a) <> "" Then cmdButton(a).Visible = True
    Next a
    tabLaunch.SetFocus
    Unload frmEdit
    With nid
      .cbSize = Len(nid)
      .hwnd = Me.hwnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = WM_MOUSEMOVE
      .hIcon = imgIcon.Picture
      .szTip = "Program Launcher" & vbNullChar
    End With
 
  Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'//////////////////////////////////////////////////////////////////
'//Purpose:         This is the callback function of icon in the
'//                 system tray.  This is where we will process
'//                 what the application will do when Mouse Input
'//                 is given to the icon.
'//
'//Inputs:          What Button was clicked (this is button & shift),
'//                 also, the X & Y coordinates of the mouse.
'//////////////////////////////////////////////////////////////////

  Dim msg As Long     '//The callback value
  
  '//The value of X will vary depending
  '//upon the ScaleMode setting.  Here
  '//we are using that fact to determine
  '//what the value of 'msg' should really be
  If (Me.ScaleMode = vbPixels) Then
    msg = X
  Else
    msg = X / Screen.TwipsPerPixelX
  End If

  Select Case msg
    Case WM_LBUTTONDBLCLK    '515 restore form window
      Me.WindowState = vbNormal
      Call SetForegroundWindow(Me.hwnd)
      Me.Show
      
    Case WM_RBUTTONUP        '517 display popup menu
      Call SetForegroundWindow(Me.hwnd)
      'Me.PopupMenu Me.mnuSystray
    
    Case WM_LBUTTONUP        '514 restore form window
      'commonly an application on the
      'systray will do nothing on a
      'single mouse_click, so nothing
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    GetKeyboardState kbArray
    kbArray.kbByte(VK_CAPITAL) = 0
    ' SetKeyboardState kbArray  this line is to be checked...
End Sub

Private Sub Form_Resize()
    Left = Screen.Width - Width
    Height = 1935 + 255 ' + 255 for status bar
    Top = Screen.Height - Height - 400
End Sub

Private Sub timTZUpdate_Timer()
    UpdateTime
End Sub

Private Sub wskUDP_DataArrival(ByVal bytesTotal As Long)
    Dim strIn As String
    wskUDP.GetData strIn, vbString
End Sub

Private Sub tabLaunch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        tabLaunch.Tab = 3
        frmEdit.Show vbModal
    End If
End Sub

Private Sub timClock_Timer()
    mytime = Format(Time, "Long Time")
    mytime = DateAdd("h", gTZ, mytime)
    'MyDate = Format(Date, "Long Date")
    sbrLauncher.Panels(2).Text = Format(mytime, "hh:mm:ss") & " Local"
    'sbrLauncher.Panels(2).Text = Format(MyDate, "")
End Sub

Private Sub trayLaunch_MouseDblClick(Button As Integer, Id As Long)
    frmStart.Show
    'trayLaunch.InTray = False
End Sub

Private Sub OnTop()
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    If SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS) = True Then
        success% = SetWindowPos(frmCPU.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

Sub UpdateTime()
    'Get the local time zone from the server
    On Error Resume Next
    present = Dir(gTZFile)
    On Error GoTo 0
    If present <> "" Then
        Open gTZFile For Input As #2
        Input #2, stuff
        If Left$(stuff, 3) = "TZ=" Then
            gTZ = Right$(stuff, Len(stuff) - 3)
            If InStr(1, stuff, "-", vbTextCompare) > 0 Then gTZ = 0 - gTZ
        End If
        Close 2
    End If
End Sub
