VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   Caption         =   "Music Database (c) 2000 by D. Robinson"
   ClientHeight    =   7980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7530
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlTree 
      Left            =   6600
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":145A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":220A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoAlbums 
      Height          =   330
      Left            =   2760
      Top             =   6120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "admin"
      Password        =   ""
      RecordSource    =   "select * from albums"
      Caption         =   "adoAlbums"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoSongs 
      Height          =   330
      Left            =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Songs"
      Caption         =   "adoSongs"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   6120
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   1404
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7530
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   7530
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Song information"
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   4
         Tag             =   " ListView:"
         Top             =   0
         Width           =   3330
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Artist | Album | Song"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   " TreeView:"
         Top             =   0
         Width           =   2655
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Large Icons"
            Object.ToolTipText     =   "View Large Icons"
            ImageKey        =   "View Large Icons"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Small Icons"
            Object.ToolTipText     =   "View Small Icons"
            ImageKey        =   "View Small Icons"
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View List"
            Object.ToolTipText     =   "View List"
            ImageKey        =   "View List"
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Details"
            Object.ToolTipText     =   "View Details"
            ImageKey        =   "View Details"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7710
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7646
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "27/02/00"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "8:53 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   6600
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6600
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":265E
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2770
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2882
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2994
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AA6
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BB8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CCA
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DDC
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EEE
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3000
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3112
            Key             =   "View Details"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2760
      TabIndex        =   5
      Top             =   720
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   8467
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvData 
      Height          =   4800
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8467
      _Version        =   393217
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imlTree"
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc adoArtists 
      Height          =   330
      Left            =   2760
      Top             =   5640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Artists"
      Caption         =   "adoArtists"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInfo 
      Height          =   330
      Left            =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmMain.frx":3224
      OLEDBString     =   $"frmMain.frx":32AD
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   ""
      RecordSource    =   "select * from Info;"
      Caption         =   "adoInfo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2640
      MousePointer    =   9  'Size W E
      Top             =   720
      Width           =   135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "Sen&d to"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Rena&me"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Paste &Special..."
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "&Invert Selection"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Lar&ge Icons"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "S&mall Icons"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&Details"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Arrange &Icons"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "&Web Browser"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Data"
      Begin VB.Menu mnuDataData 
         Caption         =   "Enter Data"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  
Dim mbMoving As Boolean
Const sglSplitLimit = 500

Private Sub Form_Activate()
    mnuRefresh_Click
End Sub

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6800)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6600)
End Sub


Private Sub Form_Paint()
    lvListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    Select Case lvListView.View
        Case lvwIcon
            tbToolBar.Buttons(LISTVIEW_MODE0).Value = tbrPressed
        Case lvwSmallIcon
            tbToolBar.Buttons(LISTVIEW_MODE1).Value = tbrPressed
        Case lvwList
            tbToolBar.Buttons(LISTVIEW_MODE2).Value = tbrPressed
        Case lvwReport
            tbToolBar.Buttons(LISTVIEW_MODE3).Value = tbrPressed
    End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
End Sub

Private Sub mnuDatafrmArtists_Click()
    Dim f As New frmArtists
    f.Show
End Sub

Private Sub mnuDatafrmAlbums_Click()
    Dim f As New frmAlbums
    f.Show
End Sub

Private Sub mnuDataData_Click()
    Dim f As New Data
    f.Show vbModal
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub TreeView1_DragDrop(Source As Control, x As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls x
    End If
End Sub

Sub SizeControls(x As Single)
    On Error Resume Next
    
    'set the width
    If x < 1500 Then x = 1500
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    tvData.Width = x
    imgSplitter.Left = x
    lvListView.Left = x + 40
    lvListView.Width = Me.Width - (tvData.Width + 140)
    lblTitle(0).Width = tvData.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40

    'set the top
  
    If tbToolBar.Visible Then
        tvData.Top = tbToolBar.Height + picTitles.Height
    Else
        tvData.Top = picTitles.Height
    End If

  lvListView.Top = tvData.Top
    
    'set the height
    If sbStatusBar.Visible Then
        tvData.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvData.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    
    ' set the left side

    lvListView.Height = tvData.Height
    imgSplitter.Top = tvData.Top
    imgSplitter.Height = tvData.Height
End Sub

Private Sub mnuRefresh_Click()
    tvData.Nodes.Clear
    Dim x As Node
    adoArtists.Refresh
    adoArtists.Recordset.MoveFirst
    While Not adoArtists.Recordset.EOF
        Set x = tvData.Nodes.Add(, , "Ar" & adoArtists.Recordset.Fields("Name"), adoArtists.Recordset.Fields("Name"), 8)
        adoArtists.Recordset.MoveNext
    Wend
    
    adoAlbums.Refresh
    adoAlbums.Recordset.MoveFirst
    While Not adoAlbums.Recordset.EOF
        parentID = adoAlbums.Recordset.Fields("Artist ID")
        adoArtists.Recordset.MoveFirst
        adoArtists.Recordset.AbsolutePosition = parentID
        txtParent = adoArtists.Recordset.Fields("Name")
        Set x = tvData.Nodes.Add("Ar" & txtParent, tvwChild, "Al" & adoAlbums.Recordset.Fields("Name"), adoAlbums.Recordset.Fields("Name"), 9)
        adoAlbums.Recordset.MoveNext
    Wend
    
    adoSongs.Refresh
    adoSongs.Recordset.MoveFirst
    While Not adoSongs.Recordset.EOF
        parentID = adoSongs.Recordset.Fields("Album ID")
        adoAlbums.Recordset.MoveFirst
        adoAlbums.Recordset.AbsolutePosition = parentID
        txtParent = adoAlbums.Recordset.Fields("Name")
        Set x = tvData.Nodes.Add("Al" & txtParent, tvwChild, "So" & adoSongs.Recordset.Fields("Song Title"), adoSongs.Recordset.Fields("Song Title"), 2)
        adoSongs.Recordset.MoveNext
    Wend
    
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Back"
            'ToDo: Add 'Back' button code.
            MsgBox "Add 'Back' button code."
        Case "Forward"
            'ToDo: Add 'Forward' button code.
            MsgBox "Add 'Forward' button code."
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Delete"
            mnuFileDelete_Click
        Case "Properties"
            mnuFileProperties_Click
        Case "View Large Icons"
            lvListView.View = lvwIcon
        Case "View Small Icons"
            lvListView.View = lvwSmallIcon
        Case "View List"
            lvListView.View = lvwList
        Case "View Details"
            lvListView.View = lvwReport
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuViewWebBrowser_Click()
    'ToDo: Add 'mnuViewWebBrowser_Click' code.
    MsgBox "Add 'mnuViewWebBrowser_Click' code."
End Sub

Private Sub mnuViewOptions_Click()
    'ToDo: Add 'mnuViewOptions_Click' code.
    MsgBox "Add 'mnuViewOptions_Click' code."
End Sub

Private Sub mnuViewRefresh_Click()
    'ToDo: Add 'mnuViewRefresh_Click' code.
    MsgBox "Add 'mnuViewRefresh_Click' code."
End Sub



Private Sub mnuVAIByDate_Click()
    'ToDo: Add 'mnuVAIByDate_Click' code.
'  lvListView.SortKey = DATE_COLUMN
End Sub


Private Sub mnuVAIByName_Click()
    'ToDo: Add 'mnuVAIByName_Click' code.
'  lvListView.SortKey = NAME_COLUMN
End Sub


Private Sub mnuVAIBySize_Click()
    'ToDo: Add 'mnuVAIBySize_Click' code.
'  lvListView.SortKey = SIZE_COLUMN
End Sub


Private Sub mnuVAIByType_Click()
    'ToDo: Add 'mnuVAIByType_Click' code.
'  lvListView.SortKey = TYPE_COLUMN
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuEditInvertSelection_Click()
    'ToDo: Add 'mnuEditInvertSelection_Click' code.
    MsgBox "Add 'mnuEditInvertSelection_Click' code."
End Sub

Private Sub mnuEditSelectAll_Click()
    'ToDo: Add 'mnuEditSelectAll_Click' code.
    MsgBox "Add 'mnuEditSelectAll_Click' code."
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    'ToDo: Add 'mnuEditPaste_Click' code.
    MsgBox "Add 'mnuEditPaste_Click' code."
End Sub

Private Sub mnuEditCopy_Click()
    'ToDo: Add 'mnuEditCopy_Click' code.
    MsgBox "Add 'mnuEditCopy_Click' code."
End Sub

Private Sub mnuEditCut_Click()
    'ToDo: Add 'mnuEditCut_Click' code.
    MsgBox "Add 'mnuEditCut_Click' code."
End Sub

Private Sub mnuEditUndo_Click()
    'ToDo: Add 'mnuEditUndo_Click' code.
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub

Private Sub mnuFileClose_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileRename_Click()
    'ToDo: Add 'mnuFileRename_Click' code.
    MsgBox "Add 'mnuFileRename_Click' code."
End Sub

Private Sub mnuFileDelete_Click()
    'ToDo: Add 'mnuFileDelete_Click' code.
    MsgBox "Add 'mnuFileDelete_Click' code."
End Sub

Private Sub mnuFileNew_Click()
    'ToDo: Add 'mnuFileNew_Click' code.
    MsgBox "Add 'mnuFileNew_Click' code."
End Sub

Private Sub mnuFileSendTo_Click()
    'ToDo: Add 'mnuFileSendTo_Click' code.
    MsgBox "Add 'mnuFileSendTo_Click' code."
End Sub

Private Sub mnuFileFind_Click()
    'ToDo: Add 'mnuFileFind_Click' code.
    MsgBox "Add 'mnuFileFind_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ToDo: add code to process the opened file

End Sub

Private Sub tvData_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim nAlbum As Node
    Dim nArtist As Node
    If Left(Node.Key, 2) = "So" Then ' Check if selected node is at Song Level
        Set nAlbum = Node.Parent
        Set nArtist = nAlbum.Parent
        txtSQL = "SELECT * FROM Info WHERE [Song Title] = " & Chr(39) & Node.Text & Chr(39) & " AND Albums_Name = " & Chr(39) & nAlbum.Text & Chr(39) & " AND Artists_Name = " & Chr(39) & nArtist.Text & Chr(39)
        Debug.Print txtSQL
        adoInfo.RecordSource = txtSQL
        adoInfo.Refresh
        era = adoInfo.Recordset.Fields("Era Name")
        MsgBox era
    End If
End Sub
