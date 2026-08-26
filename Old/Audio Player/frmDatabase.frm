VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPlayer 
      Caption         =   "Player"
      Height          =   975
      Left            =   5640
      Picture         =   "frmDatabase.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Add"
      TabPicture(0)   =   "frmDatabase.frx":0884
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdBrowse"
      Tab(0).Control(1)=   "cmdSingle"
      Tab(0).Control(2)=   "cmdDir"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "frmDatabase.frx":08A0
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmbSearch"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdSearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "optSearch(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "optSearch(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "optSearch(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "optSearch(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.OptionButton optSearch 
         Caption         =   "Genre"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Song"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Album"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Artist"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox cmbSearch 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "Select..."
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   975
         Left            =   -72960
         Picture         =   "frmDatabase.frx":08BC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Browse"
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdSingle 
         Caption         =   "Single"
         Height          =   975
         Left            =   -73920
         Picture         =   "frmDatabase.frx":0CFE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Sibgle"
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "Dir"
         Height          =   975
         Left            =   -74880
         Picture         =   "frmDatabase.frx":1140
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Directory"
         Top             =   840
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList imlMusic 
      Left            =   3480
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":1582
            Key             =   "Artist"
            Object.Tag             =   "Artist"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":1662
            Key             =   "Album"
            Object.Tag             =   "Album"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":1AB4
            Key             =   "Song"
            Object.Tag             =   "Song"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMusic 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7858
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoRS As ADODB.Recordset
Dim adoConn As ADODB.Connection
Dim SQL As String
Dim DSN As String

Private Sub cmbSearch_Click()
    cmdSearch.Enabled = True
End Sub

Private Sub Form_Load()
    Set adoRS = CreateObject("ADODB.Recordset")
    Set adoConn = CreateObject("ADODB.Connection")
    
    DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Music.mdb;Persist Security Info=False"
    adoConn.Open DSN
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    adoRS.Close
    adoConn.Close
    Set adoRS = Nothing
    Set adoConn = Nothing
End Sub

Private Sub cmdPlayer_Click()
    frmPlayer.Show
End Sub

Private Sub cmdSearch_Click()
    Dim intLoop As Integer, Index As Integer
    For intLoop = optSearch.LBound To optSearch.UBound
        If optSearch(intLoop).Value = True Then Index = intLoop
    Next intLoop
    
    ' Get ID of type selected
    Select Case Index
        Case 0
            SQL = "SELECT FName, LName, ID FROM tblArtists WHERE LName = '" & cmbSearch.Text & "' ORDER by LName ASC, FName"
        Case 1
           SQL = "SELECT Name, ID FROM tblAlbums WHERE Name = '" & cmbSearch.Text & "' ORDER by Name ASC"
        Case 2
            SQL = "SELECT Name, ID FROM tblSongs WHERE Name = '" & cmbSearch.Text & "' ORDER by Name ASC"
        Case 3
            SQL = "SELECT Name, ID FROM tblGenre WHERE Name = '" & cmbSearch.Text & "' ORDER by Name ASC"
        Case 4
    End Select
    MsgBox SQL
End Sub

Private Sub optSearch_Click(Index As Integer)
    Dim strType As String, strTable As String
    cmbSearch.Clear
    cmbSearch.Text = "Select..."
    Select Case Index
        Case 0
            SQL = "SELECT FName, LName FROM tblArtists ORDER by LName ASC, FName"
        Case 1
           SQL = "SELECT Name FROM tblAlbums ORDER by Name ASC"
        Case 2
            SQL = "SELECT Name FROM tblSongs ORDER by Name ASC"
        Case 3
            SQL = "SELECT Name FROM tblGenre ORDER by Name ASC"
        Case 4
    End Select
    
    adoRS.Open SQL, adoConn
    
    Do Until adoRS.EOF
        cmbSearch.AddItem adoRS(0)
        adoRS.MoveNext
    Loop
    adoRS.Close
    cmbSearch.Enabled = True
End Sub

