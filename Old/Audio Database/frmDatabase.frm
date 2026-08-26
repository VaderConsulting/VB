VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13920
   Icon            =   "frmDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   13920
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl mciPlayer 
      Height          =   615
      Left            =   8520
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   0
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Shareable       =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame fmeKey 
      Caption         =   "Key"
      Height          =   2655
      Left            =   12000
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      Begin VB.Label Label3 
         Caption         =   "Filename"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   120
         Picture         =   "frmDatabase.frx":57E2
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Lyrics"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   120
         Picture         =   "frmDatabase.frx":BA6C
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lblMIDI 
         Caption         =   "MIDI"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   120
         Picture         =   "frmDatabase.frx":1167E
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblSong 
         Caption         =   "Song"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.Image imgSong 
         Height          =   255
         Left            =   120
         Picture         =   "frmDatabase.frx":17908
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblFolder 
         Caption         =   "Folder"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Image imgFolder 
         Height          =   255
         Left            =   120
         Picture         =   "frmDatabase.frx":1D0EA
         Stretch         =   -1  'True
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblDatabase 
         Caption         =   "Database"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgDatabase 
         Height          =   250
         Left            =   120
         Picture         =   "frmDatabase.frx":23374
         Stretch         =   -1  'True
         Top             =   360
         Width           =   250
      End
   End
   Begin MSComctlLib.ImageList imlMusic 
      Left            =   6720
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":295FE
            Key             =   "MIDI"
            Object.Tag             =   "MIDI"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":2F898
            Key             =   "Song"
            Object.Tag             =   "Song"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":3508A
            Key             =   "Lyrics"
            Object.Tag             =   "Lyrics"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":3ACAC
            Key             =   "Filename"
            Object.Tag             =   "Filename"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":40F46
            Key             =   "Folder"
            Object.Tag             =   "Folder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":471E0
            Key             =   "Database"
            Object.Tag             =   "Database"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMusic 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   16325
      _Version        =   393217
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imlMusic"
      Appearance      =   1
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   9480
      Width           =   13695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   20000
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   20000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
         Begin VB.Menu mnuDirectory 
            Caption         =   "&Directory"
         End
         Begin VB.Menu mnuSingle 
            Caption         =   "&Single"
         End
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Db"
      End
      Begin VB.Menu mnuPrefs 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "S&earch"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPlay 
      Caption         =   "Play"
      Visible         =   0   'False
      Begin VB.Menu mnuPlayMIDI 
         Caption         =   "Play MIDI"
      End
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDir_Click()
    frmBrowse.Show
End Sub

Private Sub cmdSearch_Click()
    Dim intLoop As Integer, Index As Integer
    
    MsgBox SQL
End Sub

Private Sub Form_Load()
    Dim N As Node, intLoop As Integer, strTemp As String
    
    Me.Show
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    adoRS.Close
    adoConn.Close
    Set adoRS = Nothing
    Set adoConn = Nothing
End Sub

Private Sub mnuDirectory_Click()
    frmBrowse.Show
End Sub

Private Sub mnuExit_Click()
    Dim frm As Form
    frmDatabase.mciPlayer.Command = "Close"
    For Each frm In VB.Forms
        Unload frm
    Next
End Sub

Private Sub mnuOpen_Click()
    Dim N As Node, TargetNode As String, intLoop As Integer, strTemp As String
    Dim Songname As String, Filename As String, ismIDI As Boolean, isLyrics As Boolean
    Dim intCounter As Integer
    
    adoConn.Open DSN
    
    tvwMusic.Nodes.Clear
    
    Set N = tvwMusic.Nodes.Add(, , Db, Db, "Database")
    For intLoop = 0 To 25
        strTemp = Chr(intLoop + Asc("A"))
        Set N = tvwMusic.Nodes.Add(Db, tvwChild, "Child" & strTemp, strTemp, "Folder")
        DoEvents
    Next intLoop
    For intLoop = 0 To 9
        strTemp = CStr(intLoop)
        Set N = tvwMusic.Nodes.Add(Db, tvwChild, "Child" & strTemp, strTemp, "Folder")
        DoEvents
    Next intLoop
    
    SQL = "SELECT * FROM tblSongs ORDER BY name"
    adoRS.Open SQL, adoConn
        Do Until adoRS.EOF
            TargetNode = UCase(Left(adoRS("Name"), 1))
            Filename = adoRS("filename")
            Songname = adoRS("name")
            TargetNode = "Child" & TargetNode
            ismIDI = adoRS("isMIDI")
            isLyrics = adoRS("isLyrics")
            
            On Error Resume Next
            
            Set N = tvwMusic.Nodes.Add(TargetNode, tvwChild, Songname, Songname, "Song")
            
            If ismIDI Then
                Set N = tvwMusic.Nodes.Add(Songname, tvwChild, "MIDI" & Filename, Songname, "MIDI")
                Set N = tvwMusic.Nodes.Add("MIDI" & Filename, tvwChild, "MIDIFilename" & Filename, Filename, "Filename")
            End If
            
            If isLyrics Then
                Set N = tvwMusic.Nodes.Add(Songname, tvwChild, "Lyrics" & Filename, Songname, "Lyrics")
                Set N = tvwMusic.Nodes.Add("Lyrics" & Filename, tvwChild, "Filename" & Filename, Filename, "Filename")
            End If
            On Error GoTo 0
            
            lblStatus = "Adding " & Songname & " (" & Filename & ")"
            adoRS.MoveNext
            intCounter = intCounter + 1
            
            ' Refresh the form every so often... makes for MUCH faster loading
            If intCounter = 100 Then
                frmDatabase.Refresh
                intCounter = 0
            End If
        Loop
    adoRS.Close
    adoConn.Close
    lblStatus = "Idle"
End Sub

Private Sub mnuPlayMIDI_Click()
    With mciPlayer
        .Notify = False
        .Wait = True
        .Shareable = False
        .DeviceType = "WaveAudio"
        .Filename = SelectedSong
        .Command = "Open"
    End With
End Sub

Private Sub mnuPrefs_Click()
    frmPrefs.Show
End Sub

Private Sub mnuSearch_Click()
    frmSearch.Show
End Sub

Private Sub tvwMusic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If tvwMusic.Nodes.Count = 0 Then Exit Sub
    If Button = vbRightButton Then
        If UCase(Right(tvwMusic.SelectedItem.Text, 3)) = "MID" Then
            SelectedSong = tvwMusic.SelectedItem.Text
            frmDatabase.PopupMenu mnuPlay, , x, y
        End If
    End If
End Sub
