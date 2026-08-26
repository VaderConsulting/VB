VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   6960
      Top             =   2400
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5400
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdWriteTage 
      Caption         =   "Save Tag"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtGenreCode 
      Height          =   285
      Left            =   6960
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   5400
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtAlbum 
      Height          =   285
      Left            =   5400
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   5400
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2880
      Pattern         =   "*.mp3"
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblMsg 
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Label lblElapsedTime 
      Caption         =   "Elapsed Time"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblTotalTime 
      Caption         =   "Total Time"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   3360
      Width           =   1455
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   7335
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -150
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GenresTypes
Dim Min As Integer
Dim Sec As Integer

Dim FileName As String
Dim FileOpen As Boolean
Dim CurrentTag As TagInfo

Private Type TagInfo
  Tag As String * 3
  Songname As String * 30
  artist As String * 30
  album As String * 30
  year As String * 4
  comment As String * 30
  genre As String * 1
End Type

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
  Dim temp As String
  On Error Resume Next
  EraseTXTBoxes
  
  If Right(Dir1.Path, 1) = "\" Then
    FileName = Dir1.Path & File1.FileName
  Else
    FileName = Dir1.Path & "\" & File1.FileName
  End If
      
  Open FileName For Binary As #1
  With CurrentTag
    Get #1, FileLen(FileName) - 127, .Tag
    If Not .Tag = "TAG" Then
      lblMsg.Caption = "No tag"
      Close #1
      Exit Sub
    End If
    Get #1, , .Songname
    Get #1, , .artist
    Get #1, , .album
    Get #1, , .year
    Get #1, , .comment
    Get #1, , .genre
    Close #1
  
    txtTitle = RTrim(.Songname)
    txtArtist = RTrim(.artist)
    txtAlbum = RTrim(.album)
    txtYear = RTrim(.year)
    txtComment = RTrim(.comment)
    
    temp = RTrim(.genre)
    txtGenreCode = Asc(temp)
    Combo1.ListIndex = CInt(txtGenreCode) - 1
  End With
End Sub
Private Sub Dir1_change()
  File1.FileName = Dir1.Path
End Sub

Private Sub cmdWriteTag_Click()
  If FileOpen Then
    MsgBox "You can't save to an open file", vbCritical, "MP3 Tag Save Error"
    Exit Sub
  End If
      
  If Right(Dir1.Path, 1) = "\" Then
    FileName = Dir1.Path & File1.FileName
  Else
    FileName = Dir1.Path & "\" & File1.FileName
  End If
  
  With CurrentTag
    .Tag = "TAG"
    .Songname = txtTitle
    .artist = txtArtist
    .album = txtAlbum
    .year = txtYear
    .comment = txtComment
    .genre = Chr(Combo1.ListIndex + 1)
      
    Open FileName For Binary Access Write As #1
      Seek #1, FileLen(FileName) - 127
      Put #1, , .Tag
      Put #1, , .Songname
      Put #1, , .artist
      Put #1, , .album
      Put #1, , .year
      Put #1, , .comment
      Put #1, , .genre
    Close #1
  End With
End Sub

Private Sub cmdOpen_Click()
  With MediaPlayer1
    If Not FileOpen Then
      .FileName = FileName
      .AutoStart = True
      cmdOpen.Caption = "Close"
    Else
      .FileName = ""
      cmdOpen.Caption = "Open"
    End If
  End With
End Sub

Private Sub MediaPlayer1_OpenStateChange(ByVal OldState As Long, ByVal NewState As Long)
  
  Min = MediaPlayer1.Duration \ 60
  Sec = MediaPlayer1.Duration - (Min * 60)
  lblTotalTime = "Total Time: " & Format(Min, "0#") _
    & ":" & Format(Sec, "0#") 'format time to 00:00
      
  FileOpen = CBool(NewState)
End Sub

Private Sub Timer1_Timer()
  Min = MediaPlayer1.CurrentPosition \ 60
  Sec = MediaPlayer1.CurrentPosition - (Min * 60)
  If Min > 0 Or Sec > 0 Then
    lblElapsedTime = "Elapsed Time: " & Format(Min, "0#") & ":" & Format(Sec, "0#")
  Else
    lblElapsedTime = "Elapsed Time: 00:00"
  End If
End Sub

Private Sub EraseTXTBoxes()
  lblMsg.Caption = ""
  txtTitle = ""
  txtArtist = ""
  txtAlbum = ""
  txtYear = ""
  txtComment = ""
  txtGenreCode = ""
  Combo1.ListIndex = -1
End Sub

Private Sub Form_Load()
  Dim x As Integer
  Dim iLower As Integer
  Dim iUpper As Integer
  
  Drive1.Drive = "C:"
  Dir1.Path = "C:\"
  
  GenresTypes = Array("Blues", "Classic Rock", "Country", _
    "Dance", "Disco", "Funk", "Grunge", "Hip -Hop", _
    "Jazz", "Metal", "New Age", "Oldies", "Other", _
    "Pop", "R&b", "Rap", "Reggae", "Rock", "Techno", _
    "Industrial", "Alternative", "Ska", "Death Metal", _
    "Pranks", "Soundtrack", "Euro -Techno", "Ambient", _
    "Trip -Hop", "Vocal", "Jazz Funk", "Fusion", _
    "Trance", "Classical", "Instrumental")
  
  iLower = LBound(GenresTypes)
  iUpper = UBound(GenresTypes)
  For x = iLower To iUpper
    Combo1.AddItem GenresTypes(x)
  Next x
End Sub



