VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3975
   ScaleWidth      =   5040
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstPlay 
      Height          =   2790
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Idle"
            TextSave        =   "Idle"
            Object.ToolTipText     =   "Application Status"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgTrash 
      Height          =   480
      Left            =   4440
      Picture         =   "frmMain.frx":0442
      Top             =   3000
      Width           =   480
   End
   Begin MediaPlayerCtl.MediaPlayer mpPlayer 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4290
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   -1  'True
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
      InvokeURLs      =   0   'False
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
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -110
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AddFile(strFilename As String, strFiletype As String)
    Select Case strFiletype
        Case ""
            ' Invalid
        Case "txt"
            ' Text file
        Case "mid"
            ' MIDI file
            AddtoPlaylist strFilename
        Case "wav"
            ' WAVE file
            AddtoPlaylist strFilename
        Case "mp3"
            ' MP3 file
            AddtoPlaylist strFilename
         Case "doc"
            ' DOC file
        Case Else
            ' unrecognised file
    End Select
End Sub

Sub AddtoPlaylist(strFilename As String)
    lstPlay.AddItem strFilename
    sbrStatus.SimpleText = strFilename & " added to playlist."
End Sub

Public Function GetDroppedFiles(Files As Collection)
    Dim strFilename As String
    Dim strFiletype As String
    Dim bResult As Boolean
    For lp = 1 To Files.Count
        strFilename = Files.Item(lp)
        bResult = GetFileType(strFilename, strFiletype)
        If bResult Then
            AddFile strFilename, strFiletype
        Else
            strFiletype = ""     ' invalid result = invalid data
        End If
    Next lp
End Function

Public Function GetFileType(strFilename As String, strFiletype As String) As Boolean
    Dim p As Integer
    GetFileType = False
    p = InStr(strFilename, ".")
    If p = 0 Then Exit Function
    strFiletype = LCase(Right(strFilename, Len(strFilename) - p))
    GetFileType = True
End Function

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim f As New Collection
    For lp = 1 To Data.Files.Count
        f.Add Data.Files(lp)
    Next lp
    GetDroppedFiles f
End Sub

Private Sub imgTrash_Click()
    Dim strFilename As String
    If lstPlay.ListIndex = -1 Then Exit Sub
    
    strFilename = lstPlay.List(lstPlay.ListIndex)
    mpPlayer.Stop
    sbrStatus.SimpleText = "Removed " & strFilename & " from playlist."
    lstPlay.RemoveItem lstPlay.ListIndex
End Sub

Private Sub lstPlay_Click()
    'Dim strFilename As String
    'strFilename = lstPlay.List(lstPlay.ListIndex)
    'mpPlayer.Stop
    'mpPlayer.FileName = strFilename
    'sbrStatus.SimpleText = strFilename & " waiting to play."
End Sub

Private Sub lstPlay_DblClick()
    Dim strFilename As String
    strFilename = lstPlay.List(lstPlay.ListIndex)
    mpPlayer.Stop
    mpPlayer.FileName = strFilename
    mpPlayer.Play
    sbrStatus.SimpleText = strFilename & " playing."
End Sub

Private Sub lstPlay_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim f As New Collection
    For lp = 1 To Data.Files.Count
        f.Add Data.Files(lp)
    Next lp
    GetDroppedFiles f
End Sub

Private Sub mpPlayer_EndOfStream(ByVal Result As Long)
    Dim strFilename As String
    If lstPlay.ListIndex <> lstPlay.ListCount - 1 Then
        lstPlay.ListIndex = lstPlay.ListIndex + 1
    Else
        lstPlay.ListIndex = 0
    End If
    
    strFilename = lstPlay.List(lstPlay.ListIndex)
    mpPlayer.FileName = strFilename
    mpPlayer.Play
    sbrStatus.SimpleText = strFilename & " playing."
End Sub
