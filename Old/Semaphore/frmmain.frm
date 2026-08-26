VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Semaphore"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5880
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFlags 
      Caption         =   "Flags"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   5640
      Width           =   735
   End
   Begin VB.CheckBox chkHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox txtType 
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   5640
      Width           =   4935
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   4680
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5280
      Top             =   4080
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   9120
      X2              =   0
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   9120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image imgStart 
      Height          =   495
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgFlag 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgSemaphore 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5040
      Left            =   120
      Picture         =   "frmMain.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4995
   End
   Begin VB.Label lblText 
      Caption         =   "Text to send"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bStart As Boolean, bContinue As Boolean

Private Sub cmdStart_Click()
  Dim strLetter As String
  tmrStart.Enabled = True
  txtSend.Enabled = False
  Do Until bStart = True
    DoEvents
  Loop
  lblHelp.Caption = ""
  If chkHelp.Value = vbChecked Then lblHelp.Visible = True
  If chkFlags.Value = vbChecked Then imgFlag.Visible = True
  imgStart.Visible = False
  For lp = 1 To Len(txtSend)
    strLetter = LCase(Mid(txtSend, lp, 1))
    lblHelp = UCase(strLetter)
    If strLetter <> " " Then
      On Error Resume Next
      imgSemaphore.Picture = LoadPicture("U:\Semaphore\" & strLetter & ".gif")
      imgFlag.Picture = LoadPicture("U:\Flags\" & strLetter & ".gif")
      On Error GoTo 0
      imgSemaphore.Refresh
      tmrWait.Enabled = True
      Do Until bContinue = True
        DoEvents
      Loop
      bContinue = False
    End If
  Next lp
  txtSend.Enabled = True
  imgSemaphore.Picture = LoadPicture("U:\Semaphore\wait.gif")
  imgFlag.Visible = False
  lblHelp.Visible = False
  tmrStart.Enabled = False
  tmrWait.Enabled = False
  bStart = False
  bContinue = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  txtType.Text = txtType.Text & Chr(KeyAscii)
End Sub

Private Sub Form_Load()
  imgSemaphore.Picture = LoadPicture("U:\Semaphore\wait.gif")
End Sub

Private Sub mnuExit_Click()
  Unload Me
  End
End Sub

Private Sub mnuOpen_Click()
  MsgBox "Put code here to open appropriate text files"
End Sub

Private Sub tmrStart_Timer()
  Static iInterval As Integer
  iInterval = iInterval + 1
  imgStart.Picture = LoadPicture("U:\Semaphore\" & 6 - iInterval & ".gif")
  imgStart.Visible = True
  If iInterval <> 5 Then
    tmrStart.Enabled = True
    Exit Sub
  End If
  bStart = True
  tmrStart.Enabled = False
  iInterval = 0
End Sub

Private Sub tmrWait_Timer()
  bContinue = True
  tmrWait.Enabled = False
End Sub
