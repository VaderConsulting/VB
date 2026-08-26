VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to Catch A Orbb"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ScMar ScMarquee1 
      Height          =   615
      Index           =   0
      Left            =   2880
      TabIndex        =   6
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      Caption         =   "Welcome To Catch a Orbb"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   20
   End
   Begin VB.CommandButton cmdExit 
      DownPicture     =   "frmIntro.frx":12FA
      Height          =   735
      Left            =   4560
      Picture         =   "frmIntro.frx":4DD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox picTarget 
      Height          =   3015
      Left            =   3600
      Picture         =   "frmIntro.frx":528E
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   2
      Tag             =   "1"
      Top             =   3600
      Width           =   3015
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drag over here to Activate the Blocks"
         Height          =   795
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      X1              =   8520
      X2              =   8400
      Y1              =   1920
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   8520
      X2              =   8640
      Y1              =   1920
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   8520
      X2              =   8520
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Label lblRead 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Please read the instructions first!!!"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblDragin2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drag this block into the Target area to see Instructions"
      Height          =   555
      Index           =   1
      Left            =   7320
      TabIndex        =   4
      Top             =   2880
      Width           =   2280
   End
   Begin VB.Label lblDragin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drag this block into the Target area to Play..."
      Height          =   435
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   2880
      Width           =   1680
   End
   Begin VB.Image imgOrbb2 
      Height          =   1455
      Index           =   1
      Left            =   9000
      Picture         =   "frmIntro.frx":5EA4
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   975
   End
   Begin VB.Image imgOrbb2 
      Height          =   1455
      Index           =   0
      Left            =   120
      Picture         =   "frmIntro.frx":C1AA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgIns 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   3060
      Left            =   6960
      Picture         =   "frmIntro.frx":124B0
      Tag             =   "2"
      Top             =   2040
      Width           =   3060
   End
   Begin VB.Label lblPrNameDate 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Coenraad  Steenkamp 09/05/2000"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image imgPlay 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   3060
      Left            =   120
      Picture         =   "frmIntro.frx":13042
      Tag             =   "1"
      Top             =   2040
      Width           =   3060
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************'
'Filename :       frmIntro.frm                               '
'Author:          Coenraad Steenkamp                         '
'Created:         09/05/2000                                 '
'Operating System Windows 9x & NT ,2000 ,Me                  '
'                                                            '
'Description:                                                '
'This form gives the user access to the instructions form and'
'the game.The user can drag-drop the images to the targer    '
'area to play .The form also include a scrolling marquee.    '
'                                                            '
'************************************************************'
Option Explicit
   'Declare Variable For the sound
   Dim Sound
   Const Music = "intro.wav"
Private Sub cmdExit_Click()
   'Declare Local Variables
   Dim strAnswer As String
   'Ask's the user if he wants to Exit
   strAnswer = MsgBox("Are you sure?", vbYesNo, "Confirmation...")
   If strAnswer = vbYes Then
      'Error handeling goto EH
      On Error GoTo EH
      Unload frmIntro
      'Error handle for "Disk not Ready"
EH:
      If Err.Number = 71 Then
         intRet = MsgBox("The drive is not ready.", _
                  vbExclamation + vbAbortRetryIgnore, "Disk not in Drive!")
         Select Case intRet
            Case vbRetry
               Resume
            Case vbAbort
               End
            Case vbIgnore
               End
         End Select
      End If
   End If
End Sub
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
   'Enables the DragDrop
   Source.Move X, Y
End Sub
Private Sub Form_Load()
   'Error handeling goto EH
   On Error GoTo EH
   'Declare varibles
   Dim iRecLength As Integer
   'Play Some music
   Sound = mciSendString("play " & App.Path & "\snd\" & Music, 0&, 0, 0)
   miFileNum = FreeFile
   iRecLength = Len(gudRec)
   'Open the file from the Disk
   Open "A:\Orbb.txt" For Random As #miFileNum Len = iRecLength
   'Error handle for "Disk not Ready"
EH:
   If Err.Number = 71 Then
      intRet = MsgBox("The drive is not ready.", _
               vbExclamation + vbAbortRetryIgnore, "Disk not in Drive!")
      Select Case intRet
         Case vbRetry
            Resume
         Case vbAbort
            End
         Case vbIgnore
            End
      End Select
   End If
   'Use the LOF()Function(Length of File)to Calculate number of records in the file
   giNumRecords = LOF(miFileNum) / iRecLength
   If giNumRecords < 0 Then
      Get miFileNum, giNumRecords, gudRec
   Else
      giNumRecords = 1
      'Sets the filename to dummy informtion
      gudRec.StrName = "Nexgen"
      gudRec.iScore = "24"
      Put miFileNum, giNumRecords, gudRec
   End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   'Stop the music
   Sound = mciSendString("close " & App.Path & "\snd\" & Music, 0&, 0, 0)
   'Close the file miFileNum
   Close #miFileNum
   Unload Me
End Sub
Private Sub picTarget_DragDrop(Source As Control, X As Single, Y As Single)
   'Stop the music
   Sound = mciSendString("close " & App.Path & "\snd\" & Music, 0&, 0, 0)
   Dim iX As Integer
   'If the tag value is 1 then frmplay will show
   'otherwise frmInstructions will Show
   If Val(Source.Tag) = 1 Then
      frmPlay.Show
      frmIntro.Hide
      For iX = 1 To giNumRecords
        'Read the records form the disk and write them
        'to a Array
        Get #miFileNum, iX, gudRec
        gudtRecordArray(iX).StrName = gudRec.StrName
        gudtRecordArray(iX).iScore = gudRec.iScore
      Next iX
   Else
      'Show the form instructions and Hide the Intro Form
      frmInstruc.Show
      frmIntro.Hide
   End If
End Sub
