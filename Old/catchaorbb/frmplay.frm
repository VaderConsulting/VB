VERSION 5.00
Begin VB.Form frmPlay 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play the Game"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7800
   Icon            =   "frmPlay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPlay.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   6030
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStar 
      Left            =   1080
      Top             =   2400
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00008000&
      Caption         =   "&Reset"
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer tmrTotal 
      Enabled         =   0   'False
      Interval        =   12500
      Left            =   1080
      Top             =   1920
   End
   Begin VB.Timer tmrRan 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   1440
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblTime2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time Remaining :"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image imgOrbb2 
      Height          =   1335
      Left            =   120
      Picture         =   "frmPlay.frx":0D0C
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   11
      Left            =   3720
      Picture         =   "frmPlay.frx":7012
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   10
      Left            =   3360
      Picture         =   "frmPlay.frx":7F2C
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   9
      Left            =   3000
      Picture         =   "frmPlay.frx":8E46
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   8
      Left            =   2640
      Picture         =   "frmPlay.frx":9D60
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   7
      Left            =   2280
      Picture         =   "frmPlay.frx":AC7A
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   6
      Left            =   1920
      Picture         =   "frmPlay.frx":BB94
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   5
      Left            =   1560
      Picture         =   "frmPlay.frx":CAAE
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   4
      Left            =   1200
      Picture         =   "frmPlay.frx":D9C8
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   12
      Left            =   4080
      Picture         =   "frmPlay.frx":E8E2
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   3
      Left            =   840
      Picture         =   "frmPlay.frx":F7FC
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   1
      Left            =   120
      Picture         =   "frmPlay.frx":10716
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgOrb 
      Height          =   750
      Index           =   2
      Left            =   480
      Picture         =   "frmPlay.frx":11630
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewgame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
         Begin VB.Menu mnuEasy 
            Caption         =   "&Easy"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuDifficult 
            Caption         =   "&Difficult"
         End
      End
      Begin VB.Menu mnuTop5 
         Caption         =   "&Top 5 Scores"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuInstructions 
         Caption         =   "&Instructions"
      End
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************'
'Filename :       frmPlay.frm                                '
'Author:          Coenraad Steenkamp                         '
'Created:         09/05/2000                                 '
'Operating System Windows 9x & NT ,2000 ,Me                  '
'                                                            '
'Description:                                                '
'This form gives access to the instructions form , a top form'
'and the main menu.When you press on new game you can start  '
'a new game there is two difficult levels that you can play  '
'on easy and difficult .You have 12 seconds on each level.   '
'                                                            '
'                                                            '
'************************************************************'
Option Explicit
   'Declare Variables for the starfield
   Dim I As Integer
   Dim X(100), Y(100), Z(100) As Integer
   Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
   'Declare constants for Both the music sounds
   Const Music = "passport.mid"
   Const Music2 = "shot.wav"
   Dim Sound
Private Sub cmdreset_Click()
   'Declare variables
   Dim X As Integer
   tmrRan.Enabled = False
   tmrTotal.Enabled = False
   'Resets the score to 0
   lblScore.Caption = 0
   'Makes the images invisible to the user
   'when he clicks on the reset button
   If mnuEasy.Checked = True Then
       For X = 1 To 12
           imgOrb(X).Visible = False
       Next X
   Else
       For X = 1 To 24
           imgOrb(X).Visible = False
       Next X
   End If
   lblTime.Caption = 12
End Sub
Private Sub Form_Activate()
   'Set's the speed for the Starfield - lets it go
   'backwards + lets it go forward
   iSpeed = -20
   iK = 2038
   iZoom = 256
   tmrStar.Interval = 1
   For I = 0 To 100
      X(I) = Int(Rnd * 1024) - 512
      Y(I) = Int(Rnd * 1024) - 512
      Z(I) = Int(Rnd * 512) - 256
   Next I
End Sub
Private Sub Form_Load()
   'Play Sound
   Sound = mciSendString("play " & App.Path & "\snd\" & Music, 0&, 0, 0)
   'Set the number of orb to 12
   gOrb = 12
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'When user clicks with his mouse anyware on the game form
   'a popup menu will appear
   If Button = vbRightButton Then
       PopupMenu mnuGame
   End If
End Sub
Private Sub imgOrb_Click(Index As Integer)
   'PLay sound
   Sound = mciSendString("play " & App.Path & "\snd\" & Music2, 0&, 0, 0)
   PlaySound "snd\shot.wav", 1, 2
   'Increments the game score with 1
   lblScore.Caption = lblScore.Caption + 1
   frmPlay.Caption = "Play the Game " & "  Score > " & lblScore.Caption
   'Makes Image orb invisible
   imgOrb(Index).Visible = False
   'Sets the property for image Orbb 2
   With imgOrbb2
      .Top = imgOrb(Index).Top
      .Left = imgOrb(Index).Left
      .Visible = True
   End With
End Sub
Private Sub mnuDifficult_Click()
   Dim iY As Integer
   'Loads the controls on the fly
   If mnuDifficult.Checked = True Then
      MsgBox "The level is on Difficult ", vbOKOnly, "Level Allready Set"
   Else
      For iY = 13 To 24
         Load imgOrb(iY)
      Next iY
      'Makes the level difficult
      tmrTotal.Interval = 19000
      tmrRan.Interval = 1500
      mnuEasy.Checked = False
      mnuDifficult.Checked = True
      gOrb = 24
      If tmrRan = True Then
         MsgBox "Please press reset!! ", vbOKOnly, "Press Reset"
      Else
      End If
   End If
End Sub
Private Sub mnuEasy_Click()
   'Makes the level Easy
   mnuEasy.Checked = True
   mnuDifficult.Checked = False
   gOrb = 12
   If tmrRan = True Then
       MsgBox "Please press reset!! ", vbOKOnly, "Press Reset"
   Else
   End If
End Sub
Private Sub mnuExit_Click()
   'Declare Local Variables
   Dim strAnswer As String
   'Ask's the user if he wants to Exit
   strAnswer = MsgBox("Are you sure?", vbYesNo, "Confirmation...")
   If strAnswer = vbYes Then
      'Stop Sound
      Sound = mciSendString("close " & App.Path & "\snd\" & Music, 0&, 0, 0)
      'Unloads the current Form
      Unload frmPlay
      'Show the main form
      frmIntro.Show
   End If
End Sub
Private Sub mnuInstructions_Click()
   'Stop the sound
   Sound = mciSendString("close " & App.Path & "\snd\" & Music, 0&, 0, 0)
   'Show the instruction form and unload the play form
   frmInstruc.Show
   Unload frmPlay
End Sub
Private Sub mnuNewgame_Click()
   'Declare the Variables
   Dim iX As Integer
   lblTime.Caption = 12
   ' Initialize random-number generator.
   Randomize
   'Randomize the positions of the orbbs
   For iX = 1 To gOrb
       imgOrb(iX).Top = ((Rnd * frmPlay.ScaleHeight - imgOrb(1).Height)) + 250
       imgOrb(iX).Left = ((Rnd * frmPlay.ScaleWidth - imgOrb(1).Width))
   Next iX
   'Resets the score to 0
   lblScore.Caption = 0
   'Enable my random timer
   tmrRan.Enabled = True
   tmrTotal.Enabled = True
End Sub
Private Sub mnuTop5_Click()
   'Stop te sound
   Sound = mciSendString("close " & App.Path & "\snd\" & Music, 0&, 0, 0)
   'Unload the form
   Unload frmPlay
   'Declare the variables
   Dim iX As Integer, iY As Integer, udtTemp As Recordings
   'Read the records into an Array
   For iX = 1 To giNumRecords
       Get #miFileNum, iX, gudRec
       gudtRecordArray(iX).StrName = gudRec.StrName
       gudtRecordArray(iX).iScore = gudRec.iScore
   Next iX
   'Sort's the records
   For iX = 1 To giNumRecords - 1
       For iY = iX + 1 To giNumRecords
           If gudtRecordArray(iX).iScore < gudtRecordArray(iY).iScore Then
               udtTemp = gudtRecordArray(iX)
               gudtRecordArray(iX) = gudtRecordArray(iY)
               gudtRecordArray(iY) = udtTemp
           End If
       Next iY
   Next iX
   'Display the Records on the labels
   For iX = 1 To 5
       frmTopF.lblName(iX) = gudtRecordArray(iX).StrName
       frmTopF.lblScore(iX) = gudtRecordArray(iX).iScore
   Next iX
   'Show's the Top 5 form
   frmTopF.Show
End Sub
Private Sub tmrStar_Timer()
   For I = 0 To 100
      Circle (tmpX(I), tmpY(I)), 5, BackColor
      Z(I) = Z(I) + iSpeed
      If Z(I) > 255 Then Z(I) = -255
      If Z(I) < -255 Then Z(I) = 255
      tmpZ(I) = Z(I) + iZoom
      tmpX(I) = (X(I) * iK / tmpZ(I)) + (frmPlay.Width / 2)
      tmpY(I) = (Y(I) * iK / tmpZ(I)) + (frmPlay.Height / 2)
      iRadius = 1
      iStarcolor = 256 - Z(I)
      Circle (tmpX(I), tmpY(I)), 5, RGB(iStarcolor, iStarcolor, iStarcolor)
    Next I
End Sub
Private Sub tmrRan_Timer()
   'Declare Variables
   Dim iX As Integer
   Dim iRndNum As Integer
   'Makes the images invisible
   For iX = 1 To gOrb
       imgOrb(iX).Visible = False
   Next iX
   'If gOrb is 24 then display the new pictures
   If gOrb = 24 Then
       For iX = 1 To 2
           'Generate's the random number generator
           Randomize
           'Generate a random number between 1 and gOrb
           iRndNum = Int(gOrb * Rnd + 1)
           imgOrb(iRndNum).Visible = True
           imgOrbb2.Visible = False
       Next iX
   Else
   'Generate's the random number generator
   Randomize
   'Generate a random number between 1 and gOrb
   iRndNum = Int(gOrb * Rnd + 1)
   imgOrb(iRndNum).Visible = True
   imgOrbb2.Visible = False
   End If
   lblTime.Caption = lblTime.Caption - 1
End Sub
Private Sub tmrTotal_Timer()
   'Error handeling goto EH
   On Error GoTo EH
   'Declare the Varibles
   Dim iX As Integer
   Dim iRecLength As Integer
   Dim strMsg As String
   'Disable The random timer
   tmrRan.Enabled = False
   'Disable the Total Timer
   tmrTotal.Enabled = False
   'Makes all the images invisible
   If mnuEasy.Checked = True Then
      For iX = 1 To 12
          imgOrb(iX).Visible = False
      Next iX
   Else
      For iX = 1 To 24
          imgOrb(iX).Visible = False
      Next iX
   End If
   'Display a mesage box that say's the user is game over
   MsgBox " Score " & lblScore.Caption, , "GAME OVER"
   'Ask the person for his name
   gudRec.StrName = (InputBox("Please enter your name ?", _
   "Name Please?", , 5000, 4000))
   miFileNum = FreeFile
   iRecLength = Len(gudRec)
   'Open the file from the Disk
   Open "A:\Orbb.txt" For Random As #miFileNum Len = iRecLength
   'Error Handeling for "Disk not ready"
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
   'Get the informatin from the current game
   gudRec.StrName = gudRec.StrName
   gudRec.iScore = lblScore.Caption
   'Increment the number of records
   giNumRecords = giNumRecords + 1
   'Put the records on the disk
   Put miFileNum, giNumRecords, gudRec
End Sub
