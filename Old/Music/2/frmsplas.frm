VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   3105
   ClientTop       =   2595
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3375
   ScaleWidth      =   4665
   Begin VB.Timer timSplash 
      Interval        =   2500
      Left            =   4080
      Top             =   0
   End
   Begin VB.Image Image3 
      Height          =   1560
      Left            =   1680
      Picture         =   "FRMSPLAS.frx":0000
      Top             =   120
      Width           =   2835
   End
   Begin VB.Image Image2 
      Height          =   1740
      Left            =   960
      Picture         =   "FRMSPLAS.frx":1EC8
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2790
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   120
      Picture         =   "FRMSPLAS.frx":24A2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
    Rem frmHelp.lblHelp.Caption = "Don't forget to load the database..."
    Rem frmHelp.Show 1
    frmSplash.Hide
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
    timSplash.Enabled = 1
    Location = CurDir()
    Rem soundname$ = "D:\PROJECTS\MUSICD~1\" & "START.WAV"
    soundname$ = Location & "\START.WAV"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Rem X% = sndPlaySound(soundname$, wFlags%)
End Sub


Private Sub timSplash_Timer()
    timSplash.Enabled = 0
    frmSplash.Hide
End Sub
