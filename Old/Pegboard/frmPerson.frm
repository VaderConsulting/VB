VERSION 5.00
Begin VB.Form frmPerson 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2055
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
   Begin VB.Image imgRank 
      Height          =   735
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgPicture 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblBillet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblRankName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblNumber 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    a = Val(lblNumber)
    lblRankName = gRank(a) & " " & gName(a)
    lblBillet = gBilletNumber(a) & " " & gBilletText(a)
    imgPicture = LoadPicture("c:\DemoPic.BMP")
    If gRank(a) = "CMDR" Then imgRank.Picture = frmLeave.imgCMDR.Picture
    If gRank(a) = "LCDR" Then imgRank.Picture = frmLeave.imgLCDR.Picture
    If gRank(a) = "LEUT" Then imgRank.Picture = frmLeave.imgLEUT.Picture
    If gRank(a) = "SBLT" Then imgRank.Picture = frmLeave.imgSBLT.Picture
    If gRank(a) = "MIDN" Then imgRank.Picture = frmLeave.imgMIDN.Picture
End Sub
