VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP4 Display"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptDisplay 
      Caption         =   "Off"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.OptionButton OptDisplay 
      Caption         =   "Off"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image imgMain 
      Height          =   4305
      Left            =   120
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgMain_Click()

End Sub
