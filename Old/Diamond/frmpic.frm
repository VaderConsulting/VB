VERSION 5.00
Begin VB.Form frmPic 
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgPause 
      Height          =   480
      Left            =   600
      Picture         =   "frmPic.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgDiamond 
      Height          =   1800
      Left            =   120
      Picture         =   "frmPic.frx":0442
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Image imgStop 
      Height          =   480
      Left            =   120
      Picture         =   "frmPic.frx":2A52
      Top             =   720
      Width           =   480
   End
   Begin VB.Image imgStart 
      Height          =   480
      Left            =   120
      Picture         =   "frmPic.frx":2E94
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
