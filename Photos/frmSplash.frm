VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Splash"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   5000
      Left            =   8400
      Top             =   3840
   End
   Begin VB.Image imgCrest 
      Height          =   8340
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgCrest_Click()
    Unload Me
End Sub

Private Sub tmrSplash_Timer()
    Unload Me
End Sub
