VERSION 5.00
Object = "{18D91AD0-D0BE-11D1-A6B4-00AA002075DA}#1.0#0"; "flshtray.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Tusk Marquee Settings"
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   1305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TrayIconPrj.TrayIcon TrayIcon1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1905
      _ExtentY        =   953
      Icon            =   "frmMain.frx":08CA
      ToolTipText     =   "DOJ Marquee Settings"
      Enabled         =   -1  'True
      TrueClick       =   -1  'True
      Visible         =   -1  'True
      FlashSound      =   0
      FlashIcon       =   "frmMain.frx":0BE4
      FlashInterval   =   1000
      FlashEnabled    =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Hide
End Sub

Private Sub TrayIcon1_LeftButtonClick()
    Dim cmd As String
    cmd = "c:\winnt\system32\DOJ Marquee.scr /c"
    Shell cmd
End Sub

Private Sub TrayIcon1_LeftButtonDoubleClick()
    Dim cmd As String
    cmd = "c:\winnt\system32\DOJ Marquee.scr /c"
    Shell cmd
End Sub
