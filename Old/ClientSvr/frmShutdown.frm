VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmShutdown 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System event"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "frmShutdown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pbrShutdown 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmrSec 
      Interval        =   1000
      Left            =   3000
      Top             =   360
   End
   Begin VB.Timer tmrShutdown 
      Left            =   120
      Top             =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmShutdown.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblShutdown 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    tmrShutdown.Enabled = True
    tmrShutdown.Interval = CInt(frmClient.txtShutdown) * 1000
    pbrShutdown.Max = frmClient.txtShutdown
    pbrShutdown.Value = frmClient.txtShutdown
End Sub

Private Sub tmrSec_Timer()
    lblShutdown = "Shutdown in .. " & pbrShutdown.Value
    If pbrShutdown.Value > 0 Then pbrShutdown.Value = pbrShutdown.Value - 1
End Sub

Private Sub tmrShutdown_Timer()
    uFlags = ewx_force Or ewx_logoff Or ewx_shutdown
    ExitWindowsEx uFlags, 0
End Sub
