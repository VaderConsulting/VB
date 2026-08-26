VERSION 5.00
Begin VB.Form frmEnterIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Server IP Address"
   ClientHeight    =   1665
   ClientLeft      =   1380
   ClientTop       =   3945
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1665
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "10.1.1.1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   $"enterip.frx":0000
      Height          =   975
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmEnterIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Form1.SocketX.RemoteAddress = ""
   Unload Me
End Sub


Private Sub cmdOK_Click()
   Form1.SocketX.RemoteAddress = Trim(txtIP)
   Unload Me
End Sub

