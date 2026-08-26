VERSION 5.00
Begin VB.Form frmConfigure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmConfigure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3375
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "IP Port Addressing"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtClientPort 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "2000"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtHostPort 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Text            =   "3000"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblClient 
         Alignment       =   2  'Center
         Caption         =   "Client"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblPort 
         Alignment       =   2  'Center
         Caption         =   "Host"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    frmServer.Show
    Unload Me
End Sub

Private Sub txtClientPort_Change()
    frmServer.txtClientPort = txtClientPort
End Sub

Private Sub txtHostPort_Change()
    frmServer.txtHostPort = txtHostPort
End Sub
