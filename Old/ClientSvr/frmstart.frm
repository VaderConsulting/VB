VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Client-Server"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   2160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdServer 
      Caption         =   "Server"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "Client"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblRole 
      Alignment       =   2  'Center
      Caption         =   "Select a role"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClient_Click()
    frmStart.Hide
    frmClient.Show
    frmServer.Show
End Sub

Private Sub cmdServer_Click()
    frmStart.Hide
    frmServer.Show
End Sub

