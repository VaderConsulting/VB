VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proxy Switcher"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   1830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "New..."
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox cmbConfig 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNew_Click()
    frmAddProxy.Show vbModal
End Sub
