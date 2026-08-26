VERSION 5.00
Begin VB.Form frmScores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scores Help"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   9
      Left            =   0
      TabIndex        =   8
      Top             =   3960
      Width           =   6000
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   8
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   6000
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   6000
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   6000
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   6000
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   6000
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   6000
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   6000
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6000
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    frmPERS.SetFocus
End Sub

