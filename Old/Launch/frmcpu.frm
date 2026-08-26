VERSION 5.00
Begin VB.Form frmCPU 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   990
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   15
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   15
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   14
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   14
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   13
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   13
      Top             =   240
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   12
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   12
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   11
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   11
      Top             =   480
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   10
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   600
      Width           =   135
   End
   Begin VB.Timer tmrCPU 
      Interval        =   150
      Left            =   240
      Top             =   120
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   0
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   9
      Top             =   1800
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   1
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   1680
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   2
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   1560
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   3
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   1440
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   4
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   1320
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   5
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   1200
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   6
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   1080
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   7
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   960
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   8
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   840
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   9
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "frmCPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Width = 135
    Top = Screen.Height - Height - 400
    Left = Screen.Width - Width
    OnTop
End Sub
