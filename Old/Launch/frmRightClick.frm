VERSION 5.00
Begin VB.Form frmRightClick 
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   0
      Top             =   1680
   End
   Begin VB.CommandButton cmdDummy 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "properties"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label lblIcon 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label lblDirectory 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label lblProgram 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmRightClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    Top = frmStart.Top
    Left = frmStart.Left - Width
    tmrTime.Enabled = True
    Visible = True
End Sub

Private Sub tmrTime_Timer()
    frmRightClick.Visible = False
    Enabled = False
    Unload frmRightClick
End Sub
