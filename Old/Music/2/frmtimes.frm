VERSION 5.00
Begin VB.Form frmTimesUp 
   Caption         =   "Your time is up"
   ClientHeight    =   3795
   ClientLeft      =   2385
   ClientTop       =   3270
   ClientWidth     =   6690
   Icon            =   "FRMTIMES.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3795
   ScaleWidth      =   6690
   Begin VB.CommandButton cmdOK 
      Caption         =   "Oh Well"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "If you registered this great program, you would have more than a measly 5 minutes"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Times up!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
   End
End
Attribute VB_Name = "frmTimesUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    End
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub
