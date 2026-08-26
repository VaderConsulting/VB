VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hint..."
   ClientHeight    =   1275
   ClientLeft      =   3840
   ClientTop       =   4695
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1275
   ScaleWidth      =   6690
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblHelp 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub
