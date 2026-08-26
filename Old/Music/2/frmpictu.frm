VERSION 5.00
Begin VB.Form frmPicture 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4065
   ClientLeft      =   3255
   ClientTop       =   2775
   ClientWidth     =   4890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4065
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   285
      Left            =   4230
      TabIndex        =   0
      Top             =   3690
      Width           =   555
   End
   Begin VB.Image imgPicture 
      Height          =   4065
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    frmPicture.Hide
End Sub

Private Sub Form_Activate()
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
    If frmStart.txtPicture.Text <> "" Then imgPicture = LoadPicture(frmStart.txtPicture.Text)
End Sub

