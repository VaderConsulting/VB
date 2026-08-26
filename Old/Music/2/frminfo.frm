VERSION 5.00
Begin VB.Form frmInfo 
   Caption         =   "Program Info"
   ClientHeight    =   6915
   ClientLeft      =   1140
   ClientTop       =   1530
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6915
   ScaleWidth      =   9600
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblCost 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   9375
   End
   Begin VB.Label lblTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    frmRegister.Show
    frmInfo.Hide
End Sub

Private Sub Form_Activate()
    lblTime = "This program was my major project over a 3 month period.  "
    lblTime = lblTime & "The background programming took a little over 200 hours alone." & Chr$(10)
    lblCost = "Cost: if you already own a previous version of " & gProgramName & " then you can "
    lblCost = lblCost & "buy this upgrade for A$" & gUpgrade & Chr$(10)
    lblCost = lblCost & "To buy the full priced product is A$" & gFullPrice
End Sub

Private Sub Form_Load()
    Width = gVGAWidth
    Height = gVGAHeight
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub
