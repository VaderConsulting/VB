VERSION 5.00
Begin VB.Form frmMe 
   Caption         =   "Current Settings"
   ClientHeight    =   4080
   ClientLeft      =   3990
   ClientTop       =   6585
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4080
   ScaleWidth      =   6570
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change password"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtLevel 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblInvisible 
      Caption         =   "<----- Invisible Controls"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "User Level"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblLogin 
      Caption         =   "Login name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmStart.Show
    frmMe.Hide
End Sub

Private Sub cmdChange_Click()
    frmMe.cmdOK.Enabled = 0
    frmChange.Show
End Sub

Private Sub cmdOK_Click()
    If gPassword = 1 Then
        Result = fctSaveSettings(txtLogin, txtPassword, txtLevel, gUserLock)
    End If
    frmStart.Show
    frmMe.Hide
End Sub

Private Sub Form_Activate()
    txtLogin = gUserName
    txtLevel = gUserLevel
    txtPassword = gUserPassword
End Sub

Private Sub Form_Load()
    Width = 6690
    Height = 4485
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub

