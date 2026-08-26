VERSION 5.00
Begin VB.Form frmChange 
   Caption         =   "Change Password"
   ClientHeight    =   1545
   ClientLeft      =   5085
   ClientTop       =   2250
   ClientWidth     =   5370
   Icon            =   "FRMCHNGE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1545
   ScaleWidth      =   5370
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtOld 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblConfirm 
      Caption         =   "Confirm new password:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblNew 
      Caption         =   "New password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblOld 
      Caption         =   "Old password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    frmMe.cmdOK.Enabled = 1
    frmChange.Hide
End Sub

Private Sub cmdOK_Click()
    txtOld = UCase$(txtOld)
    txtNew = UCase$(txtNew)
    txtConfirm = UCase$(txtConfirm)
    Result = 0
    If txtOld = gUserPassword Then
        If txtNew = txtConfirm Then
            gUserPassword = txtNew
            frmMe.txtPassword = txtNew
            frmMe.Refresh
            Result = 1
        Else
            msg = "The new box is different to the confirm box"
            varStyle = vbOKOnly + vbCritical + vbDefaultButton1
            response = MsgBox(msg, varStyle, "Incorrect passwords")
            txtConfirm = ""
            txtNew = ""
            txtNew.SetFocus
        End If
    Else
        msg = "The old Password is incorrect"
        varStyle = vbOKOnly + vbCritical + vbDefaultButton1
        response = MsgBox(msg, varStyle, "Incorrect password")
        txtOld = ""
        txtNew = ""
        txtConfirm = ""
        txtOld.SetFocus
    End If
    If Result = 1 Then
        frmMe.cmdOK.Enabled = 1
        frmChange.Hide
    End If
End Sub

Private Sub Form_Activate()
    txtOld = ""
    txtNew = ""
    txtConfirm = ""
    txtOld.SetFocus
End Sub

Private Sub Form_Load()
    Width = 5490
    Height = 1950
    'these sizes equate to a full 640x480 screen
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub


Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub


Private Sub txtNew_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub


Private Sub txtOld_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub


