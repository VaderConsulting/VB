VERSION 5.00
Begin VB.Form frmAdmin 
   Caption         =   "Admin"
   ClientHeight    =   4080
   ClientLeft      =   4245
   ClientTop       =   3450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4080
   ScaleWidth      =   6570
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.CheckBox chkLocked 
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtLevel 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.ListBox lstUsers 
      Height          =   3375
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblUsers 
      Alignment       =   2  'Center
      Caption         =   "Users:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblNewPassword 
      Alignment       =   1  'Right Justify
      Caption         =   "New password:"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblLocked 
      Alignment       =   1  'Right Justify
      Caption         =   "Account locked:"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblLevel 
      Alignment       =   1  'Right Justify
      Caption         =   "Level:"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      Caption         =   "Current password:"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    frmStart.Show
    frmAdmin.Hide
End Sub

Private Sub cmdDelete_Click()
    varUsername = frmAdmin.lstUsers.List(frmAdmin.lstUsers.ListIndex)
    DeleteSetting gINIFile, "Passwords", varUsername
    DeleteSetting gINIFile, "Locks", varUsername
    DeleteSetting gINIFile, "Levels", varUsername
    subRefresh
End Sub

Private Sub cmdNew_Click()
    cmdUpdate.Visible = 0
    cmdRefresh.Visible = 1
    lblPassword = "User name:"
    txtPassword = ""
    txtNewPassword = ""
    txtPassword.SetFocus
    frmAdmin.txtLevel = 1
    chkLocked.Value = 0
    lstUsers.Enabled = 0
    'subGetSettings
End Sub


Private Sub cmdOK_Click()
    If txtNewPassword <> "" Then gUserPassword = txtNewPassword
    frmStart.Show
    frmAdmin.Hide
End Sub

Private Sub cmdRefresh_Click()
    cmdRefresh.Visible = 0
    cmdUpdate.Visible = 1
    If chkLocked = 0 Then
        UserLock = 0
    Else
        UserLock = 1
    End If
    Result = fctSaveSettings(txtPassword, txtNewPassword, Val(txtLevel), UserLock)
    subRefresh
End Sub

Private Sub cmdUpdate_Click()
    UserLevel = Val(txtLevel)
    If chkLocked = 0 Then
        UserLock = 0
    Else
        UserLock = 1
    End If
    If txtNewPassword <> "" Then
        UserPassword = txtNewPassword
    Else
        UserPassword = txtPassword
    End If
    Result = fctSaveSettings(lstUsers.List(lstUsers.ListIndex), UserPassword, UserLevel, UserLock)
    If gUserName = lstUsers.List(lstUsers.ListIndex) Then gUserPassword = UserPassword  ' because user modified self
    txtPassword = txtNewPassword
    txtNewPassword = ""
End Sub

Private Sub Form_Activate()
    subRefresh
End Sub

Private Sub Form_Load()
    Width = 6690
    Height = 4485
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub


Private Sub lstUsers_Click()
    'this sub has been tested and is OK
    varName = lstUsers.List(lstUsers.ListIndex)
    varEncrypted = GetSetting(gINIFile, "Passwords", varName)
    varPass = fctDecode(varEncrypted)
    varLevel = GetSetting(gINIFile, "Levels", varName)
    varLock = GetSetting(gINIFile, "Locks", varName)
    txtPassword = varPass
    txtLevel = varLevel
    cmdUpdate.Enabled = 1
    If Val(varLock) = 1 Then
        chkLocked.Value = 1
    Else
        chkLocked.Value = 0
    End If
End Sub


Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub


