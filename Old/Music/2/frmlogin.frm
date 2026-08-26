VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to "
   ClientHeight    =   2280
   ClientLeft      =   3810
   ClientTop       =   3855
   ClientWidth     =   6750
   ClipControls    =   0   'False
   Icon            =   "FRMLOGIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2280
   ScaleWidth      =   6750
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox chkRemember 
      Alignment       =   1  'Right Justify
      Caption         =   "Remember user name"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgUser 
      Height          =   735
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgCBB 
      Height          =   735
      Left            =   120
      Picture         =   "FRMLOGIN.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblWelcome 
      Caption         =   "Type a user name and password to log on to "
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image imgKeys 
      Height          =   825
      Left            =   240
      Picture         =   "FRMLOGIN.frx":8074
      Top             =   120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblConfirm 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm password:"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblPass 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblUser 
      Alignment       =   1  'Right Justify
      Caption         =   "User name:"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    subEnd
End Sub
Private Sub cmdConfirm_Click()
    subConfirm
End Sub
Private Sub cmdOK_Click()
    Result = fctPassCheck(txtName, txtPassword)
    '
    '   codes for result:
    '   0 = NO - password incorrect
    '   1 = OK - password correct
    '   2 = OK - new user
    '   3 = NO - no password supplied
    '   4 = NO - user locked out
    '   5 = NO - no user name supplied
    
    If Result = 0 Then
        msg = "The User name or password you typed is incorrect"
        varStyle = vbOKOnly + vbCritical + vbDefaultButton1
        response = MsgBox(msg, varStyle, "Incorrect Login")
        txtName.SetFocus
        txtName.SelStart = 0
        txtName.SelLength = Len(txtName)
        txtPassword = ""
    End If
    If Result = 1 Then
        gUserName = UCase$(txtName)
        gUserPassword = UCase$(txtPassword)
        SubSaveLastName
        frmLicense.Show
        frmLogin.Hide
    End If
    If Result = 2 Then
        frmLogin.lblConfirm.Visible = 1
        frmLogin.txtConfirm.Visible = 1
        frmLogin.txtConfirm.SetFocus
        frmLogin.cmdOK.Visible = 0
        frmLogin.cmdConfirm.Visible = 1
        frmLogin.txtName.Locked = 1
        frmLogin.txtPassword.Locked = 1
        frmLogin.cmdConfirm.Default = 1
        Height = 2685
    End If
    If Result = 3 Then
        msg = "You must supply a password"
        varStyle = vbOKOnly + vbCritical + vbDefaultButton1
        response = MsgBox(msg, varStyle, "Incorrect Login")
        txtPassword.SetFocus
    End If
    If Result = 4 Then
        msg = "You are currently locked out of this program" & Chr$(10)
        msg = msg & "Contact your system administrator"
        varStyle = vbOKOnly + vbCritical + vbDefaultButton1
        response = MsgBox(msg, varStyle, "Locked Out")
        txtName.SetFocus
        txtName.SelStart = 0
        txtName.SelLength = Len(txtName)
        txtPassword = ""
    End If
    If Result = 5 Then
        msg = "Please type a user name and password to continue" & Chr$(10)
        msg = msg & "If you do not have a user name or password," & Chr$(10)
        msg = msg & "then please contact your System Administrator."
        varStyle = vbOKOnly + vbCritical + vbDefaultButton1
        response = MsgBox(msg, varStyle, "Incorrect Login")
        txtName.SetFocus
        txtName.SelStart = 0
        txtName.SelLength = Len(txtName)
        txtPassword = ""
    End If
    If gPassword = 1 Then
        If gUserLevel >= 98 Then
            frmStart.mnuModifyUser.Enabled = 1
            frmStart.mnuSetup.Enabled = 1
        End If
        If gUserLevel = 1 Then
            frmStart.mnuModifyUser.Enabled = 0
            frmStart.mnuSetup.Enabled = 0
        End If
    End If
End Sub
Private Sub Form_Activate()
    Height = 2130
    gLogo = GetSetting(gINIFile, "Setup", "Logo", 0)
    gRegistered = GetSetting(gINIFile, "Register", "Registered", 0)
    If gTimeCripple = 1 Then gTime = Timer
    If gRegistered = 1 Then
        gLogoPath = GetSetting(gINIFile, "Setup", "LogoPath", "")
        Code = GetSetting(gINIFile, "Register", "Code", "")
        Password = GetSetting(gINIFile, "Register", "Password", "")
        varPassword = Password
        'Code = txtCode
        Code = Mid$(Code, 3, Val(Left$(Code, 2)))
        varPassword = ""
        For varLoop = 1 To Len(Code) Step 2
            J = Mid$(Code, varLoop, 2)
            varPassword = varPassword + Chr$(Val("&H" + J))
        Next
        varPass = fctCrypt(gBusiness, varPassword)
        
        If varPassword <> varPass Then         ' CAUGHT!  ----- ******* INCORRECT PASSWORD IN INI FILE **********
            frmLogin.Hide
            msg = "Error in INI file." & Chr$(13)
            msg = msg & "Contact the Author or reinstall the program"
            Result = MsgBox(msg, vbCritical, "Critical Error")
            End
        End If
        If varPassword = "" Or Code = "" Then         ' CAUGHT!  ----- ******* INCORRECT PASSWORD IN INI FILE **********
            frmLogin.Hide
            msg = "Error in INI file." & Chr$(13)
            msg = msg & "Contact the Author or reinstall the program"
            Result = MsgBox(msg, vbCritical, "Critical Error")
            End
        End If
    End If
    If gLogo = 0 Then
        imgCBB.Visible = 1
        imgKeys.Visible = 0
        imgUser.Visible = 0
    End If
    If gLogo = 1 Then
        imgCBB.Visible = 0
        imgKeys.Visible = 1
        imgUser.Visible = 0
    End If
    If gLogo = 2 Then
        imgCBB.Visible = 0
        imgKeys.Visible = 0
        imgUser.Picture = LoadPicture(gLogoPath)
        imgUser.Visible = 1
    End If
    subLogin
End Sub
Private Sub Form_Load()
    Width = 6870
    Height = 2130
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
    Caption = Caption & gProgramName
    lblWelcome = lblWelcome & gProgramName
    On Error Resume Next
    gWindowsDir = Environ("windir")                     ' Get Windows Directory
    Rem varDir = Dir(gWindowsDir & "\WIN.COM")
    varDir = "WIN.COM"
    If varDir = "" Then
lblRetry:
        msg = gProgramName & " can't determine your Windows Directory." & Chr$(10) & Chr$(10)
        msg = msg & "Please enter the path where your copy of Windows is loaded."
        gWindowsDir = InputBox(msg, "Error in setup", "C:\WINDOWS")
        varDir = Dir(gWindowsDir & "\WIN.COM")
        If varDir = "" Then
            MsgBox ("Incorrect entry.  Please re-enter.")
            GoTo lblRetry
        End If
    End If
    
    DriveInfo = fctGetDriveInfo
    gSerial = Left$(DriveInfo, 9)
    gVolume = Right$(DriveInfo, Len(DriveInfo) - 9)
    
    Code = GetSetting(gINIFile, "Register", "Volume")
    If Volume = "" Then
        Volume = gVolume
    Else
        Volume = fctDecode(Code)
    End If
    If gVolume <> Volume Then
        msg = "Error in Volume Name." & Chr$(13)
        msg = msg & "Contact the Author or reinstall the program on your new Drive"
        Result = MsgBox(msg, vbCritical, "Critical Error")
        End
    End If
    Code = GetSetting(gINIFile, "Register", "Serial", gSerial)
    If Serial = "" Then
        Serial = gSerial
    Else
        Serial = fctDecode(Code)
    End If
    If gSerial <> Serial Then
        msg = "Error in Serial Number." & Chr$(13)
        msg = msg & "Contact the Author or reinstall the program on your new Drive."
        Result = MsgBox(msg, vbCritical, "Critical Error")
        End
    End If
End Sub
Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
