VERSION 5.00
Begin VB.Form FUserInfo 
   Caption         =   "NetUserInfo1 Test"
   ClientHeight    =   7995
   ClientLeft      =   1125
   ClientTop       =   1485
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7995
   ScaleWidth      =   7185
   Begin VB.ListBox lstGroups 
      Height          =   4545
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.OptionButton optType 
      Caption         =   "Level &3"
      Height          =   255
      Index           =   3
      Left            =   3900
      TabIndex        =   9
      Top             =   900
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.OptionButton optType 
      Caption         =   "Level &0"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   900
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2700
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   420
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   435
      Left            =   5280
      TabIndex        =   7
      Top             =   780
      Width           =   1635
   End
   Begin VB.OptionButton optType 
      Caption         =   "Level &2"
      Height          =   255
      Index           =   2
      Left            =   2700
      TabIndex        =   5
      Top             =   900
      Width           =   1035
   End
   Begin VB.OptionButton optType 
      Caption         =   "Level &1"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   900
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UserInfo"
      Default         =   -1  'True
      Height          =   435
      Left            =   5280
      TabIndex        =   6
      Top             =   180
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   420
      Width           =   2355
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Server:"
      Height          =   195
      Left            =   2700
      TabIndex        =   2
      Top             =   180
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "FUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private User As CNetUser

Private Const TIMEQ_FOREVER = -1&             '((unsigned long) -1L)
Private Const USER_MAXSTORAGE_UNLIMITED = -1& '((unsigned long) -1L)
Private Const USER_NO_LOGOFF = -1&            '((unsigned long) -1L)
Private Const UNITS_PER_DAY = 24
Private Const UNITS_PER_WEEK = UNITS_PER_DAY * 7

Private Const USER_PRIV_MASK = 3
Private Const USER_PRIV_GUEST = 0
Private Const USER_PRIV_USER = 1
Private Const USER_PRIV_ADMIN = 2

Private Sub Form_Load()
   On Error Resume Next
      Me.Font.Name = "Courier New"
   On Error GoTo 0
   Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
   
   Set User = New CNetUser
   Text1.Text = User.UserName
   Text2.Text = "\\" & User.Server
   Text2.Text = "\\CBDXAAA"
   Text1.Text = "PD09612"
End Sub

Private Sub Command1_Click()
   ShowUserInfo
   ShowGroupInfo
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub optType_Click(Index As Integer)
   ShowUserInfo
   ShowGroupInfo
End Sub

Private Sub ShowUserInfo()
   Dim i As Integer
   '
   ' Get info for User
   '
   User.Server = Trim(Text2.Text)
   User.UserName = Trim(Text1.Text)
   '
   ' Display Level 0 User Information
   '
   Me.Cls
   Me.CurrentY = Text1.Top + 3 * Text1.Height
   Me.Print "         Name: "; User.UserName
   If optType(0).Value Then Exit Sub
   '
   ' Display Level 1 User Information
   '
   Me.Print "      Password: "  '; User.Password
   Me.Print "   PasswordAge: "; Format(User.PasswordAge / 86400, "0.0") & " days"
   Select Case User.Privilege
      Case USER_PRIV_GUEST: Me.Print "     Privilege: Guest"
      Case USER_PRIV_USER: Me.Print "     Privilege: User"
      Case USER_PRIV_ADMIN: Me.Print "     Privilege: Administrator"
   End Select
   Me.Print "       HomeDir: "; User.HomeDir
   Me.Print "       Comment: "; User.Comment
   Me.Print "         Flags:"; User.Flags
   Me.Print "    ScriptPath: "; User.ScriptPath
   If optType(1).Value Then Exit Sub
   '
   ' Display Level 2 User Information
   '
   Me.Print "      AuthFlags: &h"; Hex(User.AuthFlags)
   Me.Print "       FullName: "; User.FullName
   Me.Print "    UserComment: "; User.UserComment
   Me.Print "          Parms: "; User.Parms
   Me.Print "   Workstations: "; User.Workstations
   Me.Print "      LastLogon: "; Format(User.LastLogonDate, "long date")
   Me.Print "     LastLogoff: "; Format(User.LastLogoffDate, "long date")
   If User.AcctExpires = TIMEQ_FOREVER Then
      Me.Print "    AcctExpires: Never"
   Else
      Me.Print "    AcctExpires: "; Format(User.AcctExpiresDate, "long date")
   End If
   If User.MaxStorage = USER_MAXSTORAGE_UNLIMITED Then
      Me.Print "     MaxStorage: Unlimited"
   Else
      Me.Print "     MaxStorage:"; User.MaxStorage
   End If
   Me.Print "   UnitsPerWeek:"; User.UnitsPerWeek
   Me.Print "     LogonHours: ";
   For i = 0 To 20
      Me.Print Right("0" & Hex(User.LogonHours(i)), 2);
   Next i
   Me.Print
   Me.Print "     BadPwCount:"; User.BadPasswordCount
   Me.Print "      NumLogons:"; User.NumLogons
   Me.Print "    LogonServer: "; User.LogonServer
   Me.Print "    CountryCode:"; User.CountryCode
   Me.Print "       CodePage:"; User.CodePage
   If optType(2).Value Then Exit Sub
   '
   ' Display Level 3 User Information
   '
   Me.Print "          UserID:"; User.UserID
   Me.Print "  PrimaryGroupID:"; User.PrimaryGroupID
   Me.Print "         Profile: "; User.Profile
   Me.Print "    HomeDirDrive: "; User.HomeDirDrive
   Me.Print " PasswordExpired: "; User.PasswordExpired
End Sub

Private Sub ShowGroupInfo()
   Dim i As Integer
   Dim grps As String
   
   Me.Print "          Groups:"; User.GroupCount;
   For i = 0 To User.GroupCount - 1
      Me.Print ", '"; User.Group(i); "'";
      lstGroups.AddItem User.Group(i)
   Next i
   Me.Print
   
   Me.Print "     LocalGroups:"; User.LocalGroupCount;
   For i = 0 To User.LocalGroupCount - 1
      Me.Print ", '"; User.LocalGroup(i); "'";
   Next i
   Me.Print
End Sub
