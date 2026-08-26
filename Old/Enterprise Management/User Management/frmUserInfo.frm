VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UserInfo"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tabUserInfo 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Basic"
      TabPicture(0)   =   "frmUserInfo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUserInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUserInfo(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUserInfo(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblUserInfo(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblUserInfo(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblUserInfo(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblUserInfo(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtUserInfo(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtUserInfo(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtUserInfo(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtUserInfo(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtUserInfo(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lstGroups"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtUserInfo(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "User Flags"
      TabPicture(1)   =   "frmUserInfo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkUserInfo(16)"
      Tab(1).Control(1)=   "chkUserInfo(15)"
      Tab(1).Control(2)=   "chkUserInfo(14)"
      Tab(1).Control(3)=   "chkUserInfo(13)"
      Tab(1).Control(4)=   "chkUserInfo(12)"
      Tab(1).Control(5)=   "chkUserInfo(11)"
      Tab(1).Control(6)=   "chkUserInfo(10)"
      Tab(1).Control(7)=   "chkUserInfo(9)"
      Tab(1).Control(8)=   "chkUserInfo(8)"
      Tab(1).Control(9)=   "chkUserInfo(7)"
      Tab(1).Control(10)=   "chkUserInfo(6)"
      Tab(1).Control(11)=   "chkUserInfo(5)"
      Tab(1).Control(12)=   "chkUserInfo(4)"
      Tab(1).Control(13)=   "chkUserInfo(3)"
      Tab(1).Control(14)=   "chkUserInfo(2)"
      Tab(1).Control(15)=   "chkUserInfo(1)"
      Tab(1).Control(16)=   "chkUserInfo(0)"
      Tab(1).ControlCount=   17
      Begin VB.TextBox txtUserInfo 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2280
         Width           =   3255
      End
      Begin VB.ListBox lstGroups 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   3000
         Width           =   9015
      End
      Begin VB.TextBox txtUserInfo 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox txtUserInfo 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtUserInfo 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtUserInfo 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtUserInfo 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   480
         Width           =   3255
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "The logon script will be executed."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   18
         Top             =   420
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "The user's account is disabled."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   17
         Top             =   660
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "The home directory is required."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   16
         Top             =   900
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "The account is currently locked out."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   15
         Top             =   1140
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "No password is required."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   14
         Top             =   1380
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "The user cannot change the password."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   13
         Top             =   1620
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "The user can send an encrypted password."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   12
         Top             =   1860
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "This is an account for users whose primary account is in another domain."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   11
         Top             =   2100
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "This is a default account type that represents a typical user."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   10
         Top             =   2340
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "This is a permit to trust account for a system domain that trusts other domains."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   9
         Top             =   2580
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "This is a computer account."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   8
         Top             =   2820
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "This is a computer account for a system backup domain controller that is a member of this domain."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   7
         Top             =   3060
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "Represents the password, which should never expire on the account."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -74880
         TabIndex        =   6
         Top             =   3300
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "This is an MNS logon account."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -74880
         TabIndex        =   5
         Top             =   3540
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "When set, this flag will force the user to log on using smart card."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   -74880
         TabIndex        =   4
         Top             =   3780
         Width           =   6855
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   "When set, the service account (user or computer account), under which a service runs, is trusted for Kerberos delegation."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   -74880
         TabIndex        =   3
         Top             =   4020
         Width           =   9015
      End
      Begin VB.CheckBox chkUserInfo 
         Caption         =   $"frmUserInfo.frx":0038
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   -74880
         TabIndex        =   2
         Top             =   4260
         Width           =   9135
      End
      Begin VB.Label lblUserInfo 
         Caption         =   "Logon Script"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblUserInfo 
         Alignment       =   2  'Center
         Caption         =   "Group Membership"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   9015
      End
      Begin VB.Label lblUserInfo 
         Caption         =   "Profile Path"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblUserInfo 
         Caption         =   "Home Path"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblUserInfo 
         Caption         =   "Home Drive"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblUserInfo 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblUserInfo 
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   5760
      Width           =   975
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormLoaded As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim UserFlag As Long
    Dim p As Integer
    Dim strTemp As String
    Dim UserInfo(16) As Boolean
    Dim User As IADsUser
    Dim Group As IADsGroup
    
    Screen.MousePointer = vbHourglass
    Me.Show
    Me.Refresh
    ' Get Username
    strTemp = frmMain.lstUsers.List(frmMain.lstUsers.ListIndex)
    p = InStr(1, strTemp, vbTab) - 1
    UserName = Left(strTemp, p)
    
    Set User = GetObject("WinNT://" & gstrDomain & "/" & UserName & ",user")
    
    txtUserInfo(0) = User.Name
    txtUserInfo(1) = User.Description
    txtUserInfo(2) = User.homedirdrive
    txtUserInfo(3) = User.HomeDirectory
    txtUserInfo(4) = User.Profile
    txtUserInfo(5) = User.LoginScript
    Me.Refresh
    ' Set tab positions
    Call SetListTabStops(lstGroups.hWnd, 28)
    
    'Get Groups
    For Each Group In User.Groups
        lstGroups.AddItem Group.Name & vbTab & Group.Description
        DoEvents
    Next
    Me.Refresh
    UserFlag = User.UserFlags
    Erase UserInfo
    ' Display Userflag info (on tab 2)
    If UserFlag And ADS_UF_SCRIPT Then UserInfo(0) = True
    If UserFlag And ADS_UF_ACCOUNTDISABLE Then UserInfo(1) = True
    If UserFlag And ADS_UF_HOMEDIR_REQUIRED Then UserInfo(2) = True
    If UserFlag And ADS_UF_LOCKOUT Then UserInfo(3) = True
    If UserFlag And ADS_UF_PASSWD_NOTREQD Then UserInfo(4) = True
    If UserFlag And ADS_UF_PASSWD_CANT_CHANGE Then UserInfo(5) = True
    If UserFlag And ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED Then UserInfo(6) = True
    If UserFlag And ADS_UF_TEMP_DUPLICATE_ACCOUNT Then UserInfo(7) = True
    If UserFlag And ADS_UF_NORMAL_ACCOUNT Then UserInfo(8) = True
    If UserFlag And ADS_UF_INTERDOMAIN_TRUST_ACCOUNT Then UserInfo(9) = True
    If UserFlag And ADS_UF_WORKSTATION_TRUST_ACCOUNT Then UserInfo(10) = True
    If UserFlag And ADS_UF_SERVER_TRUST_ACCOUNT Then UserInfo(11) = True
    If UserFlag And ADS_UF_DONT_EXPIRE_PASSWD Then UserInfo(12) = True
    If UserFlag And ADS_UF_MNS_LOGON_ACCOUNT Then UserInfo(13) = True
    If UserFlag And ADS_UF_SMARTCARD_REQUIRED Then UserInfo(14) = True
    If UserFlag And ADS_UF_TRUSTED_FOR_DELEGATION Then UserInfo(15) = True
    If UserFlag And ADS_UF_NOT_DELEGATED Then UserInfo(16) = True
    
    For lp = 0 To UBound(UserInfo())
        If UserInfo(lp) Then chkUserInfo(lp).Value = vbChecked
    Next lp
    Screen.MousePointer = vbDefault
End Sub

Private Sub lstGroups_DblClick()
    Dim Groupname As String
    Dim p As Integer
    Dim strTemp As String
    strTemp = lstGroups.List(lstGroups.ListIndex)
    p = InStr(1, strTemp, vbTab) - 1
    Groupname = Left(strTemp, p)
    frmGroupInfo.lblGroupname = Groupname
    frmGroupInfo.Show vbModal
End Sub

