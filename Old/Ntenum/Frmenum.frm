VERSION 5.00
Begin VB.Form FrmEnum 
   Caption         =   "Enumeration"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstEnum 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
   Begin VB.CommandButton CmdGroups 
      Caption         =   "Enum groups"
      Height          =   495
      Left            =   2573
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TxtDomain 
      Height          =   285
      Left            =   1613
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton CmdUsers 
      Caption         =   "Enum users"
      Height          =   495
      Left            =   893
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label LblDomain 
      Caption         =   "Domain:"
      Height          =   255
      Left            =   653
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FrmEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variables must be declared
Option Explicit
Private UsersShown As Boolean

Function IsFlag(Flags As Long, TestFlag As Long) As Boolean

    ' Determines if testflag is set in flag
    
     IsFlag = ((Flags And TestFlag) = TestFlag)
     
End Function

Private Sub CmdGroups_Click()
  Dim Ok As Boolean
  Dim Size As Long
  Dim Data() As String
  Dim i As Long
  UsersShown = False
  LstEnum.Clear
  Ok = QuickEnumerate(TxtDomain, 3, Size, Data)
  If Ok Then
   For i = 1 To Size
     LstEnum.AddItem (Data(i))
   Next i
  Else
    MsgBox ("Error when trying to enumerate groups")
  End If
End Sub

Private Sub CmdUsers_Click()
  Dim Ok As Boolean
  Dim Size As Long
  Dim Data() As String
  Dim i As Long
  LstEnum.Clear
  Ok = QuickEnumerate(TxtDomain, 1, Size, Data)
  If Ok Then
   For i = 1 To Size
     LstEnum.AddItem (Data(i))
   Next i
   UsersShown = True
  Else
    MsgBox ("Error when trying to enumerate Users")
    UsersShown = False
  End If
End Sub

Private Sub Form_Load()
  UsersShown = False
End Sub

Private Sub LstEnum_Click()
Dim Ok As Boolean
Dim Info As User_Info_3
Dim Groups() As String
Dim Size As Long
Dim i As Long
  If UsersShown Then
    If LstEnum.ListIndex <> -1 Then
      Ok = GetUserInfo(TxtDomain, LstEnum.Text, Info)
      If Not Ok Then
        MsgBox ("Unable to get information about " & LstEnum.Text)
      Else
        FrmUserInfo.TxtUserName.Text = Info.name
        FrmUserInfo.TxtFullName.Text = Info.full_name
        FrmUserInfo.TxtDescription.Text = Info.comment
        FrmUserInfo.TxtPrivilege.Text = Info.priv
        If IsFlag(Info.Flags, UF_ACCOUNTDISABLE) Then
          FrmUserInfo.TxtAcctActive.Text = "No"
        Else
          FrmUserInfo.TxtAcctActive.Text = "Yes"
        End If
        If IsFlag(Info.Flags, UF_LOCKOUT) Then
          FrmUserInfo.TxtAcctLockedOut.Text = "Yes"
        Else
          FrmUserInfo.TxtAcctLockedOut.Text = "No"
        End If
        FrmUserInfo.TxtAcctExpires.Text = Info.acct_expires
        FrmUserInfo.TxtLastLogon.Text = Info.last_logon
        FrmUserInfo.TxtNumLogons.Text = Info.num_logons
        FrmUserInfo.TxtPasswordLastSet.Text = Info.password_last_set
        FrmUserInfo.TxtPasswordExpired.Text = Info.password_expired
        If IsFlag(Info.Flags, UF_PASSWD_CANT_CHANGE) Then
          FrmUserInfo.TxtPasswordChangeable.Text = "No"
        Else
          FrmUserInfo.TxtPasswordChangeable.Text = "Yes"
        End If
        If IsFlag(Info.Flags, UF_DONT_EXPIRE_PASSWD) Then
          FrmUserInfo.TxtPasswordNeverExpires.Text = "No"
        Else
          FrmUserInfo.TxtPasswordNeverExpires.Text = "Yes"
        End If
        FrmUserInfo.TxtBadPwCount.Text = Info.bad_pw_count
        FrmUserInfo.TxtUserProfilePath.Text = Info.profile
        FrmUserInfo.TxtLoginScriptName.Text = Info.script_path
        FrmUserInfo.TxtHomedir.Text = Info.home_dir
        FrmUserInfo.TxtHomedirDrive.Text = Info.home_dir_drive
        FrmUserInfo.TxtWorkstationsAllowed.Text = Info.workstations_allowed
      
        FrmUserInfo.CmbGlobalGroups.Clear
        FrmUserInfo.CmbGlobalGroups.Text = "Press button to view memberships"
        Ok = EnumerateUsersGroupMembships(TxtDomain.Text, LstEnum.Text, Size, Groups)
        If Not Ok Then
          MsgBox ("Unable to retrieve information about Global Group Memberships for " & LstEnum.Text)
        Else
          For i = 1 To Size
            FrmUserInfo.CmbGlobalGroups.AddItem (Groups(i))
          Next i
          FrmUserInfo.Show vbModal, Me
        End If
      End If
    End If
  End If
End Sub
