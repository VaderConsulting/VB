VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeTemplate 
      Height          =   3135
      Left            =   6240
      TabIndex        =   27
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   6120
      Width           =   735
   End
   Begin VB.Frame fmeNoTemplate 
      Height          =   5535
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   6015
      Begin VB.CommandButton cmdCheck 
         Caption         =   "Check"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdRemoveGroup 
         Caption         =   "<"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton cmdAddGroup 
         Caption         =   ">"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   4080
         Width           =   495
      End
      Begin VB.ListBox lstUserGroups 
         Height          =   2010
         Left            =   3360
         TabIndex        =   13
         Top             =   3360
         Width           =   2535
      End
      Begin VB.ListBox lstGroups 
         Height          =   2010
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   3360
         Width           =   2535
      End
      Begin VB.ComboBox cmbSiteloc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Text            =   "SELECT A SITELOC"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optDuty 
         Caption         =   "OIC"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   9
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton optDuty 
         Caption         =   "Detective"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton optDuty 
         Caption         =   "General"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtProfileDir 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtHomeDir 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtLogonScript 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "begin.bat"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtFullname 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   5880
         Y1              =   2895
         Y2              =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   5880
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label lblGroups 
         Alignment       =   2  'Center
         Caption         =   "Group Membership"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   5775
      End
      Begin VB.Label lblSiteloc 
         Caption         =   "Siteloc"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label lblProfileDir 
         Caption         =   "Profile Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2055
         Width           =   1575
      End
      Begin VB.Label lblHomeDir 
         Caption         =   "Home Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1695
         Width           =   1575
      End
      Begin VB.Label lblLogonScript 
         Caption         =   "Logon Script"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label lblFullname 
         Caption         =   "Fullname"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   615
         Width           =   855
      End
      Begin VB.Label lblUsername 
         Caption         =   "Username"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtDomain 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "WARNBRO"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   6120
      Width           =   4335
   End
   Begin VB.Label lblDomain 
      Caption         =   "Domain"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DomainName As String, DSN As String
Dim Conn As ADODB.Connection, RS As ADODB.Recordset
Dim Domain As IADsDomain, User As IADsUser
Dim SQL As String

Private Sub Form_Load()
  Me.Show
  Me.Refresh
  DomainName = "WARNBRO"
  DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=SERVER"
  Set Conn = CreateObject("ADODB.Connection")
  Set RS = CreateObject("ADODB.Recordset")
  Conn.Open DSN
  SQL = "SELECT * FROM tblSiteloc ORDER by Siteloc"
  lblStatus = "Retrieving siteloc's"
  lblStatus.Refresh
  RS.Open SQL, Conn, adOpenStatic
  Do Until RS.EOF
    cmbSiteloc.AddItem RS("Siteloc")
    RS.MoveNext
  Loop
  RS.Close
  Conn.Close
  lblStatus = "Retrieving Groupnames from " & DomainName
  lblStatus.Refresh
  Set Domain = GetObject("WinNT://" & DomainName)
  Domain.Filter = Array("Group")
  For Each Group In Domain
    lstGroups.AddItem Group.Name
  Next
  lblStatus = "Idle"
  lblStatus.Refresh
End Sub

Private Sub cmbSiteloc_LostFocus()
  If cmbSiteloc.Text = "SELECT A SITELOC" Then Exit Sub
  Conn.Open DSN
  SQL = "SELECT * FROM tblSiteloc WHERE Siteloc='" & cmbSiteloc.Text & "' ORDER by Siteloc"
  lblStatus = "Retrieving siteserver for " & cmbSiteloc.Text
  lblStatus.Refresh
  RS.Open SQL, Conn, adOpenStatic
  txtHomeDir = "\\" & RS("Siteserver") & "\" & txtUsername & "$"
  txtProfileDir = "\\" & RS("Siteserver") & "\Profiles\" & txtUsername
  RS.Close
  Conn.Close
  lblStatus = "Idle"
  lblStatus.Refresh
  
End Sub

Private Sub txtUsername_LostFocus()
  If txtUsername <> "" Then
    cmdCheck.Enabled = True
    txtFullname.Enabled = True
    txtLogonScript.Enabled = True
    cmbSiteloc.Enabled = True
    txtHomeDir.Enabled = True
    txtProfileDir.Enabled = True
    ' The following code is required to ensure that using the [TAB] key
    ' moves the focus to the correct control.
    If frmMain.ActiveControl.Name = "optDuty" Then
      cmdCheck.SetFocus
    End If
  Else
    cmdCheck.Enabled = False
    txtFullname.Enabled = False
    txtLogonScript.Enabled = False
    cmbSiteloc.Enabled = False
    txtHomeDir.Enabled = False
    txtProfileDir.Enabled = False
  End If
End Sub

Private Sub cmdAddGroup_Click()
  If lstGroups.ListIndex = -1 Then Exit Sub
  lstUserGroups.AddItem lstGroups.List(lstGroups.ListIndex)
  lstGroups.RemoveItem lstGroups.ListIndex
End Sub

Private Sub cmdCancel_Click()
  Set Domain = Nothing
  Set User = Nothing
  Set RS = Nothing
  Set Conn = Nothing
  Unload Me
  End
End Sub

Private Sub cmdRemoveGroup_Click()
  If lstUserGroups.ListIndex = -1 Then Exit Sub
  lstGroups.AddItem lstUserGroups.List(lstUserGroups.ListIndex)
  lstUserGroups.RemoveItem lstUserGroups.ListIndex
End Sub


