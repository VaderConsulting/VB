VERSION 5.00
Begin VB.Form FRM_Secure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon"
   ClientHeight    =   1500
   ClientLeft      =   3675
   ClientTop       =   3765
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "FRM_Secure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1500
   ScaleWidth      =   4800
   Begin VB.CommandButton CMD_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton CMD_OK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox TXT_Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox TXT_Name 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "&Password"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "&Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu MNU_Changewg 
      Caption         =   "&Change Workgroup File"
      Index           =   100
   End
End
Attribute VB_Name = "FRM_Secure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_Cancel_Click()
    Unload FRM_Secure
End Sub

Private Sub CMD_OK_Click()
' this is where we check to see if the person logging in should be doing that!
' Robert May 1-30-98
    On Error GoTo errortrip
    Dim dexwks As Workspace
    Dim degroup As Group
    Dim priv As Boolean
    Dim groups(1) As String
    priv = False
    If TXT_Name = "" Then
        Call MsgBox("You must fill out the username field!", vbOKOnly + vbCritical, "Missing Field")
        Exit Sub
    End If
    groups(0) = "Domain Admins"
    groups(1) = "Management"
    priv = IsMember(FRM_Secure.TXT_Name, groups)
    DBEngine.SystemDB = GetSetting("EnumUsers", "Workgroup", "Location", "\\dbsserver\access\starwest.mdw")
    DBEngine.DefaultUser = TXT_Name
    DBEngine.DefaultPassword = TXT_Password
    Set dexwks = DBEngine.CreateWorkspace(TXT_Name & TXT_Password, TXT_Name, TXT_Password)
    For Each degroup In dexwks.Users(TXT_Name).groups
        If degroup.Name = "Management" Or degroup.Name = "Admins" Then
            priv = True
        End If
    Next
    If priv = False Then
        Call MsgBox("You do not have the proper privleges to run this program.", vbOKOnly + vbCritical, "Insufficient Privledges")
        Unload FRM_Secure
        Unload FRM_Users
        Exit Sub
    End If
    dexwks.Close
    FRM_Users.Show
    FRM_Secure.Hide
    Exit Sub
errortrip:
    Select Case Err.Number
        Case 3029
            Call MsgBox("Invalid account name or password.", vbOKOnly + vbExclamation, "Logon")
            Resume ending:
            Exit Sub
        Case Else
            Call MsgBox("Unable to log in at this time, please try again later.", vbOKOnly + vbExclamation, "Logon")
            Unload FRM_Secure
            Unload FRM_Users
            Exit Sub
    End Select
ending:
End Sub

Private Sub Form_Load()
    FRM_Secure.Move (Screen.Height - FRM_Secure.Height) / 2, (Screen.Width - FRM_Secure.Width) / 2

    FRM_Secure.Show
    
End Sub


Private Sub MNU_Changewg_Click(Index As Integer)
    ' this is your microsoft access workgroup location
    SaveSetting "EnumUsers", "Workgroup", "Location", InputBox("Enter New WorkGroup Location", "WorkGroup Location", GetSetting("Dex", "Workgroup", "Location", "\\dbsserver\access\starwest.mdw"))
End Sub


