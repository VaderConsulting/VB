VERSION 5.00
Begin VB.Form FRM_Users 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administer Users"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5970
   Icon            =   "FRM_Users.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LST_Temp 
      Height          =   255
      Left            =   3990
      TabIndex        =   21
      Top             =   105
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox TXT_Verify 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   2940
      Width           =   2325
   End
   Begin VB.TextBox TXT_Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   105
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2940
      Width           =   1695
   End
   Begin VB.CommandButton CMD_CServer 
      Caption         =   "Change Server"
      Height          =   330
      Left            =   4620
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton CMD_DropG 
      Caption         =   "-->"
      Height          =   255
      Left            =   1950
      TabIndex        =   3
      Top             =   1470
      Width           =   375
   End
   Begin VB.CommandButton CMD_AddG 
      Caption         =   "<--"
      Height          =   255
      Left            =   1950
      TabIndex        =   2
      Top             =   885
      Width           =   375
   End
   Begin VB.TextBox TXT_Last 
      Height          =   285
      Left            =   1830
      TabIndex        =   6
      Top             =   2415
      Width           =   2415
   End
   Begin VB.TextBox TXT_First 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2415
      Width           =   1725
   End
   Begin VB.ListBox LST_PosGroups 
      Height          =   1425
      Left            =   2430
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   645
      Width           =   1815
   End
   Begin VB.CommandButton CMD_ShowAll 
      Caption         =   "Reset"
      Height          =   420
      Left            =   1470
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ListBox LST_Groups 
      Height          =   1425
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   645
      Width           =   1815
   End
   Begin VB.CommandButton CMD_DelUser 
      Caption         =   "&Delete User"
      Height          =   420
      Left            =   2910
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton CMD_AddUser 
      Caption         =   "&Add User"
      Height          =   420
      Left            =   30
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ListBox LST_Name 
      Height          =   3180
      Left            =   4380
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   630
      Width           =   1545
   End
   Begin VB.Label Label7 
      Caption         =   "Verify"
      Height          =   225
      Left            =   1890
      TabIndex        =   20
      Top             =   2730
      Width           =   2115
   End
   Begin VB.Label Label6 
      Caption         =   "Password"
      Height          =   225
      Left            =   105
      TabIndex        =   19
      Top             =   2730
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Users"
      Height          =   225
      Left            =   4410
      TabIndex        =   18
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   1830
      TabIndex        =   17
      Top             =   2205
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "First Name/User Name"
      Height          =   255
      Left            =   105
      TabIndex        =   16
      Top             =   2205
      Width           =   1665
   End
   Begin VB.Label Label2 
      Caption         =   "Available Groups"
      Height          =   255
      Left            =   2415
      TabIndex        =   15
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Assigned Groups"
      Height          =   255
      Left            =   30
      TabIndex        =   14
      Top             =   405
      Width           =   1575
   End
   Begin VB.Label LAB_DC 
      Caption         =   "Server"
      Height          =   255
      Left            =   105
      TabIndex        =   12
      Top             =   0
      Width           =   3375
   End
   Begin VB.Menu MNU_File 
      Caption         =   "&File"
      Index           =   100
      Begin VB.Menu MNU_Exit 
         Caption         =   "E&xit"
         Index           =   102
      End
   End
   Begin VB.Menu MNU_Help 
      Caption         =   "Help"
      Index           =   200
      Begin VB.Menu MNU_about 
         Caption         =   "About"
         Index           =   201
      End
   End
End
Attribute VB_Name = "FRM_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMD_AddG_Click()
' Oh the joy of this mess!  THis will add a person to a group
' Robert May 1-30-98
    Dim found As Boolean
    Dim dewks As Workspace
    Dim mygroup As Group
    found = False
    For i = 0 To LST_Groups.ListCount - 1
        If LST_PosGroups = LST_Groups.List(i) Then
            found = True
        End If
    Next i
    If LST_PosGroups.ListIndex <> -1 And found = False Then
        If LST_Name.ListIndex <> -1 Then
            With FRM_Secure
                Set dewks = DBEngine.CreateWorkspace(.TXT_Name & Time, .TXT_Name, .TXT_Password)
            End With
            Set mygroup = dewks.Users(LST_Name).CreateGroup(LST_PosGroups)
            dewks.Users(LST_Name).groups.Append mygroup
            AddUserToGroup GLBServer, LST_PosGroups, LST_Name
        End If
        LST_Groups.AddItem LST_PosGroups
    End If
End Sub

Private Sub CMD_AddUser_Click()
' this is the sweet heart of the program.  NOTE:  I am adding or updating a
' record in our employee information database.  You'll either need to change
' your table names, or elimiate the code all together.
' Robert May 1-30-98
    On Error Resume Next
    Dim dewks As Workspace
    Dim deuser As User
    Dim degroup As Group
    Dim dedb As Database
    Dim derecordset As Recordset
    Dim GoError As Boolean
    GoError = False
    FRM_Users.MousePointer = vbHourglass
' this is all validation crap--too bad we have to waste our time here!
    If TXT_First <> "" Then
        If Len(TXT_First) + Len(TXT_Last) < 4 Then
                Call MsgBox("The combined length of the First Name and Last name fields must be at least 4 characters in length.", vbOKOnly + vbExclamation, "Not Enough Info")
                FRM_Users.MousePointer = vbDefault
                Exit Sub
        End If
        If TXT_Last = "" Then TXT_Last = " "
        If TXT_Password = "" Or TXT_Verify = "" Then
                Call MsgBox("You must enter a password into both the password box and the verify box.", vbOKOnly + vbExclamation, "Bad Password")
                FRM_Users.MousePointer = vbDefault
                Exit Sub
        End If
        If Len(TXT_Password) < 6 Then
            Call MsgBox("Your password must be at least 6 characters or numbers.", vbOKOnly + vbExclamation, "Bad Password")
            FRM_Users.MousePointer = vbDefault
            Exit Sub
        End If
        For i = 0 To LST_Name.ListCount - 1
            If UCase$(LST_Name.List(i)) = UCase$(TXT_First) Then
                Call MsgBox("User already exists.  Please select another first name/username", vbOKOnly + vbExclamation, "User Exists")
                FRM_Users.MousePointer = vbDefault
                Exit Sub
            End If
        Next i
        If StrComp(TXT_Password, TXT_Verify) <> 0 Then
            Call MsgBox("The Password box and the verify box do not match.  Remember, passwords are case sensetive.  Please re-type the passwords.", vbOKOnly + vbExclamation, "Bad Passwords")
            FRM_Users.MousePointer = vbDefault
            Exit Sub
        End If
'create the user on the NT domain
        res = AddUser(GLBServer, TXT_First, TXT_Password)
        If res <> 0 Then
            Call MsgBox("An error occured adding this user.  User not added.  Result code " & res, vbOKOnly + vbExclamation, "User not added")
            Exit Sub
        End If
' open the access workgroup file and add the user
        With FRM_Secure
            Set dewks = DBEngine.CreateWorkspace(.TXT_Name & Date, .TXT_Name, .TXT_Password)
        End With
        Set deuser = dewks.CreateUser(TXT_First, Left$(TXT_First & TXT_Last, 20), TXT_Password)
        dewks.Users.Append deuser
' this is the database code.  If you don't want to use a database, you'll need
' to get rid of this
        Set dedb = dewks.OpenDatabase(GetSetting("EnumUsers", "Settings", "DBLOC", "\\dbsserver\access\starwest.mde"))
        Set derecordset = dedb.OpenRecordset("Select [Employee Information].* from [Employee Information] where [First Name]='" & TXT_First & "'", dbOpenDynaset)
        derecordset.MoveLast
        If Err.Number = 3021 Then GoError = True
        derecordset.MoveFirst
        If GoError = False Then
            derecordset.Edit
            derecordset.Fields("Active") = True
            derecordset.Fields("Password") = TXT_Password
            derecordset.Update
        Else
            derecordset.AddNew
            derecordset.Fields("First Name") = TXT_First
            derecordset.Fields("Last Name") = TXT_Last
            derecordset.Fields("Password") = TXT_Password
            derecordset.Update
        End If
' that's the end of the db code.  Next we add groups.
        For i = 0 To LST_Groups.ListCount - 1
            If LST_Groups.List(i) = "Domain Users" Then
                Set degroup = deuser.CreateGroup("Users")
            Else
                Set degroup = deuser.CreateGroup(LST_Groups.List(i))
            End If
            dewks.Users(deuser.Name).groups.Append degroup
            AddUserToGroup GLBServer, LST_Groups.List(i), TXT_First
        Next i
        LST_Name.Clear
        ShowUsers "", FRM_Secure.TXT_Name, LST_Name
    Else
        Call MsgBox("You must enter a unique first name/username.", vbOKOnly + vbExclamation, "No Username")
    End If
    FRM_Users.MousePointer = vbDefault
End Sub

Private Sub CMD_CServer_Click()
    ' change the server from the primary dc to any nt server
    ' Robert May 1-30-98
    On Error Resume Next
    GLBServer = InputBox$("Please enter the server name:", "Enter Server", "\\DBSSERVER")
    If Left$(GLBServer, 2) <> "\\" Then
        GLBServer = "\\" & GLBServer
    End If
    FRM_Users.MousePointer = 13
    FRM_Users.Refresh
    LST_Name.Clear
    LST_Groups.Clear
    LST_Groups.AddItem "Domain Users"
    LST_PosGroups.Clear
    ShowUsers "", FRM_Secure.TXT_Name, LST_Name
    ShowGroups "", FRM_Secure.TXT_Name, LST_PosGroups
    LAB_DC = "Modifying users found on " & GLBServer
    FRM_Users.MousePointer = 0
    
End Sub

Private Sub CMD_DelUser_Click()
    'kill the user--I find it ironic that I have to write so much more code for
    'the add user part than the delete user part!!!
    'Robert May 1-30-98
    Dim dewks As Workspace
    If LST_Name.ListIndex <> -1 Then
        res = MsgBox("WARNING!!!  YOU WILL NOT BE ABLE TO RECREATE " & LST_Name & "'S ACCOUNT AS IT WAS BEFORE (You can add a new user, but privledges will have to be re-established).  ARE YOU SURE YOU WANT TO DELETE " & LST_Name & "?", vbYesNo, "Confirm Delete")
        If res = vbYes Then
            DelUser GLBServer, LST_Name
            With FRM_Secure
                Set dewks = DBEngine.CreateWorkspace(.TXT_Name & Time, .TXT_Name, .TXT_Password)
            End With
            dewks.Users.Delete LST_Name
            ShowUsers "", FRM_Secure.TXT_Name, LST_Name
        End If
    End If
End Sub

Private Sub CMD_DropG_Click()
    ' dump a group
    ' Robert May 1-30-98
    Dim dewks As Workspace
    If LST_Groups = "Domain Users" Then
        Call MsgBox("You cannot remove the Domain Users Group from any users.", vbOKOnly + vbExclamation, "Unable to comply")
        Exit Sub
    End If
    If LST_Groups.ListIndex <> -1 Then
        If LST_Name.ListIndex <> -1 Then
            res = MsgBox("Are you sure you want to remove " & LST_Name & " from the " & LST_Groups & " group?", vbYesNo + vbCritical, "Removing Group")
            If res = vbNo Then Exit Sub
            With FRM_Secure
                Set dewks = DBEngine.CreateWorkspace(.TXT_Name & Time, .TXT_Name, .TXT_Password)
            End With
            dewks.Users(LST_Name).groups.Delete LST_Groups
            DelUserFromGroup GLBServer, LST_Groups, LST_Name
        End If
        LST_Groups.RemoveItem LST_Groups.ListIndex
    End If
End Sub

Private Sub CMD_ShowAll_Click()
    ' reset the form
    Dim domain As String
    Dim Users() As String
    Dim groups() As String
    Dim drops() As String
    Dim AdminGs() As String
    FRM_Users.MousePointer = 13
    FRM_Users.Refresh
    ShowUsers "", FRM_Secure.TXT_Name, LST_Name
    ShowGroups "", FRM_Secure.TXT_Name, LST_PosGroups
    LST_Groups.Clear
    LST_Groups.AddItem "Domain Users"
    FRM_Users.Refresh
    FRM_Users.MousePointer = 0
End Sub

Private Sub Form_Load()
    ' set things up and show initial users
    ' Robert May 1-30-98
    Dim domain As String
    Dim FixedDomain As String
    FRM_Users.Show
    FRM_Users.MousePointer = 13
    FRM_Users.Refresh
    domain = GetPrimaryDCName("", "")
    i = 1
    While Asc(Mid$(domain, i, 1)) <> 0
        FixedDomain = FixedDomain & Mid$(domain, i, 1)
        i = i + 1
    Wend
    GLBServer = FixedDomain
    LAB_DC = "Modifying users found on " & Right(FixedDomain, Len(FixedDomain) - 2)
    ShowUsers "", FRM_Secure.TXT_Name, LST_Name
    ShowGroups "", FRM_Secure.TXT_Name, LST_PosGroups
    LST_Groups.AddItem "Domain Users"
    FRM_Users.MousePointer = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload FRM_Secure
    Unload frmAbout
    Unload FRM_Users
End Sub

Private Sub LST_Name_Click()
    Dim tempgroups() As String
    Dim deme(0) As String
    deme(0) = "none"
    FRM_Users.MousePointer = 13
    FRM_Users.Refresh
    LST_Groups.Clear
    EnumerateGroups GLBServer, LST_Name, tempgroups
    POPLST tempgroups(), LST_Groups, deme
    FRM_Users.MousePointer = 0
    
End Sub

Private Sub LST_PosGroups_DblClick()
    On Error Resume Next
    FRM_Users.MousePointer = 13
    FRM_Users.Refresh
    ShowUsers LST_PosGroups, FRM_Secure.TXT_Name, LST_Name
    FRM_Users.MousePointer = 0

End Sub

Private Sub MNU_about_Click(Index As Integer)
    frmAbout.Show
End Sub

Private Sub MNU_Exit_Click(Index As Integer)
    Unload FRM_Users
End Sub
