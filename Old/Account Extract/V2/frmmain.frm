VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Domain Accounts"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11205
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11205
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbDomains 
      Height          =   315
      ItemData        =   "frmMain.frx":0442
      Left            =   1320
      List            =   "frmMain.frx":0444
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.OptionButton optOS 
      Caption         =   "2K"
      Height          =   255
      Index           =   1
      Left            =   12480
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton optOS 
      Caption         =   "NT"
      Height          =   255
      Index           =   0
      Left            =   12480
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5520
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ImageList imlMembers 
      Left            =   13680
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0446
            Key             =   "User"
            Object.Tag             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0526
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":070B
            Key             =   "Computer"
            Object.Tag             =   "Computer"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":083D
            Key             =   "Computer_Error"
            Object.Tag             =   "Computer_Error"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":108F
            Key             =   "Computer_OK"
            Object.Tag             =   "Computer_OK"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18E1
            Key             =   "Server"
            Object.Tag             =   "Server"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C34
            Key             =   "Server_Error"
            Object.Tag             =   "Server_Error"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2488
            Key             =   "Server_OK"
            Object.Tag             =   "Server_OK"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CDC
            Key             =   "Computer_Unknown"
            Object.Tag             =   "Computer_Unknown"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMembers 
      Height          =   6735
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11880
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "imlMembers"
      SmallIcons      =   "imlMembers"
      ColHdrIcons     =   "imlMembers"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   9360
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtAuthDomain 
      Height          =   285
      Left            =   6360
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtPDC 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CheckBox chkEnum 
      Caption         =   "Computers"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox chkEnum 
      Caption         =   "Groups"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CheckBox chkEnum 
      Caption         =   "Users"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox lstBat 
      Height          =   645
      Left            =   12480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
   End
   Begin VB.ListBox lstMembers 
      Height          =   2595
      Left            =   12480
      MouseIcon       =   "frmMain.frx":3530
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username"
      Height          =   255
      Left            =   5160
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      Height          =   255
      Left            =   8160
      TabIndex        =   21
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblDomain 
      Caption         =   "Domain"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblPDC 
      Caption         =   "PDC:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblGroups 
      Height          =   255
      Left            =   3720
      TabIndex        =   18
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label lblGroupNo 
      Caption         =   "No of Groups:"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label lblComputerNo 
      Caption         =   "No of Computers:"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label lblComputers 
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label lblUsers 
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label lblUserNo 
      Caption         =   "No of Users:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label lblTargetDC 
      Caption         =   "Domain Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   8040
      Width           =   10935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   15000
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Paused As Boolean
Public bExtracting As Boolean

Const ERR_ACCESS_DENIED As Long = 70
Const ERR_SUCCESS As Long = 0
Const ERR_MISMATCH As Long = 13

'Private Sub cmbDomains_Change()
'    If cmbDomains.Text = "" Then Exit Sub
'    txtPDC = GetPDCName(cmbDomains.Text)
'End Sub

Private Sub cmbDomains_Click()
    'Dim tmpDomain() As Byte
    'If cmbDomains.Text = "" Then Exit Sub
    'tmpDomain() = cmbDomains.Text
    'txtPDC = GetPDCName(tmpDomain())
End Sub

Private Sub Form_Load()
    Dim commandLine As Boolean, msg As String
    msg = "Application Started."
    LogLocal msg
    Paused = False
    txtUsername = Environ("username")
    txtAuthDomain = Environ("userdomain")
    txtdomain = Environ("userdomain")
    GetDomainList (SV_TYPE_DOMAIN_ENUM)
    Me.Show
    Me.Refresh
End Sub

Sub LogLocal(msg As String)
    Open App.Path & "\appLog" & Format(Date, "DD-MM-YY") & ".log" For Append As #1
        Print #1, Format(Date, "DD/MM/YYYY") & " " & Format(Time, "Long Time") & " " & msg
    Close 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msg As String
    msg = "Application Ended."
    LogLocal msg
End Sub


Private Sub lvwMembers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If bExtracting Then Exit Sub
    Select Case Left(Item.Key, 1)
        Case "U"
            lblStatus = Item.Text & " (User)"
        Case "G"
            lblStatus = Item.Text & " (Group)"
        Case "C"
            If Len(Trim(Item.Tag)) > 1 Then
                lblStatus = Item.Text & " (IP Address " & Item.Tag & ")"
            Else
                lblStatus = Item.Text & " (IP Address UNKNOWN)"
            End If
    End Select
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuExport_Click()
    Dim Filename As String
    Filename = "c:\temp\"
    If chkEnum(0).Value = vbChecked Then Filename = Filename & "Users, "
    If chkEnum(1).Value = vbChecked Then Filename = Filename & "Groups and "
    If chkEnum(2).Value = vbChecked Then Filename = Filename & "Computers "
    Filename = Filename & "of " & cmbDomains.Text & ".csv"
    Open Filename For Output As #1
        Print #1, "Name,Class,IP Address/Full Name,Type/Home,State/Profile,Type/Description"
        'For lp = 0 To lstMembers.ListCount - 1
        For lp = 1 To lvwMembers.ListItems.Count
            'Print #1, lstMembers.List(lp)
            Select Case Left(lvwMembers.ListItems(lp).Key, 1)
                Case "U"
                    Print #1, lvwMembers.ListItems(lp).Text & ",User," & lvwMembers.ListItems(lp).Tag
                Case "G"
                    Print #1, lvwMembers.ListItems(lp).Text & ",Group"
                Case "C"
                    p = InStr(1, lvwMembers.ListItems(lp).Tag, ",")
                    If p = 0 Then
                        'Print #1, lvwMembers.ListItems(lp).Text & ",Computer,Not Resolved"
                        Print #1, lvwMembers.ListItems(lp).Text & ",Computer," & lvwMembers.ListItems(lp).Tag
                    Else
                        If Right(lvwMembers.ListItems(lp).Tag, Len(lvwMembers.ListItems(lp).Tag) - p) = "Up" Then
                            Print #1, lvwMembers.ListItems(lp).Text & ",Computer," & lvwMembers.ListItems(lp).Tag
                        Else
                            Print #1, lvwMembers.ListItems(lp).Text & ",Computer," & lvwMembers.ListItems(lp).Tag
                        End If
                    End If
            End Select
        Next lp
    Close 1
    MsgBox "File saved in " & Filename
End Sub

Private Sub mnuPause_Click()
    If Paused Then
        Paused = False
        mnuPause.Caption = "Pause"
        lblStatus = "Continuing"
    Else
        Paused = True
        mnuPause.Caption = "Continue"
        lblStatus = "Paused"
    End If
    frmMain.Refresh
    If Paused = True Then
        Do Until Paused = False
            DoEvents
        Loop
    End If
End Sub

Private Sub mnuStart_Click()
    Dim msg As String
    Dim TotalUsers As Integer
    Dim TotalGroups As Integer
    Dim TotalComputers As Integer
    
    'On Error Resume Next
    'If optOS(0).Value = True Then
    '    Dim Domain As IADsDomain  ' <--- WinNT
    'Else
        'Dim domain As Object '  Object      ' retrieves users and groups
        Dim domain As IADsDomain
    'End If
    Dim Computer As IADsComputer
    Dim DomainError As Boolean
    Dim HostIP As String
    Dim MemberName As String
    Dim strDomain As String
    Dim strPDC As String
    Dim Ping As DSPINGLib.Ping
    Dim tmpDomain As String
    Dim strType As String
    
    If cmbDomains.Text = "" Then
        lblStatus = "Domain name must be specified"
        Exit Sub
    End If
    If txtPDC = "" Then
        lblStatus = "PDC name must be specified"
        Exit Sub
    End If
    
    Set Ping = CreateObject("DSPing.Ping")
    bExtracting = True
    
    lstMembers.Clear
    lvwMembers.ListItems.Clear
    mnuPause.Enabled = True
    mnuExport.Enabled = False
    mnuStart.Enabled = False
    lblStatus = "Initialising..."
    lblStatus.Refresh
    lblUsers = 0
    lblGroups = 0
    lblComputers = 0
    
    strDomain = cmbDomains.Text
    strPDC = txtPDC.Text
    
    lblStatus = "Authenticating to " & strPDC
    lblStatus.Refresh
    
    SocketsInitialize
    
    'retval = Connect2("", "\\" & strPDC & "\IPC$", txtAuthDomain & "\" & txtUsername, txtPassword)
    
    If retval = 0 Then
        lblStatus = "Contacting " & strDomain
        lblStatus.Refresh
        'Set domain = GetObject("WinNT://" & strPDC & ",computer")   ' <--- Retrieves Users and Groups only
        Set domain = GetObject("WinNT://" & strDomain)
        Select Case Err.Number
            Case ERR_SUCCESS
                msg = "Retrieving members of " & strDomain
                LogLocal msg
                DomainError = False
            Case ERR_MISMATCH
                msg = "Error contacting " & strDomain & ". Is this a Domain name?"
                'lstResults.AddItem msg
                lblStatus = msg
                LogLocal msg
                DomainError = True
            Case ERR_ACCESS_DENIED
                msg = "Access denied whilst contacting " & strDomain
                'lstResults.AddItem msg
                lblStatus = msg
                LogLocal msg
                DomainError = True
            Case Else
                msg = "Error " & Err.Number & " (" & Err.Description & ") whilst contacting " & strDomain
                msg = Replace(msg, vbCr, " ")
                msg = Replace(msg, vbLf, " ")
                'lstResults.AddItem msg
                lblStatus = msg
                LogLocal msg
                DomainError = True
        End Select
        DoEvents
        If Not DomainError Then
            TotalMembers = 0
            For Each member In domain
                If (chkEnum(0) And member.Class = "User") Or (chkEnum(1) And member.Class = "Group") Or (chkEnum(2) And member.Class = "Computer") Then
                MemberName = UCase(member.Name)
                
                lblStatus = "Retrieving info for " & MemberName
                lblStatus.Refresh
                lstMembers.AddItem MemberName & "," & member.Class
                Select Case member.Class
                    Case "User"
                        TotalUsers = TotalUsers + 1
                        lvwMembers.ListItems.Add lvwMembers.ListItems.Count + 1, "U" & MemberName, MemberName, "User"
                        lvwMembers.ListItems(lvwMembers.ListItems.Count).Tag = member.FullName & "," & member.HomeDirectory & "," & member.Profile & "," & member.Description & "," & member.userflags
                    Case "Group"
                        TotalGroups = TotalGroups + 1
                        lvwMembers.ListItems.Add lvwMembers.ListItems.Count + 1, "G" & MemberName, MemberName, "Group"
                    Case "Computer"
                        TotalComputers = TotalComputers + 1
                        HostIP = ""
                        HostIP = Trim(ResolveHostname(MemberName))
                        If HostIP <> "" Then
                            result = Ping.DoPing(MemberName)
                            If result = 0 Then ' ping was successful
                                strType = ""
                                strType = HostType(MemberName)
                                If strType = "" Then strType = "Workstation" ' Assume workstation if not determined
                                If strType = "Server" Then
                                    lvwMembers.ListItems.Add lvwMembers.ListItems.Count + 1, "C" & MemberName, MemberName, "Server_OK"
                                ElseIf strType = "Workstation" Then
                                    lvwMembers.ListItems.Add lvwMembers.ListItems.Count + 1, "C" & MemberName, MemberName, "Computer_OK"
                                ElseIf strType = "Server and Workstation" Then
                                    lvwMembers.ListItems.Add lvwMembers.ListItems.Count + 1, "C" & MemberName, MemberName, "Server_OK"
                                Else
                                    lvwMembers.ListItems.Add lvwMembers.ListItems.Count + 1, "C" & MemberName, MemberName, "Computer_OK"
                                End If
                                lvwMembers.ListItems(lvwMembers.ListItems.Count).Tag = HostIP & "," & strType & ",Up"
                                'tmpDomain = strDomain
                                'GetServerInfo MemberName, strDomain
                                'strDomain = tmpDomain
                                'ListServers SV_TYPE_ALL, strDomain, MemberName
                                'strDomain = tmpDomain
                            Else ' ping not successful
                                lvwMembers.ListItems.Add lvwMembers.ListItems.Count + 1, "C" & MemberName, MemberName, "Computer_Unknown"
                                lvwMembers.ListItems(lvwMembers.ListItems.Count).Tag = HostIP & ",Unknown Type,Down"
                            End If
                        Else ' Name could not be resolved
                            lvwMembers.ListItems.Add lvwMembers.ListItems.Count + 1, "C" & MemberName, MemberName, "Computer_Error"
                            lvwMembers.ListItems(lvwMembers.ListItems.Count).Tag = "Unknown IP,Unknown Type,Not Resolved"
                        End If
                End Select
                
                lblUsers = TotalUsers
                lblGroups = TotalGroups
                lblComputers = TotalComputers
                frmMain.Refresh
                DoEvents
                End If
            Next
            lblStatus = "Idle"
            lblStatus.Refresh
        End If
        
        msg = "Completed task"
        LogLocal msg
        Set domain = Nothing
    Else
        lblStatus = "Error " & retval & " authenticating to " & txtPDC
    End If
    If lstMembers.ListCount = 0 Then
        mnuExport.Enabled = False
    Else
        mnuExport.Enabled = True
    End If
    
    SocketsCleanup
    Set Ping = Nothing
    
    mnuPause.Enabled = False
    mnuStart.Enabled = True
    bExtracting = False
End Sub

Function HostType(ServerName As String) As String
    Dim strFilename As String, tmpString As String
    Dim WFlag As Boolean, SFlag As Boolean ' Workstation and Server
    strFilename = "\\" & ServerName & "\c$\boot.ini"
    lblStatus = "Determining OS on " & ServerName
    lblStatus.Refresh
    retval = Connect2("", "\\" & ServerName & "\IPC$", txtAuthDomain & "\" & txtUsername, txtPassword)
    On Error Resume Next
    Open strFilename For Input As #1
        If Err.Number = 0 Then
        Do Until EOF(1) Or Err.Number <> 0
            Line Input #1, tmpString
            If InStr(1, UCase(tmpString), "WORKSTATION") > 0 Then WFlag = True
            If InStr(1, UCase(tmpString), "PROFESSIONAL") > 0 Then WFlag = True
            If InStr(1, UCase(tmpString), "SERVER") > 0 Then SFlag = True
        Loop
        End If
    Close 1
    If WFlag Then HostType = "Workstation"
    If SFlag Then HostType = "Server"
    If WFlag And SFlag Then HostType = "Server and Workstation"
    DoEvents
    retval = DisConnect2("\\" & ServerName & "\IPC$", True)
End Function

