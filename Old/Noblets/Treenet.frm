VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form TreeNet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TreeNet"
   ClientHeight    =   5820
   ClientLeft      =   2520
   ClientTop       =   1935
   ClientWidth     =   5625
   Icon            =   "TreeNet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5820
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ListView lvwLinks 
      Height          =   1275
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   2249
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.TreeView tvwNet 
      Height          =   3855
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   6800
      _Version        =   327682
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5220
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeNet.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeNet.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeNet.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeNet.frx":095A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeNet.frx":0C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeNet.frx":0F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TreeNet.frx":12A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Workstation Connections:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Network Resources:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   75
      Width           =   1740
   End
   Begin VB.Menu popDomain 
      Caption         =   "popDomain"
      Visible         =   0   'False
      Begin VB.Menu popDomainGroups 
         Caption         =   "Groups"
      End
      Begin VB.Menu popDomainUsers 
         Caption         =   "Users"
      End
      Begin VB.Menu popDomSep0 
         Caption         =   "-"
      End
      Begin VB.Menu popDomainRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popServer 
      Caption         =   "popServer"
      Visible         =   0   'False
      Begin VB.Menu popServerInfo 
         Caption         =   "Info"
      End
      Begin VB.Menu popServerSep0 
         Caption         =   "-"
      End
      Begin VB.Menu popServerGroups 
         Caption         =   "Global Groups"
      End
      Begin VB.Menu popServerLocalGroups 
         Caption         =   "Groups"
         Enabled         =   0   'False
      End
      Begin VB.Menu popServerUsers 
         Caption         =   "Users"
      End
      Begin VB.Menu popServerSep1 
         Caption         =   "-"
      End
      Begin VB.Menu popServerServices 
         Caption         =   "Services"
      End
      Begin VB.Menu popServerSep2 
         Caption         =   "-"
      End
      Begin VB.Menu popServerRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu popShare 
      Caption         =   "popShare"
      Visible         =   0   'False
      Begin VB.Menu popShareInfo 
         Caption         =   "Info"
         Enabled         =   0   'False
      End
      Begin VB.Menu popShareOpenFiles 
         Caption         =   "Open Files..."
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "TreeNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public fcolNodes    As Collection

Dim mApp        As NOblets.CApplication
Dim mNetwork    As NOblets.CNetwork
Dim mDomain     As NOblets.CDomain
Dim mServer     As NOblets.CServer
Dim mShare      As NOblets.CShare

Const LVL_NETWORK = 1
Const LVL_DOMAIN = 2
Const LVL_SERVER = 3
Const LVL_SHARE = 4

Const DRIVE_SHARE = 1

Const IMG_ENTIRENET = 1
Const IMG_NETHOOD = 2
Const IMG_MYCOMP = 3
Const IMG_PC02 = 4
Const IMG_DRIVENET = 5
Const IMG_PRINTER = 6
Const IMG_DRIVE01 = 7

Const NETROOT_NT = "Microsoft Windows Network"
Const NETROOT_95 = "Microsoft Network"

Const NERR_ServerNotStarted = 2114 ' The Server service is not started.
Const ERROR_NO_BROWSER_SERVERS = 6118   ' No browser server was found in the domain.
Private Sub ShowLinks()
    
    ' Create an object variable for the ColumnHeader object.
    Dim clmX    As ColumnHeader
    Dim oShare  As NOblets.CShare
    Dim oShare2 As NOblets.CShare
    
    ' Add ColumnHeaders.  The width of the columns is the width
    ' of the control divided by the number of ColumnHeader objects.
    Set clmX = lvwLinks.ColumnHeaders.Add(, , "Resource", lvwLinks.Width * 0.3)
    Set clmX = lvwLinks.ColumnHeaders.Add(, , "Comment", lvwLinks.Width * 0.53, lvwColumnCenter)
    
    lvwLinks.View = lvwReport ' Set View property to Report.

    ' Create a variable to add ListItem objects.
    Dim itmX        As ListItem
    Dim colLinks    As Collection
    
    Set colLinks = mApp.WorkstationLinks()
    For Each oShare In colLinks
        Set itmX = lvwLinks.ListItems.Add(, , oShare.RemoteName, 1)   ' Disk
        itmX.SmallIcon = IMG_DRIVENET   ' NOT ACCURATE!!!
    Next
End Sub

Private Sub ShowNet()
    Dim nodX        As Node
    Dim oNet        As NOblets.CNetwork
    Dim oDomain     As NOblets.CDomain
    Dim colDomains  As Collection
    
    Set mApp = New NOblets.CApplication
    Set fcolNodes = New Collection
    
    tvwNet.Style = tvwTreelinesPlusMinusPictureText     ' Style 7.
    tvwNet.BorderStyle = vbFixedSingle
    
    mApp.Init
    ' Add the Networks
    For Each oNet In mApp.Networks
        Set nodX = tvwNet.Nodes.Add(, , oNet.RemoteName, oNet.RemoteName, IMG_ENTIRENET)
        '                   relative, relationship, key, text, image
        nodX.EnsureVisible
        fcolNodes.Add oNet, Key:=oNet.RemoteName
        
        ' And Add Domains
        oNet.Enumerate
        Set colDomains = oNet.Domains
        For Each oDomain In colDomains
            Set nodX = tvwNet.Nodes.Add(oNet.RemoteName, tvwChild, _
                oDomain.RemoteName, oDomain.RemoteName, IMG_NETHOOD)
            '                   relative, relationship, key, text, image
            fcolNodes.Add oDomain, Key:=oDomain.RemoteName
        Next
    Next
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    Call ShowNet
    Call ShowLinks
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set mApp = Nothing
    Set mNetwork = Nothing
    Set mDomain = Nothing
    Set mServer = Nothing
    Set mShare = Nothing
End Sub


Private Sub popDomainGroups_Click()
    Dim nodX        As Node
    Dim oDomain     As NOblets.CDomain
    Dim oGroup      As NOblets.CGroup
    Dim colGroups   As Collection
    
    ptrHourglass
    
    Set nodX = tvwNet.SelectedItem
    Set oDomain = fcolNodes(nodX.Key)
    
    With ShowUsers
        .lvwUsers.ListItems.Clear
        .Caption = "Show groups in domain " & nodX.Key
        .ShowingType = SHOWING_GROUPS
        
        ' Create a variable to add ListItem objects.
        Dim itmX        As ListItem
        
        Set colGroups = oDomain.Groups()
        For Each oGroup In colGroups
            Set itmX = .lvwUsers.ListItems.Add(, , oGroup.GroupName)   ' Disk
            itmX.SubItems(1) = oGroup.Comment
        Next
        
        ptrNormal
        .Show vbModal
    End With
End Sub

Private Sub popDomainUsers_Click()
    Dim nodX    As Node
    Dim oDomain As NOblets.CDomain
    Dim oUser   As NOblets.CUser
    Dim colUsers    As Collection
    
    ptrHourglass
    
    Set nodX = tvwNet.SelectedItem
    Set oDomain = fcolNodes(nodX.Key)
    
    With ShowUsers
        .lvwUsers.ListItems.Clear
        .Caption = "Show users in domain " & nodX.Key
        .ShowingType = SHOWING_USERS
    
        ' Create a variable to add ListItem objects.
        Dim itmX        As ListItem
        
        Set colUsers = oDomain.Users()
        For Each oUser In colUsers
            Set itmX = .lvwUsers.ListItems.Add(, , oUser.UserName)   ' Disk
            itmX.SubItems(1) = oUser.Comment
        Next
        ptrNormal
        .Show vbModal
    End With
    Set colUsers = Nothing
End Sub

Private Sub popServerGroups_Click()
    Dim nodX        As Node
    Dim oServer     As NOblets.CServer
    Dim oGroup      As NOblets.CGroup
    Dim colGroups   As Collection
    
    ptrHourglass
    
    Set nodX = tvwNet.SelectedItem
    Set oServer = fcolNodes(nodX.Key)
    
    With ShowUsers
        .lvwUsers.ListItems.Clear
        .Caption = "Show global groups on server " & nodX.Key
        .ShowingType = SHOWING_GROUPS
        
        ' Create a variable to add ListItem objects.
        Dim itmX        As ListItem
        
        Set colGroups = oServer.GlobalGroups()
        For Each oGroup In colGroups
            Set itmX = .lvwUsers.ListItems.Add(, , oGroup.GroupName)   ' Disk
            itmX.SubItems(1) = oGroup.Comment
        Next
        ptrNormal
        .Show vbModal
    End With
    
    Set colGroups = Nothing
End Sub

Private Sub popServerInfo_Click()
    Dim nodX    As Node
    Dim sMsg    As String
    Dim oServer As NOblets.CServer
    
    ptrHourglass
    
    Set nodX = tvwNet.SelectedItem
    Set oServer = fcolNodes(nodX.Key)
    
    With oServer
        If .IsNT Then
            sMsg = "Windows NT Platform, "
        Else
            sMsg = "OS2 Platform, "
        End If
        sMsg = sMsg & "version " & .Version & vbCrLf

        If .IsPDC Then
            sMsg = sMsg & "Primary Domain Controller" & vbCrLf
        ElseIf .IsBDC Then
            sMsg = sMsg & "Backup Domain Controller" & vbCrLf
        End If
        
        If .IsBrowser Then
            If .IsMasterBrowser Then
                sMsg = sMsg & "Master Browser" & vbCrLf
            Else
                sMsg = sMsg & "Backup Browser" & vbCrLf
            End If
        End If
        
        If .IsSQLServer Then _
            sMsg = sMsg & "SQL Server" & vbCrLf
    End With
    
    ptrNormal
    MsgBox sMsg
End Sub

Private Sub popServerOpenFiles_Click()
    ' NOT YET IMPLEMENTED
    Exit Sub

    Dim nodX        As Node
    Dim colFiles    As Collection
    Dim oServer     As NOblets.CServer
    
    Set nodX = tvwNet.SelectedItem
    Set oServer = fcolNodes(nodX.Key)
End Sub


Private Sub popServerSchedule_Click()
    Dim nodX    As Node
    Dim colJobs As Collection
    Dim oServer As NOblets.CServer
    
    Set nodX = tvwNet.SelectedItem
    Set oServer = fcolNodes(nodX.Key)
    Set colJobs = oServer.Schedule

End Sub


Private Sub popServerServices_Click()
    Dim nodX        As Node
    Dim oServer     As NOblets.CServer
    Dim oService    As NOblets.CService
    Dim colServices As Collection
    
    ptrHourglass
    
    Set nodX = tvwNet.SelectedItem
    Set oServer = fcolNodes(nodX.Key)
    
    With ShowUsers
        .lvwUsers.ListItems.Clear
        .Caption = "Show Services on server " & nodX.Key
        .ShowingType = SHOWING_SERVICES
        .Server = nodX.Key
    
        ' Create a variable to add ListItem objects.
        Dim itmX        As ListItem
        
        On Local Error Resume Next
        Set colServices = oServer.Services
        Set colServices = oServer.Services
        On Local Error GoTo 0
        
        For Each oService In colServices
            Set itmX = .lvwUsers.ListItems.Add(, oService.Name, oService.Name)
            itmX.SubItems(1) = oService.DisplayName
            itmX.SubItems(2) = oService.StatusString
        Next
        
        ptrNormal
        .Show vbModal
    End With
    
    Set colServices = Nothing
End Sub

Private Sub popServerUsers_Click()
    Dim nodX    As Node
    Dim oServer As NOblets.CServer
    Dim oUser   As NOblets.CUser
    Dim colUsers    As Collection
    
    ptrHourglass
    
    Set nodX = tvwNet.SelectedItem
    Set oServer = fcolNodes(nodX.Key)
    
    With ShowUsers
        .lvwUsers.ListItems.Clear
        If oServer.IsPDC Or oServer.IsBDC Then
            .Caption = "Show users on server " & nodX.Key _
                & " (domain controller, domain users)"
        Else
            .Caption = "Show users on server " & nodX.Key
        End If
        .ShowingType = SHOWING_USERS
    
        ' Create a variable to add ListItem objects.
        Dim itmX        As ListItem
        
        Set colUsers = oServer.Users()
        For Each oUser In colUsers
            Set itmX = .lvwUsers.ListItems.Add(, , oUser.UserName)   ' Disk
            itmX.SubItems(1) = oUser.Comment
        Next
        
        ptrNormal
        .Show vbModal
    End With
    
    Set colUsers = Nothing
End Sub


Private Sub tvwNet_Expand(ByVal Node As Node)
    
    Dim oNetwork    As NOblets.CNetwork
    Dim oDomain     As NOblets.CDomain
    Dim oServer     As NOblets.CServer
    Dim oShare      As NOblets.CShare
    
    Dim oObject     As Object
    Dim colObjects  As Collection
    Dim colChilds   As Collection
    
    Dim nodChild    As Node
    Dim nodX        As Node
    Dim iLevel      As Integer
    Dim iIdx        As Integer
    Dim iStart      As Integer
    Dim sKey        As String
    
    Dim sDir        As String
    
    Me.MousePointer = vbHourglass
    
    If Node.Parent Is Nothing Then
        iLevel = 1
    Else
        iLevel = Val(Node.Parent.Tag) + 1
    End If
    sKey = Node.Key
    
    ' If this has not been expanded before,
    ' expand it now
    If Len(Node.Tag) = 0 Then
    
        On Local Error Resume Next
    
        Set oObject = fcolNodes(sKey)
        If Err.Number > 0 Then
            Screen.MousePointer = vbNormal
            MsgBox Err.Description, vbCritical + vbOKOnly, App.Title & ".Enumerate"
            Exit Sub
        End If
        oObject.Enumerate
        If Err.Number = ERROR_NO_BROWSER_SERVERS Then   ' No browser
            Screen.MousePointer = vbNormal
            MsgBox Err.Description, vbCritical + vbOKOnly, App.Title & ".Enumerate"
            Exit Sub
        End If
        
        Select Case iLevel
            Case LVL_NETWORK
                Set colObjects = oObject.Domains
                For Each oDomain In colObjects
                    oDomain.Enumerate
                    Set colChilds = oDomain.Servers
                    For Each oServer In colChilds
                        ' relative, relationship, key, text, image, selectedimage
                        Set nodX = tvwNet.Nodes.Add(oDomain.RemoteName, tvwChild, _
                            oServer.RemoteName, _
                            oServer.RemoteName & "      - " & oServer.Comment, _
                            IMG_MYCOMP)
                        nodX.Tag = LVL_SERVER
                        fcolNodes.Add oServer, Key:=oServer.RemoteName
                    Next
                Next
            Case LVL_DOMAIN
                Set colObjects = oObject.Servers
                For Each oServer In colObjects
                    On Local Error Resume Next
                    oServer.Enumerate
                    ' Server service not started on a workstation (no shares)
                    If Err.Number = NERR_ServerNotStarted Then
                        Screen.MousePointer = vbNormal
                        Exit Sub
                    End If
                    Set colChilds = oServer.Shares
                    For Each oShare In colChilds
                        If oShare.ShareType = DRIVE_SHARE Then
                            Set nodX = tvwNet.Nodes.Add(oServer.RemoteName, tvwChild, oShare.RemoteName, _
                                oShare.RemoteName & "      - " & oShare.Comment, IMG_DRIVENET)
                        Else
                            Set nodX = tvwNet.Nodes.Add(oServer.RemoteName, tvwChild, oShare.RemoteName, _
                                oShare.RemoteName & "      - " & oShare.Comment, IMG_PRINTER)
                        End If
                        nodX.Tag = LVL_SHARE
                        fcolNodes.Add oShare, Key:=oShare.RemoteName
                    Next
                Next
            Case LVL_SERVER, LVL_SHARE
                Screen.MousePointer = vbNormal
                Exit Sub
        End Select
        
    End If
    
    ' Set tag to indicate this node has been expanded
    Node.Tag = iLevel
    
    ' Get a Paint,
    ' because listing files could also be lengthy
    tvwNet.Refresh
    
    Me.MousePointer = vbNormal
    
End Sub


Private Sub tvwNet_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim nodX    As Node
    If Button = vbRightButton Then
        Set nodX = tvwNet.SelectedItem
        Select Case Val(nodX.Tag)
            Case LVL_DOMAIN
                PopupMenu popDomain
            Case LVL_SERVER
                PopupMenu popServer
            Case LVL_SHARE
                PopupMenu popShare
        End Select
    End If
End Sub


