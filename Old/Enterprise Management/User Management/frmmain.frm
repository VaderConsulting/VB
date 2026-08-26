VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enterprise User Management"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
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
      Height          =   2535
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   3240
      Width           =   6855
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5790
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7197
            MinWidth        =   7197
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label lblUserInfo 
      Caption         =   "Groupname"
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
      Left            =   120
      TabIndex        =   11
      Top             =   3000
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
      Index           =   7
      Left            =   3120
      TabIndex        =   10
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label lblUserInfo 
      Alignment       =   2  'Center
      Caption         =   "P"
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
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblUserInfo 
      Caption         =   "Flags"
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
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblUserInfo 
      Alignment       =   2  'Center
      Caption         =   "E"
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
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblUserInfo 
      Alignment       =   2  'Center
      Caption         =   "D"
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
      Left            =   5880
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblUserInfo 
      Alignment       =   2  'Center
      Caption         =   "L"
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
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   135
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
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   3735
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   15000
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   15000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuDefaults 
         Caption         =   "Defaults"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim ObjectList() As ScreenObject
'Dim CurrentObject As Object
'Dim intReferenceHeight As Integer
'Dim intReferenceWidth As Integer

'Sub AutoScale()
'    Dim dblXMultiplier As Double
'    Dim dblYMultiplier As Double
'    Dim intObjectNumber As Integer
'    Dim intFontSize As Integer
'
'    '   Get ratio of initial form size to current form size
'    dblXMultiplier = Me.Height / intReferenceHeight
'    dblYMultiplier = Me.Width / intReferenceWidth
'
'    On Error Resume Next
'    '   resize each object
'    For intObjectNumber = 0 To UBound(ObjectList)
'        For Each CurrentObject In Me
'            If CurrentObject.TabIndex = ObjectList(intObjectNumber).Index Then
'                 If Err.Number = 0 Then
'                 With CurrentObject
'                    .Left = ObjectList(intObjectNumber).Left * dblYMultiplier
'                    .Width = ObjectList(intObjectNumber).Width * dblYMultiplier
'                    .Height = ObjectList(intObjectNumber).Height * dblXMultiplier
'                    .Top = ObjectList(intObjectNumber).Top * dblXMultiplier
'                 End With
'                 End If
'            End If
'        Next CurrentObject
'    Next intObjectNumber
'
'    '   This is a sample of how to rescale the font of an object as well.  You can see the
'    '   effect best when you maximize the form.
'    If Int(dblXMultiplier) > 0 Then
'        intFontSize = Int(dblXMultiplier * 8)
'        lblFontResize.FontSize = intFontSize
'        txtSample1.FontSize = intFontSize
'        lblInstructions.FontSize = intFontSize
'    End If
'End Sub

'Sub GetCurrentPositions()
'    Dim intObjectNumber As Integer
'    '   Load the current positions of each object into a user defined type array.
'    '   This information will be used to rescale them in the AutoScale function.
'    On Error Resume Next
'    For Each CurrentObject In Me
'        ReDim Preserve ObjectList(intObjectNumber)
'        With ObjectList(intObjectNumber)
'            .Name = CurrentObject
'            .Index = CurrentObject.TabIndex
'            .Left = CurrentObject.Left
'            .Top = CurrentObject.Top
'            .Width = CurrentObject.Width
'            .Height = CurrentObject.Height
'        End With
'        intObjectNumber = intObjectNumber + 1
'    Next CurrentObject
'
'    '   This is what the object sizes will be compared to on rescaling.
'        intReferenceHeight = Me.Height
'        intReferenceWidth = Me.Width
'End Sub

Sub GetUsers()
    Dim Domain As Object
    Dim member As IADsUser
    Dim UserInfo(15) As String
    Dim TotalUsers As Long
    Dim UserFlags As Long
    
    lstUsers.Clear
    
    sbrInfo.Panels(3).Text = "Opening ADSI Object: WinNT://" & gstrDomain & ",computer"
    Set Domain = GetObject("WinNT://" & gstrDomain & ",computer")
    Domain.Filter = Array("user")
'    If Err.Number <> 0 Then
'        msg = "Error contacting " & strDomain
'        lstResults.AddItem msg
'        End
'    End If
    TotalUsers = 0
    For Each member In Domain
        UserInfo(0) = member.Name
        sbrInfo.Panels(3).Text = "Retrieving info for " & UserInfo(0)
        UserInfo(1) = Left(member.Description, 60)
        UserFlags = member.UserFlags
        If UserFlags And ADS_UF_DONT_EXPIRE_PASSWD Then
            UserInfo(2) = "Y"
        Else
            UserInfo(2) = "N"
        End If
        If member.IsAccountLocked = True Then
            UserInfo(3) = "Y"
        Else
            UserInfo(3) = "N"
        End If
        If member.AccountDisabled = True Then
            UserInfo(4) = "Y"
        Else
            UserInfo(4) = "N"
        End If
        If (member.PasswordExpirationDate < Now) And ((UserFlags And ADS_UF_DONT_EXPIRE_PASSWD) = 0) Then
            UserInfo(5) = "Y"
        Else
            UserInfo(5) = "N"
        End If
        
        lstUsers.AddItem UserInfo(0) & vbTab & _
                         UserInfo(1) & vbTab & _
                         UserInfo(2) & vbTab & _
                         UserInfo(3) & vbTab & _
                         UserInfo(4) & vbTab & _
                         UserInfo(5) & vbTab & _
                         UserFlags
        TotalUsers = TotalUsers + 1
        
        sbrInfo.Panels(1).Text = "Users: " & TotalUsers
        
        Erase UserInfo ' clear UserInfo array
        UserFlags = 0
    Next
    sbrInfo.Panels(1).Text = "Total Users: " & TotalUsers
    sbrInfo.Panels(3).Text = "Idle"
    Set Domain = Nothing
End Sub

Sub GetGroups()
    Dim Domain As Object
    Dim Group As IADsGroup
    Dim TotalGroups As Long
    
    lstGroups.Clear
    
    sbrInfo.Panels(3).Text = "Opening ADSI Object: WinNT://" & gstrDomain & ",computer"
    Set Domain = GetObject("WinNT://" & gstrDomain & ",computer")
    Domain.Filter = Array("group")
    For Each Group In Domain
        sbrInfo.Panels(3).Text = "Retrieving info for " & Group.Name
        lstGroups.AddItem Group.Name & vbTab & Group.Description
        TotalGroups = TotalGroups + 1
        sbrInfo.Panels(2).Text = "Groups: " & TotalGroups
    Next
    sbrInfo.Panels(2).Text = "Total Groups: " & TotalGroups
    sbrInfo.Panels(3).Text = "Idle"
End Sub

Private Sub Form_Load()
    gstrDomain = "WARNBRO"
    gstrPDC = "SERVER"
    
    ' For form resizing
    'GetCurrentPositions
    
    Me.Show
    Me.Refresh
    
    Screen.MousePointer = vbHourglass
    ' Set tab positions
    Call SetListTabStops(lstUsers.hWnd, 18, 70, 73, 76, 79, 83)
    GetUsers
    ' Set tab positions
    Call SetListTabStops(lstGroups.hWnd, 40)
    GetGroups
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

'Private Sub Form_Resize()
'    Call AutoScale
'End Sub

Private Sub lstUsers_DblClick()
    On Error Resume Next
    If Trim(lstUsers.List(lstUsers.ListIndex)) = "" Then Exit Sub
    frmUserInfo.Show vbModal
End Sub

Private Sub mnuDefaults_Click()
    frmDefaults.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub
