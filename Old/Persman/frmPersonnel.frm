VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmPersonnel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personnel Manager"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15270
   Icon            =   "frmPersonnel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   13320
      Pattern         =   "*.bmp"
      TabIndex        =   38
      Top             =   8160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame fmeGroups 
      Caption         =   "Groups"
      Height          =   2175
      Left            =   120
      TabIndex        =   34
      Top             =   8280
      Width           =   2895
      Begin VB.ListBox lstGroups 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1856
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   35
         Top             =   240
         Width           =   2624
      End
   End
   Begin VB.Frame fmeTools 
      Caption         =   "Tools"
      Height          =   975
      Left            =   3120
      TabIndex        =   31
      Top             =   9480
      Width           =   4455
      Begin VB.CommandButton cmdWords 
         Height          =   645
         Left            =   1560
         Picture         =   "frmPersonnel.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton cmdSounds 
         Height          =   645
         Left            =   840
         Picture         =   "frmPersonnel.frx":1934
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Sounds"
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton cmdLeave 
         Height          =   645
         Left            =   120
         Picture         =   "frmPersonnel.frx":1D76
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Leave now"
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear details"
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   5160
      Width           =   975
   End
   Begin VB.Frame fmeDetails 
      Caption         =   "Details"
      Height          =   5295
      Left            =   10560
      TabIndex        =   15
      Top             =   360
      Width           =   4575
      Begin VB.TextBox txtLocation 
         Height          =   304
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2520
         Width           =   3045
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtRank 
         Height          =   288
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1005
         Width           =   3015
      End
      Begin VB.TextBox txtBillet 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1395
         Width           =   3015
      End
      Begin VB.TextBox txtBilletDesc 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   1770
         Width           =   3015
      End
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtChristian 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   615
         Width           =   3015
      End
      Begin VB.Image imgPhoto 
         Height          =   2175
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label lblPhoto 
         Caption         =   "Photo"
         Height          =   270
         Left            =   120
         TabIndex        =   30
         Top             =   3000
         Width           =   1170
      End
      Begin VB.Label lblLocation 
         Caption         =   "Location"
         Height          =   270
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label lblSurname 
         Caption         =   "Surname"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblRank 
         Caption         =   "Rank"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1005
         Width           =   1215
      End
      Begin VB.Label lblBillet 
         Caption         =   "Billet"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Label lblBilletDesc 
         Caption         =   "Billet Description"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblChristian 
         Caption         =   "Christian Name"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   630
         Width           =   1215
      End
   End
   Begin VB.Frame fmeMessages 
      Caption         =   "Messages"
      Height          =   2295
      Left            =   10560
      TabIndex        =   12
      Top             =   5760
      Width           =   4575
      Begin VB.CommandButton cmdClearMail 
         Height          =   645
         Left            =   3120
         Picture         =   "frmPersonnel.frx":21B8
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Update message"
         Top             =   1560
         Width           =   675
      End
      Begin VB.CommandButton cmdMail 
         Height          =   645
         Left            =   3840
         Picture         =   "frmPersonnel.frx":24C2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Update message"
         Top             =   1560
         Width           =   675
      End
      Begin VB.TextBox txtMessages 
         Height          =   1245
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame fmeSearch 
      BackColor       =   &H8000000A&
      Caption         =   "Search"
      Height          =   1095
      Left            =   3120
      TabIndex        =   8
      Top             =   8280
      Width           =   4455
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   645
         Left            =   3480
         Picture         =   "frmPersonnel.frx":27CC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Search for personnel"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Timer timClock 
      Interval        =   30000
      Left            =   5040
      Top             =   3960
   End
   Begin VB.CommandButton cmdIn 
      Caption         =   "<"
      Height          =   400
      Left            =   4920
      TabIndex        =   6
      Top             =   3328
      Width           =   645
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   ">"
      Height          =   400
      Left            =   4920
      TabIndex        =   5
      Top             =   2688
      Width           =   645
   End
   Begin ComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   10590
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   23548
            MinWidth        =   23548
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16/09/98"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1501
            MinWidth        =   1501
            TextSave        =   "6:16"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstOut 
      DragIcon        =   "frmPersonnel.frx":2AD6
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7710
      Left            =   5632
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Personnel currently ashore"
      Top             =   384
      Width           =   4752
   End
   Begin VB.ListBox lstIn 
      DragIcon        =   "frmPersonnel.frx":2F18
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7710
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Personnel currently onboard"
      Top             =   384
      Width           =   4720
   End
   Begin VB.Label lblOut 
      Alignment       =   2  'Center
      Caption         =   "Out"
      Height          =   256
      Left            =   5664
      TabIndex        =   3
      Top             =   128
      Width           =   4672
   End
   Begin VB.Label lblIn 
      Alignment       =   2  'Center
      Caption         =   "In"
      Height          =   256
      Left            =   128
      TabIndex        =   2
      Top             =   128
      Width           =   4720
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "Add New"
         Begin VB.Menu mnuNewUser 
            Caption         =   "User"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuNewGroup 
            Caption         =   "Group"
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuNewID 
            Caption         =   "ID"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuNewLocation 
            Caption         =   "Location"
            Shortcut        =   {F5}
         End
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuMessages 
      Caption         =   "Messages"
      Begin VB.Menu mnuAddMessages 
         Caption         =   "Add/View"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuDeleteMessages 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuOrders 
         Caption         =   "Daily Orders"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Begin VB.Menu mnuSearchGroup 
            Caption         =   "Group"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnuSearchLocation 
            Caption         =   "Location"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnSearchUser 
            Caption         =   "User"
            Shortcut        =   ^U
         End
      End
      Begin VB.Menu mnuSounds 
         Caption         =   "Sounds"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmPersonnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
    With frmPersonnel
        .txtSurname = ""
        .txtChristian = ""
        .txtRank = ""
        .txtBillet = ""
        .txtBilletDesc = ""
        .txtComment = ""
        .txtLocation = ""
    End With

End Sub

Private Sub cmdIn_Click()
    'On Error GoTo erra
    If lstOut.ListCount = 0 Then Exit Sub
    For a = 0 To lstOut.ListCount - 1
        If lstOut().Selected(a) = True Then lstIn.AddItem lstOut.List(a)
        fctOut lstOut.List(a), "came onboard at"
    Next a
    a = 0
redo:
    If lstOut.ListCount = 0 Then Exit Sub
    Do
        If lstOut().Selected(a) = True Then
            lstOut.RemoveItem a
            If a > 0 Then a = 0
            GoTo redo
        End If
        a = a + 1
        If a > lstOut.ListCount - 1 Then Exit Do
    Loop
erra:
    On Error GoTo 0
End Sub

Private Sub cmdLeave_Click()
    reply = MsgBox("This action will log all personnel that are not onboard as being late.  Do you want to do this ?", vbYesNo + vbQuestion, "Confirm")
    If reply = vbYes Then
        MsgBox "Logging personnel as late", vbOKOnly
    Else
        MsgBox "Cancelling last operation", vbOKOnly
    End If
End Sub

Private Sub cmdMail_Click()
    If lstIn.ListIndex = -1 And lstIn.ListIndex = -1 Then Exit Sub
    
End Sub

Private Sub cmdOut_Click()
    'On Error GoTo erra
    If lstIn.ListCount = 0 Then Exit Sub
    For a = 0 To lstIn.ListCount - 1
        If lstIn().Selected(a) = True Then lstOut.AddItem lstIn.List(a)
        fctOut lstIn.List(a), "left at"
    Next a
    a = 0
redo:
    If lstIn.ListCount = 0 Then Exit Sub
    Do
        If lstIn().Selected(a) = True Then
            lstIn.RemoveItem a
            If a > 0 Then a = 0
            GoTo redo
        End If
        a = a + 1
        If a > lstIn.ListCount - 1 Then Exit Do
    Loop
erra:
    On Error GoTo 0
End Sub

Private Sub cmdSearch_Click()
    txtName.SetFocus
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
    SearchName = UCase$(txtName)
    For a = 0 To lstIn.ListCount - 1
        PersName = Right$(lstIn.List(a), Len(lstIn.List(a)) - 13)
        If InStr(1, PersName, SearchName) >= 1 Then
            lstIn.ListIndex = a
            Refresh
            gClickOrigin = "Search"
            gSearch = PersName
            fctPopulateTextBoxes
            Exit Sub
        End If
    Next a
    For a = 0 To lstOut.ListCount - 1
        PersName = Right$(lstOut.List(a), Len(lstOut.List(a)) - 13)
        If InStr(1, PersName, SearchName) >= 1 Then
            lstOut.ListIndex = a
            Refresh
            gClickOrigin = "Search"
            gSearch = PersName
            fctPopulateTextBoxes
            Exit Sub
        End If
    Next a
End Sub

Private Sub cmdSounds_Click()
    frmSounds.Show vbModal
End Sub

Private Sub cmdWords_Click()
    frmOrders.Show vbModal
End Sub

Private Sub Command1_Click()

End Sub

Private Sub fmeSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtName.SetFocus
End Sub

Public Sub Form_Load()
    On Error GoTo erra
    GoTo cont
erra:
    reply = InputBox("Enter a valid path to the PersMan database", "More Information required")
    
cont:
    File1.Path = "D:\Program files\devstudio\vb\graphics\bitmaps\assorted"
    Dim UserNo As Integer
    If reply = "" Then
        dbPath = gDbPath
    Else
        dbPath = reply
    End If
    Dim db As Database
    Dim rst As Recordset
    Set db = OpenDatabase(dbPath & "\PersMan")
    Set rst = db.OpenRecordset("T_Location")
    rst.MoveFirst
    UserNo = 1
    Do Until rst.EOF          ' process all personnel in Persman database
                              ' start by loading each record from appropriate fields
                              ' use '& ""' in the following to get around Nulls in the data
        User(UserNo).ID = rst(0) & ""
        User(UserNo).Rank = rst(1) & ""
        User(UserNo).Christian = rst(2) & ""
        User(UserNo).Surname = rst(3) & ""
        User(UserNo).Number = rst(4) & ""
        User(UserNo).Billet = rst(5) & ""
        User(UserNo).Department = rst(6) & ""
        User(UserNo).BilletDesc = rst(7) & ""
        User(UserNo).Location = rst(8) & ""
        User(UserNo).Tout = rst(9) & ""
        User(UserNo).DOut = rst(10) & ""
        User(UserNo).TIn = rst(11) & ""
        User(UserNo).DIn = rst(12) & ""
        User(UserNo).Message = rst(13) & ""
        User(UserNo).Photo = rst(14) & ""
        User(UserNo).Comment = rst(15) & ""
        Billet = User(UserNo).Billet
        If Right$(Billet, 1) = "N" Then              ' modify billet naming structure
            Billet = Left$(Billet, Len(Billet) - 1)  ' from D001N to 01D
            Billet = Right$(Billet, 2) & Left$(Billet, 1)
            User(UserNo).Billet = Billet
        End If
        spaceinname = InStr(1, User(UserNo).Christian, " ", vbTextCompare)
        If spaceinname > 0 Then ' Multiple firstnames
            User(UserNo).Christian = Left$(User(UserNo).Christian, spaceinname - 1)
        End If
                              ' set up the groups
        For a = 0 To lstGroups.ListCount - 1
            If User(UserNo).Department = lstGroups.List(a) Then flag = 1
        Next a
        If flag = 1 Then
            flag = 0          ' do nothing but reset the flag
        Else
            lstGroups.AddItem User(UserNo).Department
        End If
                              ' set up the temporary variables
        Surname = User(UserNo).Surname & ", "
        Rank = User(UserNo).Rank & " "
        Christian = User(UserNo).Christian & " "
        Billet = User(UserNo).Billet & " "
                              ' Now add the personnel to the appropriate listbox
        PersonnelName = Billet & Rank & Surname & Christian & "(" & User(UserNo).ID & ")"
        If User(UserNo).Location <> "ONBOARD" Then
            frmPersonnel.lstOut.AddItem PersonnelName
        End If
        If User(UserNo).Location = "ONBOARD" Then
            frmPersonnel.lstIn.AddItem PersonnelName
        End If
        rst.MoveNext
        UserNo = UserNo + 1
    Loop
End Sub

Private Sub lstIn_Click()
    'On Error Resume Next
    'lstOut.ListIndex = -1
    'lstIn.SetFocus
    'On Error GoTo 0
    gClickOrigin = "In"
    fctPopulateTextBoxes
End Sub

Private Sub lstIn_DblClick()
    Dim temp As String
    If lstIn.ListIndex <> -1 Then
        'If gDefaultAction = "Properties" Then
        '    gClickOrigin = "In"
        '    If lstIn.ListIndex >= 0 Then frmProperties.Show vbModal
        'End If
        If gDefaultAction = "In/Out" Then
            temp = lstIn.List(lstIn.ListIndex)
            lstIn.RemoveItem (lstIn.ListIndex)
            lstOut.AddItem temp
            fctOut temp, "left at"
        End If
    End If
End Sub

Private Sub lstOut_Click()
    gClickOrigin = "Out"
    fctPopulateTextBoxes
    'On Error Resume Next
    'lstIn.ListIndex = -1
    'lstOut.SetFocus
    'On Error GoTo 0
End Sub

Private Sub lstOut_DblClick()
    Dim temp As String
    If lstOut.ListIndex <> -1 Then
        'If gDefaultAction = "Properties" Then
        '    gClickOrigin = "Out"
        '    If lstOut.ListIndex >= 0 Then frmProperties.Show vbModal
        'End If
        If gDefaultAction = "In/Out" Then
            temp = lstOut.List(lstOut.ListIndex)
            lstOut.RemoveItem (lstOut.ListIndex)
            lstIn.AddItem temp
            fctOut temp, "came onboard at"
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuDelete_Click()
    reply = MsgBox("This will delete the currently selected person.  Are you sure ?", vbYesNo + vbExclamation, "Confirm Action")
    If reply = vbYes Then
        'delete this user
    End If
    
End Sub

Private Sub mnuExit_Click()
    reply = MsgBox("This will exit the application.  Are you sure ?", vbYesNo + vbExclamation, "Confirm Action")
    If reply = vbYes Then
        End
    End If
End Sub

Private Sub mnuNewUser_Click()
    frmNewUser.Show vbModal
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuOrders_Click()
    frmOrders.Show vbModal
End Sub

Private Sub mnuProperties_Click()
    If lstIn.ListIndex <> -1 And lstOut.ListIndex <> -1 Then Exit Sub
    gClickOrigin = "Menu"
    frmProperties.Show vbModal
End Sub

Private Sub mnuSounds_Click()
    frmSounds.Show vbModal
End Sub

Public Function fctOut(PersName As String, action As String)
    On Error Resume Next
    If gLog = True Then
        Open gLogFile For Append As #1
            Print #1, PersName & " " & action & " " & Now
        Close 1
    End If
End Function

Function fctPopulateTextBoxes()
    On Error Resume Next
    If gClickOrigin = "In" Then ' from lstIn
        PersString = frmPersonnel.lstIn.List(frmPersonnel.lstIn.ListIndex)
    End If
    If gClickOrigin = "Out" Then ' from lstOut
        PersString = frmPersonnel.lstOut.List(frmPersonnel.lstOut.ListIndex)
    End If
    If gClickOrigin = "Menu" Then ' select source
        'if lstin.
        reply = MsgBox("Select 'Yes' for In data, Select 'No' for Out data, Select 'Cancel' to cancel", vbYesNoCancel + vbDefaultButton1, "Select data source")
        If reply = vbYes Then PersString = frmPersonnel.lstIn.List(frmPersonnel.lstIn.ListIndex)
        If reply = vbNo Then PersString = frmPersonnel.lstOut.List(frmPersonnel.lstOut.ListIndex)
        If reply = vbCancel Then Exit Function
    End If
    If gClickOrigin = "Search" Then ' from Search function
        PersString = gSearch
    End If
    '
    '
    rn = File1.ListCount - 1
    rn = Int(Rnd(1) * rn) + 1
    
    '
    '
    PersLen = Len(PersString) - 1
    Start = InStr(1, PersString, "(")
    PersID = Mid$(PersString, Start + 1, PersLen)
    PersID = Val(Left$(PersID, Len(PersID) - 1))
    'frmPersonnel.fctPopulateTextBoxes (PersID)
    With frmPersonnel
    'With frmProperties
        .txtSurname = User(PersID).Surname
        .txtChristian = User(PersID).Christian
        .txtRank = User(PersID).Rank
        .txtBillet = User(PersID).Billet
        .txtBilletDesc = User(PersID).BilletDesc
        .txtComment = User(PersID).Comment
        .txtLocation = User(PersID).Location
        .imgPhoto.Picture = LoadPicture(File1.Path & "\" & File1.List(rn))
    End With
End Function

Private Sub timClock_Timer()
    If Time() >= gjs Then
        
    End If
End Sub
