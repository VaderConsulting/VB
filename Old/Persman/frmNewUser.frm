VERSION 5.00
Begin VB.Form frmNewUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New User"
   ClientHeight    =   5856
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9088
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5856
   ScaleWidth      =   9088
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGetPhoto 
      Caption         =   "Get..."
      Height          =   272
      Left            =   8320
      TabIndex        =   27
      Top             =   1408
      Width           =   656
   End
   Begin VB.PictureBox picPhoto 
      Height          =   3088
      Left            =   5888
      ScaleHeight     =   3024
      ScaleWidth      =   3024
      TabIndex        =   25
      ToolTipText     =   "Photo"
      Top             =   1792
      Width           =   3088
   End
   Begin VB.TextBox txtLocation 
      Enabled         =   0   'False
      Height          =   304
      Left            =   5888
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   896
      Width           =   3088
   End
   Begin VB.Frame fmeGroups 
      Caption         =   "Groups"
      Height          =   2415
      Left            =   128
      TabIndex        =   20
      Top             =   2640
      Width           =   5600
      Begin VB.ListBox lstMember 
         Height          =   1520
         ItemData        =   "frmNewUser.frx":0442
         Left            =   120
         List            =   "frmNewUser.frx":0444
         TabIndex        =   10
         Top             =   480
         Width           =   2256
      End
      Begin VB.ListBox lstGroups 
         Height          =   1520
         ItemData        =   "frmNewUser.frx":0446
         Left            =   3200
         List            =   "frmNewUser.frx":0448
         TabIndex        =   13
         Top             =   480
         Width           =   2288
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<"
         Height          =   375
         Left            =   2432
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   ">"
         Height          =   375
         Left            =   2432
         TabIndex        =   12
         Top             =   1536
         Width           =   615
      End
      Begin VB.Label lblMember 
         Alignment       =   2  'Center
         Caption         =   "Member of"
         Height          =   256
         Left            =   128
         TabIndex        =   22
         Top             =   240
         Width           =   2288
      End
      Begin VB.Label lblListed 
         Alignment       =   2  'Center
         Caption         =   "Available"
         Height          =   256
         Left            =   3200
         TabIndex        =   21
         Top             =   256
         Width           =   2288
      End
   End
   Begin VB.TextBox txtChristian 
      Height          =   285
      Left            =   1536
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   1536
      TabIndex        =   5
      Top             =   2048
      Width           =   4215
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import..."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   368
      Left            =   7952
      TabIndex        =   8
      Top             =   5408
      Width           =   976
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6752
      TabIndex        =   7
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtBilletDesc 
      Height          =   285
      Left            =   1536
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1664
      Width           =   4215
   End
   Begin VB.TextBox txtBillet 
      Height          =   285
      Left            =   1536
      TabIndex        =   3
      Top             =   1280
      Width           =   4215
   End
   Begin VB.TextBox txtRank 
      Height          =   285
      Left            =   1536
      TabIndex        =   2
      Top             =   896
      Width           =   4215
   End
   Begin VB.TextBox txtSurname 
      Height          =   285
      Left            =   1536
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.CheckBox chkOnboard 
      Caption         =   "Onboard"
      Height          =   255
      Left            =   5888
      TabIndex        =   6
      Top             =   128
      Width           =   1335
   End
   Begin VB.Label lblPhoto 
      Alignment       =   2  'Center
      Caption         =   "Photo"
      Height          =   272
      Left            =   5888
      TabIndex        =   26
      Top             =   1408
      Width           =   3088
   End
   Begin VB.Label lblLocation 
      Alignment       =   2  'Center
      Caption         =   "Location"
      Height          =   272
      Left            =   5888
      TabIndex        =   24
      Top             =   512
      Width           =   3088
   End
   Begin VB.Label lblChristian 
      Caption         =   "Christian Name"
      Height          =   256
      Left            =   128
      TabIndex        =   19
      Top             =   512
      Width           =   1328
   End
   Begin VB.Label lblComment 
      Caption         =   "Comment"
      Height          =   256
      Left            =   128
      TabIndex        =   18
      Top             =   2048
      Width           =   1328
   End
   Begin VB.Label lblBilletDesc 
      Caption         =   "Billet Description"
      Height          =   256
      Left            =   128
      TabIndex        =   17
      Top             =   1664
      Width           =   1328
   End
   Begin VB.Label lblBillet 
      Caption         =   "Billet"
      Height          =   256
      Left            =   128
      TabIndex        =   16
      Top             =   1280
      Width           =   1328
   End
   Begin VB.Label lblRank 
      Caption         =   "Rank"
      Height          =   256
      Left            =   128
      TabIndex        =   15
      Top             =   896
      Width           =   1328
   End
   Begin VB.Label lblSurname 
      Caption         =   "Surname"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkOnboard_Click()
    If chkOnboard.Value = vbUnchecked Then
        txtLocation.Enabled = True
    Else
        txtLocation.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    If lstGroups.ListIndex > -1 Then
        MsgBox "Move selected group to left hand list"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
    dbPath = gDbPath
    Dim db As Database
    Dim rst As Recordset
    Dim rank As String * 8
    Set db = opendatabase(dbPath & "\PersMan")
    Set rst = db.OpenRecordset("T_personnel_Data")
    rst.MoveFirst
    Do Until rst.EOF
        billet = rst(0)
        If Right$(billet, 1) = "N" Then
            billet = Left$(billet, Len(billet) - 1)
            billet = Right$(billet, 2) & Left$(billet, 1)
        End If
        rank = rst(2)
        surname = rst(4)
        firstname = rst(5)
        txtSurname = surname & ""
        txtRank = rank & ""
        txtChristian = firstname & ""
        txtBillet = billet & ""
        cmdOK_Click
        rst.MoveNext
    Loop
End Sub

Private Sub cmdOK_Click()
    txtSurname = txtSurname & " "
    txtChristian = txtChristian & " "
    txtRank = txtRank & " "
    txtBillet = txtBillet & " "
    If txtSurname <> "" Then
        PersonnelName = txtBillet & txtRank & txtChristian & txtSurname
        If chkOnboard.Value = vbChecked Then
            frmPersonnel.lstIn.AddItem PersonnelName
        Else
            frmPersonnel.lstOut.AddItem PersonnelName
        End If
    End If
    txtSurname = ""
    txtChristian = ""
    txtRank = ""
    txtBillet = ""
    txtBilletDesc = ""
    txtComment = ""
    subDoCheckBoxes
    lstMember.Clear
    txtSurname.SetFocus
End Sub

Private Sub cmdRemove_Click()
    If lstMember.ListIndex > -1 Then
        MsgBox "Move selected group from left hand list to right hand list"
    End If
End Sub

Private Sub Form_Load()
    subDoCheckBoxes
End Sub

Sub subDoCheckBoxes()
    If gLogOnboard = True Then
        chkOnboard.Value = vbChecked
        txtLocation = "ONBOARD"
    Else
        chkOnboard.Value = vbUnchecked
        txtLocation = "ASHORE"
    End If
End Sub
