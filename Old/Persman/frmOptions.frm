VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   7830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdgBrowse 
      Left            =   6016
      Top             =   3072
      _ExtentX        =   926
      _ExtentY        =   926
      _Version        =   327681
   End
   Begin VB.CommandButton cmdFinished 
      Caption         =   "OK"
      Height          =   375
      Left            =   6656
      TabIndex        =   0
      Top             =   3712
      Width           =   784
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   4240
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7824
      _ExtentX        =   13811
      _ExtentY        =   7488
      _Version        =   327681
      Tab             =   1
      TabHeight       =   580
      TabCaption(0)   =   "File Paths"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblOrderPath"
      Tab(0).Control(1)=   "lblDbPath"
      Tab(0).Control(2)=   "lblLog"
      Tab(0).Control(3)=   "lblWordViewPath"
      Tab(0).Control(4)=   "txtOrdersPath"
      Tab(0).Control(5)=   "txtDbPath"
      Tab(0).Control(6)=   "txtLogfile"
      Tab(0).Control(7)=   "txtWordViewPath"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblOption(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblOption(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblOption(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkNew"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkID"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmbAction"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdDoit(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdDoit(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkLog"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdView"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Times"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeHours"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtWordViewPath 
         Height          =   285
         Left            =   -71760
         TabIndex        =   32
         Top             =   2400
         Width           =   4335
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View..."
         Height          =   400
         Left            =   6656
         TabIndex        =   31
         Top             =   3072
         Width           =   784
      End
      Begin VB.CheckBox chkLog 
         Caption         =   "Create activity logfile"
         Height          =   272
         Left            =   256
         TabIndex        =   30
         Top             =   3072
         Width           =   3472
      End
      Begin VB.TextBox txtLogfile 
         Height          =   304
         Left            =   -71800
         TabIndex        =   29
         Top             =   1792
         Width           =   4368
      End
      Begin VB.Frame fmeHours 
         Caption         =   "Work Days"
         Height          =   2448
         Left            =   -74744
         TabIndex        =   14
         Top             =   512
         Width           =   7184
         Begin VB.CheckBox chkDay 
            Caption         =   "Sun"
            Height          =   272
            Index           =   0
            Left            =   128
            TabIndex        =   24
            Top             =   384
            Width           =   784
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Mon"
            Height          =   272
            Index           =   1
            Left            =   896
            TabIndex        =   23
            Top             =   384
            Value           =   1  'Checked
            Width           =   784
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Tue"
            Height          =   272
            Index           =   2
            Left            =   1664
            TabIndex        =   22
            Top             =   384
            Value           =   1  'Checked
            Width           =   784
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Wed"
            Height          =   272
            Index           =   3
            Left            =   2432
            TabIndex        =   21
            Top             =   384
            Value           =   1  'Checked
            Width           =   784
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Thu"
            Height          =   272
            Index           =   4
            Left            =   3200
            TabIndex        =   20
            Top             =   384
            Value           =   1  'Checked
            Width           =   784
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Fri"
            Height          =   272
            Index           =   5
            Left            =   3968
            TabIndex        =   19
            Top             =   384
            Value           =   1  'Checked
            Width           =   784
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "Sat"
            Height          =   272
            Index           =   6
            Left            =   4736
            TabIndex        =   18
            Top             =   384
            Width           =   784
         End
         Begin VB.TextBox txtStartJS 
            Height          =   304
            Left            =   2176
            TabIndex        =   17
            Text            =   "0750"
            Top             =   768
            Width           =   528
         End
         Begin VB.TextBox txtStartSS 
            Height          =   304
            Left            =   2176
            TabIndex        =   16
            Text            =   "0755"
            Top             =   1152
            Width           =   528
         End
         Begin VB.TextBox txtStartOfficers 
            Height          =   304
            Left            =   2176
            TabIndex        =   15
            Text            =   "0755"
            Top             =   1536
            Width           =   528
         End
         Begin VB.Label lblStartJS 
            Caption         =   "Start time - Junior Sailors"
            Height          =   272
            Left            =   128
            TabIndex        =   27
            Top             =   768
            Width           =   1936
         End
         Begin VB.Label lblStartSS 
            Caption         =   "Start time - Senior Sailors"
            Height          =   272
            Left            =   128
            TabIndex        =   26
            Top             =   1152
            Width           =   2192
         End
         Begin VB.Label lblStartOfficers 
            Caption         =   "Start time - Officers"
            Height          =   272
            Left            =   128
            TabIndex        =   25
            Top             =   1536
            Width           =   2192
         End
      End
      Begin VB.TextBox txtDbPath 
         Height          =   304
         Left            =   -71800
         TabIndex        =   12
         Top             =   1152
         Width           =   4368
      End
      Begin VB.TextBox txtOrdersPath 
         Height          =   304
         Left            =   -71800
         TabIndex        =   10
         Top             =   512
         Width           =   4368
      End
      Begin VB.CommandButton cmdDoit 
         Caption         =   "Set"
         Height          =   375
         Index           =   1
         Left            =   6656
         TabIndex        =   9
         Top             =   1024
         Width           =   855
      End
      Begin VB.CommandButton cmdDoit 
         Caption         =   "Set"
         Height          =   368
         Index           =   0
         Left            =   6656
         TabIndex        =   8
         Top             =   512
         Width           =   855
      End
      Begin VB.ComboBox cmbAction 
         Height          =   315
         ItemData        =   "frmOptions.frx":0054
         Left            =   5888
         List            =   "frmOptions.frx":0056
         TabIndex        =   7
         Top             =   1536
         Width           =   1695
      End
      Begin VB.CheckBox chkID 
         Caption         =   "Create ID's for new user's"
         Height          =   255
         Left            =   256
         TabIndex        =   4
         Top             =   2048
         Width           =   3255
      End
      Begin VB.CheckBox chkNew 
         Caption         =   "Log new personnel onboard upon creation"
         Height          =   272
         Left            =   272
         TabIndex        =   3
         Top             =   2560
         Width           =   3728
      End
      Begin VB.Label lblWordViewPath 
         Caption         =   "Default filename for Wordview path"
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblLog 
         Caption         =   "Default filename for logfile"
         Height          =   272
         Left            =   -74744
         TabIndex        =   28
         Top             =   1792
         Width           =   2704
      End
      Begin VB.Label lblDbPath 
         Caption         =   "Default path for Location database"
         Height          =   272
         Left            =   -74744
         TabIndex        =   13
         Top             =   1152
         Width           =   2832
      End
      Begin VB.Label lblOrderPath 
         Caption         =   "Default path for Daily Orders"
         Height          =   272
         Left            =   -74744
         TabIndex        =   11
         Top             =   512
         Width           =   2832
      End
      Begin VB.Label lblOption 
         Caption         =   "Log all personnel ashore now"
         Height          =   256
         Index           =   1
         Left            =   272
         TabIndex        =   6
         Top             =   1024
         Width           =   3632
      End
      Begin VB.Label lblOption 
         Caption         =   "Default double-click action..."
         Height          =   256
         Index           =   2
         Left            =   256
         TabIndex        =   5
         Top             =   1536
         Width           =   3488
      End
      Begin VB.Label lblOption 
         Caption         =   "Log all personnel onboard now"
         Height          =   256
         Index           =   0
         Left            =   256
         TabIndex        =   2
         Top             =   512
         Width           =   3632
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkLog_Click()
    If chkLog.Value = vbChecked Then
        gLog = True
        txtLogfile.Enabled = True
    Else
        gLog = False
        txtLogfile.Enabled = False
    End If

End Sub

Private Sub cmdDoit_Click(Index As Integer)
    Select Case Index
    Case 0
        If frmPersonnel.lstOut.ListCount > 0 Then
            For a = 0 To frmPersonnel.lstOut.ListCount - 1
                frmPersonnel.lstIn.AddItem frmPersonnel.lstOut.List(0)
                frmPersonnel.lstOut.RemoveItem 0
            Next a
        End If
    Case 1
        If frmPersonnel.lstIn.ListCount > 0 Then
            For a = 0 To frmPersonnel.lstIn.ListCount - 1
                frmPersonnel.lstOut.AddItem frmPersonnel.lstIn.List(0)
                frmPersonnel.lstIn.RemoveItem 0
            Next a
        End If
    End Select
End Sub

Private Sub cmdFinished_Click()
    If chkID.Value = vbChecked Then
        gAutoMakeID = True
    Else
        gAutoMakeID = False
    End If
    If chkNew.Value = vbChecked Then
        gLogOnboard = True
    Else
        gLogOnboard = False
    End If
    gDefaultAction = cmbAction.List(cmbAction.ListIndex)
    gOrderPath = txtOrdersPath
    gDbPath = txtDbPath
    gLogFile = txtLogfile
    gWordViewPath = txtWordViewPath
    Unload Me
End Sub

Private Sub cmdView_Click()
    frmLog.Show vbModal
End Sub

Private Sub Form_Load()
    Height = tabOptions.Height
    cmbAction.Clear
    cmbAction.AddItem "Mail"
    cmbAction.AddItem "In/Out"
    cmbAction.AddItem "Properties"
    cmbAction = gDefaultAction
    txtOrdersPath = gOrderPath
    txtDbPath = gDbPath
    txtWordViewPath = gWordViewPath
    If gAutoMakeID = True Then
        chkID.Value = vbChecked
    Else
        chkID.Value = vbUnchecked
    End If
    If gLogOnboard = True Then
        chkNew.Value = vbChecked
    Else
        chkNew.Value = vbUnchecked
    End If
    If gLog = True Then
        chkLog.Value = vbChecked
        txtLogfile.Enabled = True
    Else
        chkLog.Value = vbUnchecked
        txtLogfile.Enabled = False
    End If
    
End Sub

