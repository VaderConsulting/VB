VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Check"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6345
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelServer 
      Caption         =   "Del"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   18
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdAddServer 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox txtAddServer 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Del"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   13
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Del"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   11
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtAdd 
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtAdd 
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   8
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Text            =   "\\perthxsaa\stats"
      Top             =   6360
      Width           =   1815
   End
   Begin VB.ListBox lstTargets 
      Height          =   2010
      Index           =   1
      ItemData        =   "frmMain.frx":030A
      Left            =   2640
      List            =   "frmMain.frx":030C
      TabIndex        =   4
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CheckBox chkCheck 
      Caption         =   "dir *.[type]"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   5640
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkCheck 
      Caption         =   "diruse /m /*"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.ListBox lstTargets 
      Height          =   2010
      Index           =   0
      ItemData        =   "frmMain.frx":030E
      Left            =   2640
      List            =   "frmMain.frx":0310
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox lstServers 
      Height          =   5325
      ItemData        =   "frmMain.frx":0312
      Left            =   120
      List            =   "frmMain.frx":0314
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmMain.frx":0316
      Height          =   855
      Left            =   600
      TabIndex        =   20
      Top             =   6720
      Width           =   5055
   End
   Begin VB.Label lblSelect 
      Caption         =   "Select server(s) to target:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   7920
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblTaskList 
      Alignment       =   2  'Center
      Caption         =   "TaskList 2"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   15
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblTaskList 
      Alignment       =   2  'Center
      Caption         =   "TaskList 1"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   14
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblAddOutputPath 
      Caption         =   "\%computername%"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   6390
      Width           =   1575
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output path:"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh Serverlist"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Create .bat file(s)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkCheck_Click(Index As Integer)
    If chkCheck(Index).Value = vbChecked Then
        lstTargets(Index).Enabled = True
    Else
        lstTargets(Index).Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    If Index = 0 Then Text = "diruse /m /* \\[server]\" & txtAdd(0)
    If Index = 1 Then Text = "dir \\[server]\e$\*." & txtAdd(1) & " /s"
    lstTargets(Index).AddItem Text
    txtAdd(Index) = ""
    txtAdd(Index).SetFocus
End Sub


Private Sub cmdDel_Click(Index As Integer)
    If lstTargets(Index).ListCount > 0 Then
        lstTargets(Index).RemoveItem lstTargets(Index).ListIndex
    End If
End Sub

Private Sub cmdDelServer_Click()
    lstServers.RemoveItem lstServers.ListIndex
    lstServers_Click
End Sub

Private Sub Form_Load()
    lstTargets(0).AddItem "START diruse /m /* \\[server]\e$"
    lstTargets(0).AddItem "START diruse /m /* \\[server]\e$\users"
    
    lstTargets(1).AddItem "START dir \\[server]\e$\*.mp3 /s"
    lstTargets(1).AddItem "START dir \\[server]\e$\*.jpg /s"
    lstTargets(1).AddItem "START dir \\[server]\e$\*.gif /s"
    
    Randomize Timer
End Sub

Function makerndname() As String
    For lp = 1 To 8
        makerndname = makerndname & Chr((Rnd(1) * 26) + 65)
    Next lp
End Function

Private Sub lstServers_Click()
    If ListIndex <> -1 Then
        cmdDelServer.Enabled = True
        mnuStart.Enabled = True
    Else
        cmdDelServer.Enabled = False
        mnuStart.Enabled = False
    End If
End Sub

Private Sub lstTargets_Click(Index As Integer)
    If lstTargets(Index).ListIndex <> -1 Then
        cmdDel(Index).Enabled = True
    Else
        cmdDel(Index).Enabled = False
    End If
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuRefresh_Click()
    Dim adoConn As ADODB.Connection, DSN As String, SQL As String, adoData As ADODB.Recordset
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoData = CreateObject("ADODB.Recordset")
    
    lstServers.Clear
    DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=CBDXAAI"
    adoConn.Open DSN
    SQL = "SELECT * FROM tblServer WHERE Supported = '1'" ' AND PrimaryRole = 'File and Print'"
    adoData.Open SQL, adoConn
    Do Until adoData.EOF
        lstServers.AddItem adoData.Fields(0)
        adoData.MoveNext
        'UpdateDestination
    Loop
    adoConn.Close
    Set adoConn = Nothing

End Sub

Private Sub mnuStart_Click()
    Start
End Sub

Sub Start()
    For lp = 0 To lstServers.ListCount - 1
        If lstServers.Selected(lp) = True Then
            Server = lstServers.List(lp)
            If chkCheck(0).Value = vbChecked Then
                For i = 0 To lstTargets(0).ListCount - 1
                    cmd = lstTargets(0).List(i)
                    cmd = Replace(cmd, "[server]", Server)
                    cmd = cmd & " >" & txtOutput & "\" & Server & "_tasklist1" & "_subtask" & i + 1 & " " & Format(Date, "DD-MM-YY") & ".txt"
                    'Debug.Print cmd
                    'batname = makerndname & ".bat"
                    batname = Server & ".bat"
                    Open "c:\temp\" & batname For Output As #1
                        Print #1, cmd
                    Close 1
                    'Shell "c:\temp\" & batname
                Next i
            End If
            If chkCheck(1).Value = vbChecked Then
                For i = 0 To lstTargets(1).ListCount - 1
                    cmd = lstTargets(1).List(i)
                    cmd = Replace(cmd, "[server]", Server)
                    cmd = cmd & " >" & txtOutput & "\" & Server & "_tasklist2" & "_subtask" & i + 1 & " " & Format(Date, "DD-MM-YY") & ".txt"
                    'Debug.Print cmd
                    'batname = makerndname & ".bat"
                    Open "c:\temp\" & batname For Append As #1
                        Print #1, cmd
                    Close 1
                    'Shell "c:\temp\" & batname
                Next i
            End If
        End If
    Next lp
    MsgBox "Task complete.  Review the c:\temp directory for the bat files created, and copy to the selected server(s)."
End Sub

Private Sub txtAddServer_Change()
    If txtAddServer <> "" Then
        cmdAddServer.Enabled = True
    Else
        cmdAddServer.Enabled = False
    End If
End Sub
