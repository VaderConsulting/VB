VERSION 5.00
Object = "{B186D399-9B91-41CB-9242-E12B96CA99E1}#1.0#0"; "ping.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tasks"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9570
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid grdTargets 
      Height          =   6495
      Left            =   1920
      TabIndex        =   11
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   11456
      _Version        =   393216
      Rows            =   1000
      FixedRows       =   0
      FixedCols       =   0
      HighLight       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   ""
   End
   Begin Ping.drPing ocxPing 
      Left            =   7800
      Top             =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Task"
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox lstTasks 
      Height          =   1620
      Left            =   4560
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin VB.ListBox lstSitelocs 
      Height          =   6495
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblTaskInfo 
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   9
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Label lblTaskInfo 
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   8
      Top             =   2880
      Width           =   4815
   End
   Begin VB.Label lblTaskInfo 
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Label lblTaskInfo 
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   6
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label lblTasks 
      Alignment       =   2  'Center
      Caption         =   "Tasks"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   9375
   End
   Begin VB.Label lblOFOErrors 
      Alignment       =   2  'Center
      Caption         =   "Targets"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblTargets 
      Alignment       =   2  'Center
      Caption         =   "Sitelocs"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OFOPath As String
Dim oSchedule As Scheduler
Dim oTasks As Tasks
Dim oTask As Task


Private Sub cmdNew_Click()
    Open "\\" & lstTargets.List(lstTargets.ListIndex) & "\admin$\tasks\New Task.job" For Output As #1
        
    Close 1
End Sub

Private Sub Form_Load()
    ' Set up defaults
    Dim adoConn As ADODB.Connection
    Dim adoRs As ADODB.Recordset
    Dim ServerName As String
    grdTargets.ColWidth(0) = 600
    grdTargets.ColWidth(1) = 2000
    Me.Show
    Me.Refresh
    On Error Resume Next
    SQLServerName = "CBDXAAI"
    SQLDBName = "POLICE"
    OFOPath = "E$\BEOFO\"
    Set adoConn = New ADODB.Connection
    Set adoRs = New ADODB.Recordset
    
    DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & SQLDBName & " ;Data Source=" & SQLServerName
    adoConn.Open DSN
    
    SQL = "SELECT * FROM tblSiteloc "
    SQL = SQL & "ORDER BY Siteloc"
    adoRs.Open SQL, adoConn
    
    Do Until adoRs.EOF
        lstSitelocs.AddItem adoRs("Siteloc")
        adoRs.MoveNext
    Loop
    adoRs.Close
    'For lp = 0 To lstTargets.ListCount - 1
    '    ServerName = lstTargets.List(lp)
    '    UpdateStatus "Processing " & ServerName
    '    'Status = CheckforOFO(ServerName) ' True = OFO errors
    '    'If Status Then
    '    '    lstOFO.AddItem ServerName
     '   '    SQL = "INSERT INTO tblEvents (Hostname, Action, Initiated_By, Description, Complete, StartDateTime, ActionType) "
     '   '    SQL = SQL & "VALUES ('" & ServerName & "','BEOFO Diskspace','Automatic Monitoring','Check E:\BEOFO for excessive BEOFO files',0,'" & Format(Date, "yyyy") & Format(Date, "mm") & Format(Date, "dd") & " " & Format(Time, "hh") & ":" & Format(Time, "nn") & ":" & Format(Time, "ss") & "',3)"
     '   '    adoRs.Open SQL, adoConn
    ' '   'End If
    'Next lp
    adoRs.Close
    Set adoRs = Nothing
    adoConn.Close
    Set adoConn = Nothing
    'End
End Sub

Sub UpdateStatus(Message)
    lblStatus = Message
    lblStatus.Refresh
End Sub

Function CheckforOFO(ServerName As String) As Boolean
    On Error GoTo CantConnect
    CheckforOFO = False
    DirPath = "\\" & ServerName & "\" & OFOPath
    d = Dir(DirPath & "\*.sss")
    If d > "" Then
        d = Dir
        If d > "" Then CheckforOFO = True
    End If
    Exit Function
CantConnect:
    Err.Clear
    lstOFO.AddItem ServerName & " (Unknown)"
End Function

Private Sub grdTargets_Click()
    ClearInfo
    grdTargets.Col = 1
    TaskName = grdTargets.Text
    lblStatus = "Retrieving info for  " & TaskName
    lblStatus.Refresh
    'Set oTask = oSchedule.Tasks(1)
    'lblTaskInfo(0) = oTask.JobName
    'lblTaskInfo(1) = oTask.ApplicationName
    'lblTaskInfo(2) = oTask.Parameters
    'lblTaskInfo(3) = oTask.Status
    lblStatus = "Idle"
    lblStatus.Refresh
End Sub

Private Sub lstSitelocs_Click()
    Dim adoConn As ADODB.Connection
    Dim adoRs As ADODB.Recordset
    Dim ServerName As String, t As String, e As Long
    grdTargets.Clear
    On Error Resume Next
    SQLServerName = "CBDXAAI"
    SQLDBName = "POLICE"
    Set adoConn = New ADODB.Connection
    Set adoRs = New ADODB.Recordset
    
    DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & SQLDBName & " ;Data Source=" & SQLServerName
    adoConn.Open DSN
    
    SQL = "SELECT * FROM tblNodes WHERE Siteloc = '" & lstSitelocs.List(lstSitelocs.ListIndex) & "' AND (Type & 2)=2 "
    'SQL = "SELECT * FROM tblNodes WHERE Siteloc = '" & lstSitelocs.List(lstSitelocs.ListIndex) & "' AND (Type & 1)=1 "
    SQL = SQL & "ORDER BY Hostname"
    adoRs.Open SQL, adoConn
    y = 0
    Do Until adoRs.EOF
        'lstTargets.AddItem adoRs("Hostname")
        grdTargets.Row = y
        grdTargets.Col = 1
        grdTargets.Text = adoRs("Hostname")
        adoRs.MoveNext
        y = y + 1
    Loop
    adoRs.Close
    frmMain.Refresh
    
    For lp = 0 To y - 1
        grdTargets.Col = 1
        grdTargets.Row = lp
        H = Trim(grdTargets.Text)
        t = ocxPing.Ping(H, 1000)
        e = ocxPing.ReturnCode
        If t <> "" Then
            grdTargets.Col = 0
            Debug.Print H & " returns " & t & " (returncode " & e & ")"
            If t <> 1048596 Then
                grdTargets.Text = "Up"
            Else
                grdTargets.Text = "Down"
            End If
        End If
        frmMain.Refresh
    Next lp
    adoRs.Close
    Set adoRs = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub lstTargets_Click()
    Set oSchedule = New TASKSCHEDULERLib.Scheduler
    lstTasks.Clear
    ClearInfo
    On Error GoTo ErrorHandler
    TargetName = lstTargets.List(lstTargets.ListIndex)
    lblStatus = "Contacting " & TargetName
    lblStatus.Refresh
    oSchedule.TargetComputer = "\\" & TargetName
    lblStatus = "Enumerating tasks on " & TargetName
    lblStatus.Refresh
    For Each oTask In oSchedule.Tasks
        lstTasks.AddItem oTask.JobName
    Next
    lblStatus = "Idle"
    lblStatus.Refresh
    Exit Sub
ErrorHandler:
    If Err.Number = -2147024843 Then
        lblStatus = "Computer not responding"
        lblStatus.Refresh
    End If
End Sub

Sub ClearInfo()
    For lp = lblTaskInfo().LBound To lblTaskInfo().UBound
        lblTaskInfo(lp) = ""
    Next lp
End Sub
