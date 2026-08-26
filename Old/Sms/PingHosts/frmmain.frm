VERSION 5.00
Object = "{B186D399-9B91-41CB-9242-E12B96CA99E1}#1.0#0"; "DRPing.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ping Hosts"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2760
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbrStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin Ping.drPing drPing1 
      Left            =   2400
      Top             =   0
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Label lblFinish 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Requires:
' Commandline, being the full path to domain.mdb
'
Const PingTimeout As Long = 500
Const MaxTimeout As Long = 4000
Dim DoUnload As Boolean


Private Sub Form_Load()
    Dim DSN As String
    Dim SQL As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim IP As String
    Dim PCName As String
    Dim NoOfRecords As Integer
    Dim i As Integer
    Dim StartTime As Date
    Dim FinishTime As Date
    Dim PercentDone As Double
    Dim ElapsedSeconds As Integer
    Dim PingResult As String
    Dim RecordID As Long
    Dim cmd As String
    
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    
    Me.Show
    Me.Refresh
    
    'If Command$ = "" Then
    '    cmd = InputBox("Enter path to Domain.mdb", "Input Required", App.Path & "\Domain.mdb")
    'Else
    '    cmd = Command$
    'End If
    
    Screen.MousePointer = vbHourglass
    
    'DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cmd & ";Persist Security Info=False"
    DSN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Domain Control;Data Source=(Local)"
    
    adoConn.Open DSN
    
    SQL = "SELECT * FROM tblHosts WHERE (IP <> '' OR IP = NULL) AND isUp = 0 ORDER BY Hostname"
    SQL = "SELECT * FROM tblHosts WHERE (IP <> '' OR IP = NULL) ORDER BY Hostname"
        
    adoRS.Open SQL, adoConn
    
    Do Until adoRS.EOF
        NoOfRecords = NoOfRecords + 1
        DoEvents
        adoRS.MoveNext
    Loop
    
    If NoOfRecords > 0 Then pbrStatus.Max = NoOfRecords
    
    StartTime = Now
    
    If Not adoRS.BOF Then adoRS.MoveFirst
    Do Until adoRS.EOF
        PCName = adoRS("Hostname")
        IP = adoRS("IP")
        RecordID = adoRS("ID")
        'PingTimeout = adoRS("Timeout")
        
        PingResult = drPing1.Ping(IP, PingTimeout)
        
        ' On the off chance the string returned is not in the form xxxxxx where x is any digit, make it a numeric.
        If Not IsNumeric(PingResult) Then PingResult = PingTimeout + 1
        
        ' Is the host up?  If the time returned was greater than the timeout specified, then we consider it as down
        If CLng(PingResult) > PingTimeout Then
            PingResult = drPing1.Ping(IP, MaxTimeout)
            If Not IsNumeric(PingResult) Then PingResult = MaxTimeout
            If CLng(PingResult) > MaxTimeout Then
                SQL = "UPDATE tblHosts SET isUp = 0, Timeout = " & MaxTimeout & ", [DateTime] = '" & Format(Now, "DD/MM/YYYY HH:NN:S") & "' WHERE id = " & RecordID
                Debug.Print "Timeout"
            Else
                SQL = "UPDATE tblHosts SET isUp = 1, Timeout = " & PingResult & ", [DateTime] = '" & Format(Now, "DD/MM/YYYY HH:NN:S") & "' WHERE id = " & RecordID
                Debug.Print "Updated Timeout"
            End If
        Else
            SQL = "UPDATE tblHosts SET isUp = 1, Timeout = " & PingResult & ", [DateTime] = '" & Format(Now, "DD/MM/YYYY HH:NN:S") & "' WHERE id = " & RecordID
            Debug.Print "Within Timeout"
        End If
        adoConn.Execute SQL
        
        i = i + 1
        lblStatus.Caption = PCName & " (" & IP & ")"
        pbrStatus.Value = i
        PercentDone = Round((i / NoOfRecords) * 100, 2) ' Round to 2 decimal places
        ElapsedSeconds = DateDiff("s", StartTime, Now)
        a = 100 / PercentDone
        totalseconds = a * ElapsedSeconds
        FinishTime = DateAdd("s", totalseconds, Now)
        lblFinish.Caption = "Finish at " & Format(FinishTime, "HH:MM AM/PM") & " (" & PercentDone & "% Done.)"
        DoEvents
        adoRS.MoveNext
        If DoUnload Then
            Screen.MousePointer = vbDefault
            Unload Me
            End
        End If
    Loop
    
    lblStatus.Caption = "Complete"
    
    Screen.MousePointer = vbDefault
    
    adoConn.Close
    
    Unload Me
    End
    
    Set adoRS = Nothing
    Set adoConn = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DoUnload = True
    Cancel = True
End Sub

