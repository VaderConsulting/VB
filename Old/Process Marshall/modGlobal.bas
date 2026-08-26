Attribute VB_Name = "modGlobal"
Option Explicit

Public SQLServer As String
Public SQLDatabase As String
Public SQLDSN As String
Public TimerInterval As Integer
Public AgentID As Integer
Public ExclusiveMode As Boolean
Public LogSeverity As Integer

Public MaxProcesses As Integer
Public Const DefaultMaxProcesses = 5
Public Processes As Collection
Public TCPConns As Collection
Public Const MaxTCPConns = 3

Private RegisterAgentCounter As Long

'Public Sub BroadcastTCPMessage(ByVal s As String)
'    Dim oTCP As clsTCP
'    For Each oTCP In TCPConns
'        oTCP.SendBroadcast s
'        DoEvents
'    Next
'End Sub

Public Sub CloseObjects()
    'frmMain.tcpMain(0).Close
    CloseDatabase
    CloseLogDatabase
End Sub

Public Function DoubleDigit(ByVal i As Integer) As String
    Dim s As String
    s = CStr(i)
    Do While Len(s) < 2
        s = "0" + s
    Loop
    DoubleDigit = s
End Function

Public Function ExecuteCmd(Commandline As String, OwnerIP As String) As Boolean
    Dim oProc As clsProcess
    Set oProc = New clsProcess
    ExecuteCmd = False
    If Not ExecCmd(Commandline, oProc) Then
        WriteToLog 2, "Can't manually execute task (" & Commandline & ")."
        Set oProc = Nothing
        Exit Function
    End If
    oProc.DisplayName = "manually created Process - " & Commandline & "(" & OwnerIP & ")"
    Processes.Add oProc
    SetTasksState
    WriteToLog 3, "Task manually Spawned.  Client IP (" & OwnerIP & ") PID (" & oProc.dwProcessId & ")"
    ExecuteCmd = True
End Function

Public Function ExecuteTask(ByVal TaskID As Integer) As Boolean
    Dim Misc As Variant
    Dim TaskType As Integer
    Dim TaskParams As String
    Dim TaskPath As String
    Dim TaskDisplayName As String
    Dim xCmd As String
    Dim oProc As clsProcess
    
    ExecuteTask = False
    CloseDatabase
    OpenDatabase True
    
    'Check Current Status ... Should be set to 1
    Misc = GetStatus(TaskID)
    If (Misc <> 1) Then
        WriteToLog 2, "Can't fetch Task Status.  Task ID (" & TaskID & ")"
        CloseDatabase
        Exit Function
    End If
    
    If Not GetTaskInformation(TaskID, TaskType, TaskPath, TaskParams, TaskDisplayName) Then
        WriteToLog 2, "Can't fetch Task Information.  Task ID (" & TaskID & ")"
        CloseDatabase
        Exit Function
    End If
    
    ' Create Command Line
    xCmd = TaskPath + " " + TaskParams
    
    If (xCmd = "") Then
        SetStatus TaskID, -1
        WriteToLog 2, "Invalid Task Information.  Task ID (" & TaskID & ")"
        CloseDatabase
        Exit Function
    End If
    
    Set oProc = New clsProcess
    SetStatus TaskID, 0
    
    If Not ExecCmd(xCmd, oProc) Then
        WriteToLog 2, "Can't execute task (" & xCmd & ").  Task ID (" & TaskID & ")"
        CloseDatabase
        Set oProc = Nothing
        Exit Function
    End If
    WriteToLog 3, "Task Spawned.  Task ID (" & TaskID & ") PID (" & oProc.dwProcessId & ")"
    
    LogEntry TaskID, 1, "Task Started"
    If Not AssertAgentRunning(TaskID) Then WriteToLog 2, "Can't AssertAgentRunning.  Task ID (" & TaskID & ")"
    If Not ProcessSchedule(TaskID) Then WriteToLog 2, "Can't ProcessSchedule.  Task ID (" & TaskID & ")"
    
    oProc.TaskID = TaskID
    oProc.DisplayName = TaskDisplayName
    Processes.Add oProc
    CloseDatabase
    ExecuteTask = True
End Function

Public Function FindCompletedProcesses() As Boolean
    Dim oProc As clsProcess
    Dim i As Integer
    Dim DoAgain As Boolean
    
    FindCompletedProcesses = False
    
    Do
        DoAgain = False
        For i = 1 To Processes.Count
            Set oProc = Processes.Item(i)
            If ProcessCompleted(oProc) Then
                WriteToLog 3, "Process Stopped.  Task ID (" & oProc.TaskID & ") PID (" & oProc.dwProcessId & ")"
                LogEntry oProc.TaskID, 2, "Task Stopped"
                LogEntry oProc.TaskID, 3, GetProcessExitCode(oProc)
                
                OpenDatabase True
                If Not AssertAgentNotRunning(oProc.TaskID) Then WriteToLog 3, "Can't AssertAgentNotRunning.  Task ID (" & oProc.TaskID & ")"
                If Not AssertTaskComplete(oProc.TaskID) Then WriteToLog 3, "Can't AssertTaskComplete.  Task ID (" & oProc.TaskID & ")"
                CloseDatabase
                
                CleanupProcess oProc
                Processes.Remove i
                Set oProc = Nothing
                DoAgain = True
                Exit For
            End If
        Next
    Loop Until Not DoAgain
    
    FindCompletedProcesses = True
End Function

'Public Function FindTCPConnFromID(ByVal Index As Integer) As Integer
'    Dim oTCP As clsTCP
'    Dim i As Integer
'
'    FindTCPConnFromID = -1
'
'    For i = 1 To TCPConns.Count
'        Set oTCP = TCPConns.Item(i)
'        If (oTCP.ID = Index) Then FindTCPConnFromID = i
'    Next
'End Function

'Public Function GetValue(ByVal Hostname As String, ByVal Key As String, ByVal Value As String, ByRef ECode As Boolean) As String
'    Dim OpenKeyVal As Long
'    Dim OpenHiveVal As Long
'    Dim RResult As Long
'    Dim InfoTextStr As String
'
'    'Init
'    ECode = False
'    Hostname = Trim(Hostname)
'    Key = Trim(Key)
'    Value = Trim(Value)
'    GetValue = ""
'
'    RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
'    If (RResult <> ERROR_SUCCESS) Then Exit Function
'
'    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
'    InfoTextStr = RegistryQueryValue(OpenKeyVal, Value, REG_SZ)
'
'    RegCloseKey (OpenKeyVal)
'    RegCloseKey (OpenHiveVal)
'
'    GetValue = InfoTextStr
'    ECode = True
'End Function

Public Function InitObjects() As Boolean
    Dim Host As String
    Dim isOK As Boolean
    
    InitObjects = False
    WriteToLog 1, "Starting Agent..."
    
    Host = ""
    AgentID = -1
    ExclusiveMode = False
    MaxProcesses = DefaultMaxProcesses
    TimerInterval = 30
    SQLDSN = ""
    
    Set adoConn = New ADODB.Connection
    Set adoLogConn = New ADODB.Connection
    Set Processes = New Collection
    'Set TCPConns = New Collection
    
    App.StartLogging "", vbLogToNT
    Randomize Timer
    
    '*****  SQL Server
    SQLServer = "(Local)"
    'isOK = True
    'SQLServer = GetValue(Host, "Software\Tasks", "SQL Server", isOK)
    'If Not isOK Or (SQLServer = "") Then
    '    App.LogEvent "Invalid SQL Server", vbLogEventTypeError
    '    WriteToLog 1, "Invalid SQL Server"
    '    Exit Function
    'End If
    
    '*****  SQL Database
    SQLDatabase = "Domain Control"
    'SQLDatabase = GetValue(Host, "Software\DistroTasks", "SQL Database", isOK)
    'If Not isOK Or (SQLDatabase = "") Then
    '    App.LogEvent "Invalid SQL Database", vbLogEventTypeError
    '    WriteToLog 1, "Invalid SQL Database"
    '    Exit Function
    'End If
    SQLDSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & SQLDatabase & ";Data Source=" & SQLServer
    
    If StartTCPServer Then
        WriteToLog 1, "TCP Server Online"
    Else
        WriteToLog 1, "TCP Server FAILED"
    End If
    
    InitObjects = True
    WriteToLog 1, "Agent Started"
End Function

Public Function MinNZ(ByVal x As Integer, ByVal y As Integer) As Integer ' Weeds out 0
    If (x = 0) Then
        If (y = 0) Then
            MinNZ = 0
        Else
            MinNZ = y
        End If
    Else
        If (y = 0) Then
            MinNZ = x
        Else
            If (x < y) Then
                MinNZ = x
            Else
                MinNZ = y
            End If
        End If
    End If
End Function

Public Sub ProcessAgentKeepAlive()
    RegisterAgentCounter = RegisterAgentCounter - 1
    If (RegisterAgentCounter <= 0) Then
        RegisterAgentCounter = Int((5 * 60) / (frmMain.tmrWatchdog.Interval / 1000)) ' 5 Minutes
        AssertAgentKeepAlive
        WriteToLog 4, "Agent KeepAlive"
    End If
End Sub

Public Function ProcessListFull() As Integer
    ProcessListFull = (Processes.Count >= MaxProcesses)
End Function

Public Sub SetTasksState()
    'frmMain.lblTCPClients.Caption = CStr(TCPConns.Count) + " TCP Client/s"
    frmMain.txtNumProcesses = CStr(Processes.Count)
    If DoStop Then
        If (Processes.Count = 0) Then
            frmMain.txtState.Text = "Idle"
            WriteToLog 1, "Agent has stopped"
        Else
            frmMain.txtState.Text = "Stopping"
            WriteToLog 2, "Agent still stopping (" & Processes.Count & " Task/s Outstanding)"
        End If
    Else
        If DoPause Then
            frmMain.txtState.Text = "Paused"
            WriteToLog 1, "Agent has paused"
        Else
            frmMain.txtState.Text = "Running"
        End If
    End If
End Sub

Function ShowFreeSpace(drvPath, Optional xOption As Integer = 1)
    Dim fso, d, s As Long, t
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive(fso.GetDriveName(drvPath))
    Select Case xOption
        Case 1
            s = FormatNumber(d.FreeSpace / 1024 / 1024, 0) ' Mb
        Case 2
            s = FormatNumber(d.TotalSize / 1024 / 1024, 0) ' Mb
        Case 3
            s = FormatNumber(d.FreeSpace / 1024 / 1024, 0) ' Mb
            t = FormatNumber(d.TotalSize / 1024 / 1024, 0) ' Mb
            s = Int(((s / t) * 100) * 100) / 100 ' 2 decimal points
        Case 4
            s = FormatNumber(d.FreeSpace / 1024 / 1024, 0) ' Mb
            t = FormatNumber(d.TotalSize / 1024 / 1024, 0) ' Mb
            s = 100 - (Int(((s / t) * 100) * 100) / 100) ' 2 decimal points
    End Select
    ShowFreeSpace = s
End Function

Public Function StartTasks() As Boolean
    Dim isOK As Boolean
    Dim Misc As String
    Dim Host As String
    Dim MiscD As Double
    
    WriteToLog 2, "Agent going active..."
    StartTasks = False
    Host = ""
    MaxProcesses = DefaultMaxProcesses
    ExclusiveMode = False
    
    AgentID = RegisterAgent
    If (AgentID = -1) Then
        App.LogEvent "Invalid AgentID", vbLogEventTypeError
        WriteToLog 2, "Invalid AgentID"
        Exit Function
    End If
    
    OpenDatabase True
    FetchAgentConfiguration TimerInterval, MaxProcesses, ExclusiveMode
    CloseDatabase
    
    If LogDatabaseOpen Then CloseLogDatabase
    OpenLogDatabase
    frmMain.tmrMain.Interval = TimerInterval * 1000
    frmMain.tmrMain.Enabled = True
    SetTasksState
    
    StartTasks = True
    RegisterAgentCounter = 0
    WriteToLog 2, "Agent is active"
    DoStop = False
End Function

Private Function StartTCPServer() As Boolean
    On Error GoTo StartTCPServerError
    StartTCPServer = False
    
    'frmMain.tcpMain(0).LocalPort = 9999
    'frmMain.tcpMain(0).Listen
    
    StartTCPServer = True
    Exit Function
    
StartTCPServerError:
    Exit Function
End Function

Public Function StopTasks() As Boolean
    Dim i As Integer
    WriteToLog 2, "Agent stopping..."
    DoStop = True
    StopTasks = False
    
    'frmMain.tmrMain.Enabled = False
    '
    ' Kill each process we spawned
    '
    For i = 1 To Processes.Count
        KillProcess Processes.Item(i).dwProcessId
    Next i
    SetTasksState
    'If Processes.Count = 0 Then
    StopTasks = True
End Function

Public Sub WriteToLog(ByVal Sev As Integer, ByVal xStr As String)
    On Error GoTo WriteToLogError
    Dim s As String
    If (Sev > LogSeverity) Then Exit Sub
    s = Year(Now) & DoubleDigit(Month(Now)) & DoubleDigit(Day(Now)) & " " & DoubleDigit(Hour(Now)) & ":" & DoubleDigit(Minute(Now)) & ":" & DoubleDigit(Second(Now))
    s = s + " (" & Sev & ") " + xStr
    Open ".\TaskLog.txt" For Append As #1
    Print #1, s
    Close #1
    'BroadcastTCPMessage "L001:" + s + vbCrLf
    Exit Sub
WriteToLogError:
    Close #1
    Err.Clear
    Exit Sub
End Sub
