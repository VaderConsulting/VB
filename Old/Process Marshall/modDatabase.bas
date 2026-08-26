Attribute VB_Name = "modDatabase"
Option Explicit

Public adoConn As ADODB.Connection
Public adoLogConn As ADODB.Connection

Public Function AssertAgentKeepAlive() As Boolean
    On Error GoTo AssertAgentKeepAliveError
    AssertAgentKeepAlive = False
    
    If (AgentID < 0) Then Exit Function
    
    adoLogConn.Execute "UPDATE tblAgentReg SET LastRegistration = GETDATE() WHERE AgentID = " & AgentID
    
    AssertAgentKeepAlive = True
    Exit Function
AssertAgentKeepAliveError:
    Exit Function
End Function

Public Function FetchAgentConfiguration(ByRef TmrInt As Integer, ByRef MaxProc As Integer, ByRef ExMode As Boolean) As Boolean
    On Error GoTo FetchAgentConfigurationError
    FetchAgentConfiguration = False
    Dim adoRS As Recordset
    Set adoRS = New ADODB.Recordset
    
    adoRS.Open "SELECT * FROM tblAgentReg WHERE AgentID = " & AgentID, adoConn
    
    ' Timer Interval
    If Not IsNull(adoRS("TimerInterval")) Then TmrInt = adoRS("TimerInterval") Else TmrInt = 30
    
    ' MaxProcesses
    If Not IsNull(adoRS("MaxProcesses")) Then MaxProc = adoRS("MaxProcesses") Else MaxProc = DefaultMaxProcesses
    
    ' Exclusive Mode
    If Not IsNull(adoRS("ExclusiveMode")) Then ExMode = (adoRS("ExclusiveMode") <> 0) Else ExMode = False
    
    FetchAgentConfiguration = True
    Exit Function
FetchAgentConfigurationError:
    Exit Function
End Function

Public Function RegisterAgent() As Integer
    Dim MiscD As Double
    Dim adoRS As Recordset
    Dim RowsAffected As Integer
    Dim MiscI As Integer
    Dim isFound As Boolean
    
    On Error GoTo RegisterAgentError
    RegisterAgent = -1
    CloseDatabase
    If Not OpenDatabase(True) Then Exit Function
    
    Set adoRS = New ADODB.Recordset
    adoRS.Open "SELECT * FROM tblAgentReg WHERE Hostname = '" & UCase(Environ("ComputerName")) & "'", adoConn, 3

    If (adoRS.RecordCount > 1) Then ' Multiple Records Found
        adoRS.Close
        CloseDatabase
        Exit Function
    End If
    If (adoRS.RecordCount = 1) Then ' Existing Config Found
        MiscI = CInt(adoRS("AgentID"))
        If (MiscI > 0) Then
            MiscD = (Log(MiscI) / Log(2))
        Else
            MiscD = 0.5
        End If
        If (MiscD = Int(MiscD)) Then RegisterAgent = MiscI ' Valid AgentID
        adoRS.Close
        CloseDatabase
        Exit Function
    End If
    
    ' No Previous Record Found
    ' Parse tblDistroAgentReg searching for a new AgentID number
    
    adoRS.Close
    adoRS.Open "SELECT * FROM tblAgentReg ORDER BY AgentID", adoConn
    
    MiscI = 1
    isFound = False
    Do While Not adoRS.EOF And Not isFound
        If (Int(adoRS("AgentID")) <> MiscI) Then
            isFound = True
        Else
            MiscI = MiscI * 2
        End If
        adoRS.MoveNext
    Loop
    adoRS.Close
    adoConn.Execute "INSERT INTO tblAgentReg (AgentID,Hostname) VALUES(" & MiscI & ",'" & UCase(Environ("ComputerName")) & "')"
    RegisterAgent = MiscI
    
    CloseDatabase
    Exit Function
RegisterAgentError:
    RegisterAgent = -1
    Exit Function
End Function

Public Function OpenLogDatabase() As Boolean
    On Error GoTo OpenLogDatabaseError
    OpenLogDatabase = False
    adoLogConn.Open SQLDSN
OpenLogDatabaseError:
    OpenLogDatabase = True
    Exit Function
End Function

Public Function CloseLogDatabase() As Boolean
    On Error GoTo CloseLogDatabaseError
    CloseLogDatabase = False
    adoLogConn.Close
    Exit Function
CloseLogDatabaseError:
    Exit Function
End Function

Public Function LogDatabaseOpen() As Boolean
    LogDatabaseOpen = (adoLogConn.State = adStateOpen)
End Function

Public Function OpenDatabase(ByVal AllowRead As Boolean) As Boolean
    On Error GoTo OpenDatabaseError
    OpenDatabase = False
    If (adoConn.State = adStateOpen) Then Exit Function
    If AllowRead Then
        adoConn.Mode = adModeShareDenyNone
        adoConn.Open SQLDSN
    Else
        adoConn.Mode = adModeShareExclusive
        adoConn.Open SQLDSN
    End If
    
    OpenDatabase = True
    Exit Function
OpenDatabaseError:
    WriteToLog 0, "Error opening database!"
    Exit Function
End Function

Public Function CloseDatabase() As Boolean
    On Error GoTo CloseDatabaseError
    CloseDatabase = False
    If (adoConn.State <> adStateOpen) Then Exit Function
    adoConn.Close
    CloseDatabase = True
    Exit Function
CloseDatabaseError:
    Exit Function
End Function

Public Function GetStatus(ByVal TaskID As Integer) As Variant
    Dim adoRS As Recordset
    On Error GoTo GetStatusError
    GetStatus = -1
    
    Set adoRS = New ADODB.Recordset
    adoRS.Open "SELECT * FROM tblTasks WHERE TaskID = " & TaskID, adoConn, 3
    If (adoRS.RecordCount <> 1) Then
        adoRS.Close
        Exit Function
    End If
    
    GetStatus = adoRS("Status")
    adoRS.Close
    Exit Function
GetStatusError:
    GetStatus = -1
    Exit Function
End Function

Public Function LogEntry(ByVal TaskID As Long, ByVal EventEntry As Long, ByVal sText As String) As Boolean
    On Error GoTo LogError
    If LogDatabaseOpen Then
        adoLogConn.Execute "INSERT INTO tblLog (TaskID,Event,Result,EventDateTime,Agent) VALUES (" & TaskID & "," & EventEntry & ",'" & sText & "',DEFAULT," & AgentID & ")"
    End If
    LogEntry = True
    Exit Function
LogError:
    LogEntry = False
    WriteToLog 0, "Error writing to database log"
End Function

Public Function SetStatus(ByVal TaskID As Integer, ByVal NewStatus As Integer) As Boolean
    '
    ' Status Table
    ' 1 = Running
    '
    On Error GoTo SetStatusError
    SetStatus = False
    adoConn.Execute "UPDATE tblTasks SET Status = " & NewStatus & ", LastUpdate = GETDATE() WHERE TaskID = " & TaskID
    SetStatus = True
    Exit Function
SetStatusError:
    SetStatus = False
    Exit Function
End Function

Public Function GetTaskInformation(ByVal TaskID As Integer, ByRef TaskType As Integer, ByRef TaskPath As String, ByRef Params As String, ByRef DisplayName As String) As Boolean
    Dim adoRS As Recordset
    Dim SQL As String
    On Error GoTo GetTaskInformationError
    
    GetTaskInformation = False
    
    Set adoRS = New ADODB.Recordset
    
    SQL = "SELECT * FROM tblTasks INNER JOIN tblTaskPath ON "
    SQL = SQL + "tblTaskPath.Type = tblTasks.Type "
    SQL = SQL + "WHERE tblTasks.TaskID = " & TaskID
    
    adoRS.Open SQL, adoConn, 3
    If (adoRS.RecordCount <> 1) Then
        adoRS.Close
        Exit Function
    End If
    
    TaskType = adoRS("Type")
    If IsNull(adoRS("Params")) Then
        Params = ""
    Else
        Params = adoRS("Params")
    End If
    TaskPath = adoRS("TaskPath")
    DisplayName = adoRS("DisplayName")
    
    GetTaskInformation = True
    Exit Function
GetTaskInformationError:
    GetTaskInformation = False
    Exit Function
End Function
                                   
Public Function AssertTaskComplete(ByVal TaskID As Integer) As Boolean
    Dim adoRS As Recordset
    
    On Error GoTo AssertTaskCompleteError
    Set adoRS = New ADODB.Recordset
    
    AssertTaskComplete = False
    adoRS.Open "SELECT * FROM tblTasks WHERE NextDateTime IS NULL AND TaskID = " & TaskID, adoConn, 3
    If (adoRS.RecordCount <> 1) Then
        adoRS.Close
        Set adoRS = Nothing
        AssertTaskComplete = True
        Exit Function
    End If
    adoConn.Execute "UPDATE tblTasks SET Status = 1000 WHERE NextDateTime IS NULL AND TaskID = " & TaskID
    AssertTaskComplete = True
AssertTaskCompleteError:
    AssertTaskComplete = False
    Exit Function
End Function
                                   
Public Function AssertAgentRunning(ByVal TaskID As Integer) As Boolean
    On Error GoTo AssertAgentRunningError
    AssertAgentRunning = False
    adoConn.Execute "UPDATE tblTasks SET NumInstance = NumInstance + 1, AgentRunning = (AgentRunning | " & AgentID & ") WHERE TaskID = " & TaskID
    AssertAgentRunning = True
    Exit Function
AssertAgentRunningError:
    AssertAgentRunning = False
    Exit Function
End Function

Public Function AssertAgentNotRunning(ByVal TaskID As Integer) As Boolean
    On Error GoTo AssertAgentNotRunningError
    AssertAgentNotRunning = False
    adoConn.Execute "UPDATE tblTasks SET NumInstance = NumInstance - 1, AgentRunning = (AgentRunning & (-1 ^ " & AgentID & ")) WHERE TaskID = " & TaskID
    AssertAgentNotRunning = True
    Exit Function
AssertAgentNotRunningError:
    AssertAgentNotRunning = False
    Exit Function
End Function

Public Function ProcessSchedule(ByVal TaskID As Integer) As Boolean
    Dim adoRS As Recordset
    Dim SQL As String
    Dim SQL2 As String
    Dim xDateTime As Date
    Dim xSchedule As String
    Dim xNextDateTime As Date
    Dim xDelta As Integer
    
    On Error GoTo ProcessScheduleError
    ProcessSchedule = False
    
    Set adoRS = New ADODB.Recordset
    
    SQL = "SELECT * FROM tblTasks WHERE TaskID = " & TaskID
    adoRS.Open SQL, adoConn, 3
    If (adoRS.RecordCount <> 1) Then
        adoRS.Close
        Exit Function
    End If
    If Not IsNull(adoRS("NextDateTime")) Then
        xDateTime = adoRS("NextDateTime")
    Else
        xDateTime = 0
    End If
    xSchedule = CStr(adoRS("Schedule"))
    adoRS.Close
    
    'Check for Errors
    If (xSchedule = "") Then Exit Function
    If (xDateTime = 0) Then Exit Function
    
    SQL2 = ""
    'Parse Schedule Types
    ' Schedule is in the form: x,y, where x is the type of schedule, and y is the period between events
    If (xSchedule = "0") Then xNextDateTime = 0                    ' Single Time
    If (Left(xSchedule, 2) = "1,") Then                            ' Every x Seconds
        xDelta = CInt(Right(xSchedule, Len(xSchedule) - 2))
        SQL2 = "DATEADD(s," & xDelta & ",GETDATE())"
    End If
    If (Left(xSchedule, 2) = "2,") Then                            ' Every x Minutes
        xDelta = CInt(Right(xSchedule, Len(xSchedule) - 2))
        SQL2 = "DATEADD(n," & xDelta & ",GETDATE())"
    End If
    If (Left(xSchedule, 2) = "3,") Then                            ' Every x Days
        xDelta = CInt(Right(xSchedule, Len(xSchedule) - 2))
        SQL2 = "DATEADD(d," & xDelta & ",GETDATE())"
    End If
    
    SQL = "UPDATE tblTasks SET NextDateTime = "
    If (SQL2 = "") Then
        SQL = SQL + "NULL "
    Else
        SQL = SQL + SQL2
    End If
    SQL = SQL + " WHERE TaskID = " & TaskID
    adoConn.Execute SQL
    
    Set adoRS = Nothing
    
    ProcessSchedule = True
    Exit Function
ProcessScheduleError:
    ProcessSchedule = False
    Exit Function
End Function

Public Function FindATask(isTakeControl As Boolean) As Integer
    Dim adoRS As Recordset
    Dim SQL As String
    
    On Error GoTo FindATaskError
    FindATask = 0
    Set adoRS = New ADODB.Recordset
    
    CloseDatabase
    OpenDatabase False
    SQL = "SELECT TOP 1 * FROM tblTasks WHERE"
    SQL = SQL + " Status = 0 AND (NextDateTime <= GETDATE())"
    If (ExclusiveMode) Then
        SQL = SQL + " AND (AgentAllow = " & AgentID & ")"
    Else
        SQL = SQL + " AND (AgentAllow & " & AgentID & " <> 0)"
    End If
    'SQL = SQL + " AND (AgentRunning & " & AgentID & " = 0)"
    SQL = SQL + " AND (NumInstance < MaxInstance)"
    SQL = SQL + " ORDER BY NextDateTime DESC"
    adoRS.Open SQL, adoConn, adOpenForwardOnly
    
    If (adoRS.EOF) Then
        adoRS.Close
        CloseDatabase
        Exit Function
    End If
    If isTakeControl Then SetStatus adoRS("TaskID"), 1
    FindATask = adoRS("TaskID")
    adoRS.Close
    
    CloseDatabase
    Exit Function
FindATaskError:
    FindATask = 0
    Exit Function
End Function

'Public Function cSQLDateTime(ByVal Dte As Date) As String
'  Dim s As String
'  If Not IsDate(Dte) Then
'    cSQLDateTime = ""
'    Exit Function
'  End If
'  s = CStr(Month(Dte))
'  If (Len(s) < 2) Then s = "0" + s
'  If (Day(Dte) < 10) Then s = s + "0"
'  s = s + CStr(Day(Dte))
'  cSQLDateTime = CStr(Year(Dte)) + s + " " + CStr(DoubleDigit(Hour(Dte))) + ":" + CStr(DoubleDigit(Minute(Dte))) + ":" + CStr(DoubleDigit(Second(Dte)))
'End Function

'Function DoubleDigit(i)
'  Dim s As String
'  s = CStr(i)
'  While (Len(s) < 2)
'    s = "0" + s
'  Wend
'  DoubleDigit = s
'End Function

