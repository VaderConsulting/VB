Attribute VB_Name = "modDatabase"
Option Explicit

Public DBConn As ADODB.Connection
Public LogConn As ADODB.Connection

Public Function AssertAgentKeepAlive() As Boolean
  On Error GoTo AssertAgentKeepAliveError
  AssertAgentKeepAlive = False
  
  If (AgentID < 0) Then Exit Function
  
  LogConn.Execute "UPDATE tblDistroAgentReg SET LastRegistration = GETDATE() WHERE AgentID = " & AgentID
  
  AssertAgentKeepAlive = True
  Exit Function
AssertAgentKeepAliveError:
  Exit Function
End Function

Public Function FetchAgentConfiguration(ByRef TmrInt As Integer, _
                                        ByRef MaxProc As Integer, _
                                        ByRef ExMode As Boolean) As Boolean
                                      
  On Error GoTo FetchAgentConfigurationError
  FetchAgentConfiguration = False
  Dim CSCDB1 As Recordset
  Set CSCDB1 = New ADODB.Recordset
                                   
  CSCDB1.Open "SELECT * FROM tblDistroAgentReg WHERE AgentID = " & AgentID, DBConn
  
  ' Timer Interval
  If Not IsNull(CSCDB1("TimerInterval")) Then TmrInt = CSCDB1("TimerInterval") Else TmrInt = 30
  
  ' MaxProcesses
  If Not IsNull(CSCDB1("MaxProcesses")) Then MaxProc = CSCDB1("MaxProcesses") Else MaxProc = DefaultMaxProcesses
  
  ' Exclusive Mode
  If Not IsNull(CSCDB1("ExclusiveMode")) Then ExMode = (CSCDB1("ExclusiveMode") <> 0) Else ExMode = False

  FetchAgentConfiguration = True
  Exit Function
FetchAgentConfigurationError:
  Exit Function
End Function

Public Function RegisterAgent() As Integer
  Dim MiscD As Double
  Dim CSCDB1 As Recordset
  Dim RowsAffected As Integer
  Dim MiscI As Integer
  Dim isFound As Boolean
  
  On Error GoTo RegisterAgentError
  RegisterAgent = -1
  CloseDatabase
  If Not OpenDatabase(True) Then Exit Function
  
  Set CSCDB1 = New ADODB.Recordset
  CSCDB1.Open "SELECT * FROM tblDistroAgentReg WHERE Hostname = '" & UCase(Environ("ComputerName")) & "'", DBConn, 3
  
  If (CSCDB1.RecordCount > 1) Then ' Multiple Records Found
    CSCDB1.Close
    CloseDatabase
    Exit Function
  End If
  If (CSCDB1.RecordCount = 1) Then ' Existing Config Found
    MiscI = CInt(CSCDB1("AgentID"))
    If (MiscI > 0) Then
      MiscD = (Log(MiscI) / Log(2))
    Else
      MiscD = 0.5
    End If
    If (MiscD = Int(MiscD)) Then RegisterAgent = MiscI ' Valid AgentID
    CSCDB1.Close
    CloseDatabase
    Exit Function
  End If
  
  ' No Previous Record Found
  ' Parse tblDistroAgentReg searching for a new AgentID number
  
  CSCDB1.Close
  CSCDB1.Open "SELECT * FROM tblDistroAgentReg ORDER BY AgentID", DBConn
  
  MiscI = 1
  isFound = False
  Do While Not CSCDB1.EOF And Not isFound
    If (Int(CSCDB1("AgentID")) <> MiscI) Then
      isFound = True
    Else
      MiscI = MiscI * 2
    End If
    CSCDB1.MoveNext
  Loop
  CSCDB1.Close
  DBConn.Execute "INSERT INTO tblDistroAgentReg (AgentID,Hostname) VALUES(" & MiscI & ",'" & UCase(Environ("ComputerName")) & "')"
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
  LogConn.Open SQLDSN
OpenLogDatabaseError:
  OpenLogDatabase = True
  Exit Function
End Function

Public Function CloseLogDatabase() As Boolean
  On Error GoTo CloseLogDatabaseError
  CloseLogDatabase = False
  LogConn.Close
  Exit Function
CloseLogDatabaseError:
  Exit Function
End Function

Public Function LogDatabaseOpen() As Boolean
  LogDatabaseOpen = (LogConn.State = adStateOpen)
End Function

Public Function OpenDatabase(ByVal AllowRead As Boolean) As Boolean
  On Error GoTo OpenDatabaseError
  OpenDatabase = False
  If (DBConn.State = adStateOpen) Then Exit Function
  If AllowRead Then
    DBConn.Mode = adModeShareDenyNone
    DBConn.Open SQLDSN
  Else
    DBConn.Mode = adModeShareExclusive
    DBConn.Open SQLDSN
  End If
  
  OpenDatabase = True
  Exit Function
OpenDatabaseError:
  Exit Function
End Function

Public Function CloseDatabase() As Boolean
  On Error GoTo CloseDatabaseError
  CloseDatabase = False
  If (DBConn.State <> adStateOpen) Then Exit Function
  DBConn.Close
  CloseDatabase = True
  Exit Function
CloseDatabaseError:
  Exit Function
End Function

Public Function GetStatus(ByVal TaskID As Integer) As Variant
  Dim CSCDB1 As Recordset
  On Error GoTo GetStatusError
  GetStatus = -1
  
  Set CSCDB1 = New ADODB.Recordset
  CSCDB1.Open "SELECT * FROM tblDistroTasks WHERE TaskID = " & TaskID, DBConn, 3
  If (CSCDB1.RecordCount <> 1) Then
    CSCDB1.Close
    Exit Function
  End If
  
  GetStatus = CSCDB1("Status")
  CSCDB1.Close
  Exit Function
GetStatusError:
  GetStatus = -1
  Exit Function
End Function

Public Function LogEntry(ByVal TaskID As Long, ByVal EventEntry As Long, ByVal sText As String)
  If LogDatabaseOpen Then
    LogConn.Execute "INSERT INTO tblDistroLog (TaskID,Event,Result,EventDateTime,Agent) VALUES (" & TaskID & "," & EventEntry & ",'" & sText & "',DEFAULT," & AgentID & ")"
  End If
End Function

Public Function SetStatus(ByVal TaskID As Integer, ByVal NewStatus As Integer) As Boolean
  On Error GoTo SetStatusError
  SetStatus = False
  DBConn.Execute "UPDATE tblDistroTasks SET Status = " & NewStatus & ", LastUpdate = GETDATE() WHERE TaskID = " & TaskID
  SetStatus = True
  Exit Function
SetStatusError:
  SetStatus = False
  Exit Function
End Function

Public Function GetTaskInformation(ByVal TaskID As Integer, _
                                   ByRef TaskType As Integer, _
                                   ByRef TaskPath As String, _
                                   ByRef Params As String, _
                                   ByRef DisplayName As String) As Boolean
  Dim CSCDB1 As Recordset
  Dim SQL As String
  On Error GoTo GetTaskInformationError
  
  GetTaskInformation = False
  
  Set CSCDB1 = New ADODB.Recordset
  
  SQL = "SELECT * FROM tblDistroTasks INNER JOIN tblDistroTaskPath ON "
  SQL = SQL + "tblDistroTaskPath.Type = tblDistroTasks.Type "
  SQL = SQL + "WHERE tblDistroTasks.TaskID = " & TaskID
  
  CSCDB1.Open SQL, DBConn, 3
  If (CSCDB1.RecordCount <> 1) Then
    CSCDB1.Close
    Exit Function
  End If
  
  TaskType = CSCDB1("Type")
  If IsNull(CSCDB1("Params")) Then
    Params = ""
  Else
    Params = CSCDB1("Params")
  End If
  TaskPath = CSCDB1("TaskPath")
  DisplayName = CSCDB1("DisplayName")
  
  GetTaskInformation = True
  Exit Function
GetTaskInformationError:
  GetTaskInformation = False
  Exit Function
End Function
                                   
Public Function AssertTaskComplete(ByVal TaskID As Integer) As Boolean
  Dim CSCDB1 As Recordset
  
  On Error GoTo AssertTaskCompleteError
  Set CSCDB1 = New ADODB.Recordset
  
  AssertTaskComplete = False
  CSCDB1.Open "SELECT * FROM tblDistroTasks WHERE NextDateTime IS NULL AND TaskID = " & TaskID, DBConn, 3
  If (CSCDB1.RecordCount <> 1) Then
    CSCDB1.Close
    Set CSCDB1 = Nothing
    AssertTaskComplete = True
    Exit Function
  End If
  DBConn.Execute "UPDATE tblDistroTasks SET Status = 1000 WHERE NextDateTime IS NULL AND TaskID = " & TaskID
  AssertTaskComplete = True
AssertTaskCompleteError:
  AssertTaskComplete = False
  Exit Function
End Function
                                   
Public Function AssertAgentRunning(ByVal TaskID As Integer) As Boolean
  On Error GoTo AssertAgentRunningError
  AssertAgentRunning = False
  DBConn.Execute "UPDATE tblDistroTasks SET NumInstance = NumInstance + 1, AgentRunning = (AgentRunning | " & AgentID & ") WHERE TaskID = " & TaskID
  AssertAgentRunning = True
  Exit Function
AssertAgentRunningError:
  AssertAgentRunning = False
  Exit Function
End Function

Public Function AssertAgentNotRunning(ByVal TaskID As Integer) As Boolean
  On Error GoTo AssertAgentNotRunningError
  AssertAgentNotRunning = False
  DBConn.Execute "UPDATE tblDistroTasks SET NumInstance = NumInstance - 1, AgentRunning = (AgentRunning & (-1 ^ " & AgentID & ")) WHERE TaskID = " & TaskID
  AssertAgentNotRunning = True
  Exit Function
AssertAgentNotRunningError:
  AssertAgentNotRunning = False
  Exit Function
End Function

Public Function ProcessSchedule(ByVal TaskID As Integer) As Boolean
  Dim CSCDB1 As Recordset
  Dim SQL As String
  Dim SQL2 As String
  Dim xDateTime As Date
  Dim xSchedule As String
  Dim xNextDateTime As Date
  Dim xDelta As Integer
  
  On Error GoTo ProcessScheduleError
  ProcessSchedule = False
  
  Set CSCDB1 = New ADODB.Recordset
  
  SQL = "SELECT * FROM tblDistroTasks WHERE TaskID = " & TaskID
  CSCDB1.Open SQL, DBConn, 3
  If (CSCDB1.RecordCount <> 1) Then
    CSCDB1.Close
    Exit Function
  End If
  If Not IsNull(CSCDB1("NextDateTime")) Then
    xDateTime = CSCDB1("NextDateTime")
  Else
    xDateTime = 0
  End If
  xSchedule = CSCDB1("Schedule")
  CSCDB1.Close
  
  'Check for Errors
  If (xSchedule = "") Then Exit Function
  If (xDateTime = 0) Then Exit Function
  
  SQL2 = ""
  'Parse Schedule Types
  If (xSchedule = "0") Then xNextDateTime = 0 ' Single Time
  If (Left(xSchedule, 2) = "1,") Then ' Every x Minutes
    xDelta = CInt(Right(xSchedule, Len(xSchedule) - 2))
    SQL2 = "DATEADD(n," & xDelta & ",GETDATE())"
  End If
  If (Left(xSchedule, 2) = "2,") Then ' Every x Days
    xDelta = CInt(Right(xSchedule, Len(xSchedule) - 2))
    SQL2 = "DATEADD(d," & xDelta & ",GETDATE())"
  End If
  
  SQL = "UPDATE tblDistroTasks SET NextDateTime = "
  If (SQL2 = "") Then
    SQL = SQL + "NULL "
  Else
    SQL = SQL + SQL2
  End If
  SQL = SQL + " WHERE TaskID = " & TaskID
  DBConn.Execute SQL
  
  Set CSCDB1 = Nothing
  
  ProcessSchedule = True
  Exit Function
ProcessScheduleError:
  ProcessSchedule = False
  Exit Function
End Function

Public Function FindATask(isTakeControl As Boolean) As Integer
  Dim CSCDB1 As Recordset
  Dim SQL As String
  
  On Error GoTo FindATaskError
  FindATask = 0
  Set CSCDB1 = New ADODB.Recordset
  
  CloseDatabase
  OpenDatabase False
  SQL = "SELECT TOP 1 * FROM tblDistroTasks WHERE"
  SQL = SQL + " Status = 0 AND (NextDateTime <= GETDATE())"
  If (ExclusiveMode) Then
    SQL = SQL + " AND (AgentAllow = " & AgentID & ")"
  Else
    SQL = SQL + " AND (AgentAllow & " & AgentID & " <> 0)"
  End If
  SQL = SQL + " AND (AgentRunning & " & AgentID & " = 0)"
  SQL = SQL + " AND (NumInstance < MaxInstance)"
  SQL = SQL + " ORDER BY NextDateTime DESC"
  CSCDB1.Open SQL, DBConn, adOpenForwardOnly
  
  If (CSCDB1.EOF) Then
    CSCDB1.Close
    CloseDatabase
    Exit Function
  End If
  If isTakeControl Then SetStatus CSCDB1("TaskID"), 1
  FindATask = CSCDB1("TaskID")
  CSCDB1.Close
  
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

