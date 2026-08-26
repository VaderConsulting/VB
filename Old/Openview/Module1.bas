Attribute VB_Name = "Module1"

Sub Main()
  Dim SourcePath As String, DestPath As String, LastRunTime As Date
  Dim MRTGInfo As String, EventDate As Date
  Dim ADOConn As ADODB.Connection, ADORS As ADODB.Recordset
  Dim DSN As String, SQL As String
  DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=CBDXAAI"
  Set ADOConn = CreateObject("ADODB.Connection")
  Set ADORS = CreateObject("ADODB.Recordset")
  Dim Alert As New Monitoring.Alerter
  Dim Constants As New Monitoring.Constants
  Dim A As Boolean
  SourcePath = "\\10.1.1.28\ems\logs\event_thisweek.log"
  DestPath = "c:\temp\events.log"
  
  ' Remove existing file
  If Dir(DestPath) <> "" Then
    Kill DestPath
  End If
  
  On Error GoTo noSource
    ' Get most recent logfile
    FileCopy SourcePath, DestPath
  On Error Resume Next
    
  ' Get last run time
  SQL = "SELECT Var1 from tblMonitor WHERE Type=5"
  On Error GoTo NoSQL
    ADOConn.Open DSN
    ADORS.Open SQL, ADOConn
    LastRunTime = CDate(ADORS("Var1"))
    ' Remove 5 minutes from LastRunTime to ensure overlap takes care of time sync problems
    LastRunTime = DateAdd("d", -1, LastRunTime)
    ADORS.Close
  On Error GoTo 0
  Open DestPath For Input As #1
    Do Until EOF(1)
      Line Input #1, MRTGInfo
        ' Check if Node up or Node Down
        If InStr(1, MRTGInfo, "OV_Node_") <> 0 Then
          EventDate = Left(MRTGInfo, 17)
          ' Check if event has not already been processed
          'Debug.Print EventDate & " " & LastRunTime
          If EventDate >= LastRunTime Then
            P1 = InStr(19, MRTGInfo, ".")
            If P1 = 0 Then ' Not a DNS name, so extract NETBIOS Name
              P1 = InStr(19, MRTGInfo, " ")
            End If
            HostName = Mid(MRTGInfo, 19, P1 - 19)
            P2 = InStr(19, MRTGInfo, " ")
            P3 = InStr(P2 + 1, MRTGInfo, " ")
            MRTGType = Mid(MRTGInfo, P2 + 1, P3 - P2 - 1)
            AlertSeverity = Right(MRTGInfo, Len(MRTGInfo) - P3)
            Select Case MRTGType
              Case "OV_Node_Down"
                SQL = "UPDATE tblNodes SET State = 0, Updated = '" & Format(EventDate, "YYYYMMDD HH:MM:SS") & "' WHERE Hostname like '" & HostName & "'"
                A = Alert.AddEvent(CStr(Date), CStr(Time), Constants.asCritical, "HP Openview", Constants.agNTSS + Constants.agWAN, HostName & " is down", HostName & " is down", "Check " & HostName)
              Case "OV_Node_Up"
                SQL = "UPDATE tblNodes SET State = 1, Updated = '" & Format(EventDate, "YYYYMMDD HH:MM:SS") & "' WHERE Hostname like '" & HostName & "'"
                A = Alert.AddEvent(CStr(Date), CStr(Time), Constants.asInformation, "HP Openview", Constants.agNTSS + Constants.agWAN, HostName & " is up", HostName & " is up", "Information only")
            End Select
            Debug.Print SQL
            ADOConn.Execute SQL
          End If
        End If
    Loop
  Close 1
  SQL = "UPDATE tblMonitor SET Var1 = '" & Format(Date, "DD/MM/YY") & " " & Format(Time, "HH:MM:SS") & "' WHERE Type=5"
  'Debug.Print SQL
  ADOConn.Execute SQL
  ADOConn.Close
  Set ADORS = Nothing
  Set ADOConn = Nothing
  End
noSource:
  Open "c:\temp\HPOVAlert.txt" For Append As #1
    Print #1, "Source file error at " & Format(Date, "HH:MM:SS")
  Close 1
  Set ADORS = Nothing
  Set ADOConn = Nothing
  End
NoSQL:
  Open "c:\temp\HPOVAlert.txt" For Append As #1
    Print #1, "SQL Error at " & Format(Date, "HH:MM:SS")
  Close 1
  Set ADORS = Nothing
  Set ADOConn = Nothing
  End
End Sub
