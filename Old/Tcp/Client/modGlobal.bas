Attribute VB_Name = "modGlobal"
Option Explicit

'Connection State Constants
Public Const csNotAuthorised = 0
Public Const csAuthorised = 1
Public Const csConfigured = 2
Public Const csReady = 3

'TCP Timeouts (all in seconds)
Public Const ttConnect = 5
Public Const ttAuthorise = 5
Public Const ttConfiguration = 5
Public Const ttReady = 5

'Environment Variables
Public MasterHost As String
Public MasterPort As Long
Public AgentID As Integer
Public LogSeverity As Integer
Public DoTCPEncrypt As Boolean
Public boolFirstRun As Boolean

Public LastSentString As String

'Encryption Settings

Public Processes As Collection
Public adoConn As ADODB.Connection
Public adoRS As ADODB.Recordset
Public DSN As String

' Character Options
Public boolAgent As Boolean
Public boolSounds As Boolean
Public boolWordBalloon As Boolean
Public boolStopBetweenActions As Boolean
Public boolAutoHideBalloon As Boolean
Public boolAutoPace As Boolean
Public boolSizeBalloontoText As Boolean
Public boolVerbose As Boolean
Public boolQuiet As Boolean
Public boolOutputSummaryOnly As Boolean
Public strSelectedAgent As String

' Sound Options
Public strInformationSound As String
Public strWarningSound As String
Public strCriticalSound As String

'voice object
Public oVoice As VTxtAuto.VTxtAuto

' Filter Options
Public boolIgnoreInformation As Boolean
Public boolIgnoreWarning As Boolean
Public boolIgnoreNTSS As Boolean
Public boolIgnoreHelpdesk As Boolean
Public boolIgnoreEUC As Boolean
Public boolIgnoreWAN As Boolean

Public Sub CloseObjects()
  frmMain.tcpMain.Close
End Sub

Public Function InitObjects() As Boolean
  Dim Host As String
  Dim Misc As String
  Dim isOK As Boolean
  
  InitObjects = False
  WriteToLog 1, "Starting Agent..."
  
  DoTCPEncrypt = True
  Host = ""
  AgentID = -1
  boolFirstRun = False
  
  Set Processes = New Collection
  
  App.StartLogging "", vbLogToNT
  Randomize Timer

  ' R1FE: = Broadcast from client (usually indicates a problem)
  
  '*****  MasterHost
  MasterHost = GetValue(Host, "Software\Alerter", "MasterHost", isOK, REG_SZ)
  If Not isOK Or (MasterHost = "") Then
    App.LogEvent "Invalid Master Hostname", vbLogEventTypeError
    WriteToLog 1, "Invalid Master Hostname"
    Exit Function
  End If
  
  '*****  MasterPort
  Misc = GetValue(Host, "Software\Alerter", "MasterPort", isOK, REG_SZ)
  If Not isOK Or (Misc = "") Or Not IsNumeric(Misc) Then
    App.LogEvent "Invalid Master Port", vbLogEventTypeError
    WriteToLog 1, "Invalid Master Port"
    Exit Function
  End If
  MasterPort = CLng(Misc)
  
  '***** Get setting for First run option
  Misc = GetValue(Host, "Software\Alerter", "First Run", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid First Run Value", vbLogEventTypeError
    WriteToLog 1, "Invalid First Run Value"
  End If
  If Trim(Misc) = "" Or Misc = "1" Then
    boolFirstRun = True  ' Set up Personal Alerter first run
    Misc = "1"
    SetValueString Host, "Software\Alerter", "First Run", Misc, REG_SZ
  End If
  
  '***** Get setting for Agent option
  Misc = GetValue(Host, "Software\Alerter", "Agent", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Agent Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Agent Value"
  End If
  If Trim(Misc) = "" Then
    boolFirstRun = True  ' Set up Personal Alerter first run
    Misc = "1"
    SetValueString Host, "Software\Alerter", "Agent", Misc, REG_SZ
  End If
  boolAgent = CBool(Misc)
  If Not boolAgent Then frmMain.mnuLoad.Enabled = False
  
  '***** Get setting for Sounds option
  Misc = GetValue(Host, "Software\Alerter", "Sounds", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Sounds Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Sounds Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "1"
    SetValueString Host, "Software\Alerter", "Sounds", Misc, REG_SZ
  End If
  boolSounds = CBool(Misc)
  
  '***** Get setting for Word Balloon option
  Misc = GetValue(Host, "Software\Alerter", "Word Balloon", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Word Balloon Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Word Balloon Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "1"
    SetValueString Host, "Software\Alerter", "Word Balloon", Misc, REG_SZ
  End If
  boolWordBalloon = CBool(Misc)
  
  '***** Get setting for Stop Between Actions option
  Misc = GetValue(Host, "Software\Alerter", "Stop Between Actions", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Stop Between Actions Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Stop Between Actions Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "1"
    SetValueString Host, "Software\Alerter", "Stop Between Actions", Misc, REG_SZ
  End If
  boolStopBetweenActions = CBool(Misc)
  
  '***** Get setting for Auto Hide Balloon option
  Misc = GetValue(Host, "Software\Alerter", "Auto Hide Balloon", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Auto Hide Balloon Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Auto Hide Balloon Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "1"
    SetValueString Host, "Software\Alerter", "Auto Hide Balloon", Misc, REG_SZ
  End If
  boolAutoHideBalloon = CBool(Misc)
  
  '***** Get setting for Auto Pace option
  Misc = GetValue(Host, "Software\Alerter", "Auto Pace", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Auto Pace Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Auto Pace Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "1"
    SetValueString Host, "Software\Alerter", "Auto Pace", Misc, REG_SZ
  End If
  boolAutoPace = CBool(Misc)
  
  '***** Get setting for Size Balloon to Text option
  Misc = GetValue(Host, "Software\Alerter", "Size Balloon to Text", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Size Balloon to Text Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Size Balloon to Text Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "1"
    SetValueString Host, "Software\Alerter", "Size Balloon to Text", Misc, REG_SZ
  End If
  boolSizeBalloontoText = CBool(Misc)
  
  '***** Get setting for Verbose option
  Misc = GetValue(Host, "Software\Alerter", "Verbose", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Verbose Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Verbose Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "1"
    SetValueString Host, "Software\Alerter", "Verbose", Misc, REG_SZ
  End If
  boolVerbose = CBool(Misc)
  
  '***** Get setting for Quiet option
  Misc = GetValue(Host, "Software\Alerter", "Quiet", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Quiet Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Quiet Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "1"
    SetValueString Host, "Software\Alerter", "Quiet", Misc, REG_SZ
  End If
  boolQuiet = CBool(Misc)
  
  '***** Get setting for Output Summary Only option
  Misc = GetValue(Host, "Software\Alerter", "Output Summary Only", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Output Summary Only Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Output Summary Only Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "0"
    SetValueString Host, "Software\Alerter", "Output Summary Only", Misc, REG_SZ
  End If
  boolOutputSummaryOnly = CBool(Misc)
  
  '***** Get setting for Agent
  Misc = GetValue(Host, "Software\Alerter", "Character", isOK, REG_SZ)
  If Not isOK Then
    App.LogEvent "Invalid Agent Option Value", vbLogEventTypeError
    WriteToLog 1, "Invalid Agent Option Value"
  End If
  If Trim(Misc) = "" Then
    Misc = "genie.acs"
    SetValueString Host, "Software\Alerter", "Character", Misc, REG_SZ
  End If
  strSelectedAgent = Misc
  
  '***** Get setting for Information Sound option
  Misc = GetValue(Host, "Software\Alerter", "Information Sound", isOK, REG_SZ)
  If Trim(Misc) = "" Then
    Misc = "more work.wav"
    SetValueString Host, "Software\Alerter", "Information Sound", Misc, REG_SZ
  End If
  strInformationSound = Misc
  
  '***** Get setting for Warning Sound option
  Misc = GetValue(Host, "Software\Alerter", "Warning Sound", isOK, REG_SZ)
  If Trim(Misc) = "" Then
    Misc = "im not listening.wav"
    SetValueString Host, "Software\Alerter", "Warning Sound", Misc, REG_SZ
  End If
  strWarningSound = Misc
  
  '***** Get setting for Critical Alert Sound option
  Misc = GetValue(Host, "Software\Alerter", "Critical Alert Sound", isOK, REG_SZ)
  If Trim(Misc) = "" Then
    Misc = "3.wav"
    SetValueString Host, "Software\Alerter", "Critical Alert Sound", Misc, REG_SZ
  End If
  strCriticalSound = Misc
  
  '***** Get settings for Filters
  Misc = GetValue(Host, "Software\Alerter", "Ignore Information", isOK, REG_SZ)
  If Trim(Misc) = "" Then
    Misc = "0"
    SetValueString Host, "Software\Alerter", "Ignore Information", Misc, REG_SZ
  End If
  boolIgnoreInformation = CBool(Misc)
  
  Misc = GetValue(Host, "Software\Alerter", "Ignore Warning", isOK, REG_SZ)
  If Trim(Misc) = "" Then
    Misc = "0"
    SetValueString Host, "Software\Alerter", "Ignore Warning", Misc, REG_SZ
  End If
  boolIgnoreWarning = CBool(Misc)
  
  Misc = GetValue(Host, "Software\Alerter", "Ignore NTSS", isOK, REG_SZ)
  If Trim(Misc) = "" Then
    Misc = "0"
    SetValueString Host, "Software\Alerter", "Ignore NTSS", Misc, REG_SZ
  End If
  boolIgnoreNTSS = CBool(Misc)
  
  Misc = GetValue(Host, "Software\Alerter", "Ignore Helpdesk", isOK, REG_SZ)
  If Trim(Misc) = "" Then
    Misc = "0"
    SetValueString Host, "Software\Alerter", "Ignore Helpdesk", Misc, REG_SZ
  End If
  boolIgnoreInformation = CBool(Misc)
  
  Misc = GetValue(Host, "Software\Alerter", "Ignore EUC", isOK, REG_SZ)
  If Trim(Misc) = "" Then
    Misc = "0"
    SetValueString Host, "Software\Alerter", "Ignore EUC", Misc, REG_SZ
  End If
  boolIgnoreEUC = CBool(Misc)
  
  Misc = GetValue(Host, "Software\Alerter", "Ignore WAN", isOK, REG_SZ)
  If Trim(Misc) = "" Then
    Misc = "0"
    SetValueString Host, "Software\Alerter", "Ignore WAN", Misc, REG_SZ
  End If
  boolIgnoreWAN = CBool(Misc)
    
  '***** Establish TCP Connection
  If StartTCPClient Then
    WriteToLog 1, "TCP Connection Established"
  Else
    WriteToLog 0, "TCP Connection FAILED"
    Exit Function
  End If
  
  '***** Authorise the Connection
  If Not AuthoriseConnection Then
    WriteToLog 0, "Failed to authorise to " & MasterHost
    SendString "R1FD:Client failed to authorise"
    Exit Function
  End If
  WriteToLog 4, "Authorised to " & MasterHost
  
  '***** Configure the Connection
  If Not ConfigureConnection Then
    WriteToLog 0, "Failed to Configure"
    SendString "R1FD:Client failed to configure"
    Exit Function
  End If
  'WriteToLog 4, "Successfully Configured.  AgentID set to " & AgentID & " by " & MasterHost
  
  '***** Send username to master
  SendString "R103:" & UCase(Environ$("username"))
  
  DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PAlerter.mdb;Persist Security Info=False"
  InitObjects = True
  
  RemoveOldAlerts
  
End Function

Public Sub RemoveOldAlerts()
  Dim SQL As String
  SQL = "DELETE FROM tblAlerts WHERE Complete = 1"
  Set adoConn = New ADODB.Connection
  adoConn.Open DSN
  adoConn.Execute SQL
  adoConn.Close
End Sub

Public Sub Speak(Text As String, Optional Filename As String, Optional Group As Integer = 15, Optional Severity As Integer = 3)
  Dim intTemp As Integer
  If boolQuiet = True Then
    If Time > CDate("7:00 PM") Then
      Exit Sub
    End If
    If Time < CDate("6:00 AM") Then
      Exit Sub
    End If
  End If
      
  ' Check Severity
  If boolIgnoreInformation And Severity = 1 Then Exit Sub
  If boolIgnoreWarning And Severity = 2 Then Exit Sub
  
  ' Check Group Membership
  intTemp = 0 ' intTemp represents what groups are ignored
  If boolIgnoreNTSS Then intTemp = intTemp Or 1
  If boolIgnoreHelpdesk Then intTemp = intTemp Or 2
  If boolIgnoreEUC Then intTemp = intTemp Or 4
  If boolIgnoreWAN Then intTemp = intTemp Or 8
  intTemp = Not intTemp
  If (Group And intTemp) = 0 Then Exit Sub

  If Not boolAgent Then
    Do Until oVoice.IsSpeaking = False
      DoEvents
    Loop
    If Text <> "" Then
      oVoice.Speak Text, vtxtst_READING
    End If
    Exit Sub
  End If
  
  If frmMain.CharLoaded Then
    If Text <> "" Then
      If Filename = "" Then
        frmMain.Character.Speak Text
      Else
        frmMain.Character.Speak Text, Filename
      End If
    Else
      frmMain.Character.Speak "", Filename
    End If
  End If
End Sub

Private Function ConfigureConnection() As Boolean
  Dim OldTime As Date
  
  ConfigureConnection = False
  If (frmMain.tcpMain.State <> sckConnected) Then
    WriteToLog 4, "(Configuration) Invalid Socket State - " & frmMain.tcpMain.State
    Exit Function
  End If
  
  SendString "R101:" & UCase(Environ("COMPUTERNAME")) & vbCrLf
  OldTime = Now
  
  Do While (ConnState <> csConfigured) And (DateDiff("s", OldTime, Now) < ttConfiguration)
    DoEvents
  Loop
  
  ConfigureConnection = (ConnState = csConfigured)
End Function

Private Function AuthoriseConnection() As Boolean
  Dim OldTime As Date
  
  AuthoriseConnection = False
  If (frmMain.tcpMain.State <> sckConnected) Then
    WriteToLog 4, "(Authorise) Invalid Socket State - " & frmMain.tcpMain.State
    Exit Function
  End If
    
  SendString "A001:ALERTER" + vbCrLf
  OldTime = Now
  
  Do While (ConnState <> csAuthorised) And (DateDiff("s", OldTime, Now) < ttAuthorise)
    DoEvents
  Loop
  
  AuthoriseConnection = (ConnState = csAuthorised)
End Function

Private Function StartTCPClient() As Boolean
  Dim OldTime As Date
  On Error GoTo StartTCPClientError
  StartTCPClient = False
  
  OldTime = Now
  frmMain.tcpMain.Connect MasterHost, MasterPort
  
  ' This next line ensures that the master has to respond within the allocated time
  ' Remove remarks in production!!!
  Do While (frmMain.tcpMain.State <> sckConnected) And (DateDiff("s", OldTime, Now) < ttConnect)
    DoEvents
  Loop
  
  StartTCPClient = (frmMain.tcpMain.State = sckConnected)
  Exit Function

StartTCPClientError:
  Exit Function
End Function

'Public Function FindCompletedProcesses() As Boolean
'  Dim oProc As clsProcess
'  Dim i As Integer
'  Dim DoAgain As Boolean
'
'  FindCompletedProcesses = False
'
'  Do
'    DoAgain = False
'    For i = 1 To Processes.Count
'      Set oProc = Processes.Item(i)
'      If ProcessCompleted(oProc) Then
'        WriteToLog 3, "Process Stopped.  Task ID (" & oProc.TaskID & ") PID (" & oProc.dwProcessID & ")"
'        'LogEntry oProc.TaskID, 2, "Task Stopped"
'        'LogEntry oProc.TaskID, 3, GetProcessExitCode(oProc)
'
'        'OpenDatabase True
'        'If Not AssertAgentNotRunning(oProc.TaskID) Then WriteToLog 3, "Can't AssertAgentNotRunning.  Task ID (" & oProc.TaskID & ")"
'        'If Not AssertTaskComplete(oProc.TaskID) Then WriteToLog 3, "Can't AssertTaskComplete.  Task ID (" & oProc.TaskID & ")"
'        'CloseDatabase
'
'        CleanupProcess oProc
'        Processes.Remove i
'        Set oProc = Nothing
'        DoAgain = True
'        Exit For
'      End If
'    Next
'  Loop Until Not DoAgain
'
'  FindCompletedProcesses = True
'End Function

Public Function GetValue(ByVal Hostname As String, ByVal Key As String, ByVal Value As String, ByRef ECode As Boolean, regType As Long) As String
    Dim OpenKeyVal As Long
    Dim OpenHiveVal As Long
    Dim RResult As Long
    Dim InfoTextStr As String
    
    'Init
    ECode = False
    Hostname = Trim(Hostname)
    Key = Trim(Key)
    Value = Trim(Value)
    GetValue = ""
        
    RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then Exit Function
    
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
    'InfoTextStr = RegistryQueryValue(OpenKeyVal, Value, REG_SZ)
    InfoTextStr = RegistryQueryValue(OpenKeyVal, Value, regType)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    GetValue = InfoTextStr
    ECode = True
End Function

Public Function SetValueString(Hostname As String, Key As String, Value As String, Data As String, regType As Long) As Long
    Dim OpenKeyVal As Long
    Dim RResult As Long
    Dim OpenHiveVal As Long
    Dim strValue, CMPTRName, keytogo As String, InfoTextStr As String
    Dim x As Integer
    
    'GetIPCConnection (Trim(Hostname))
    
    CMPTRName = Trim(Hostname)
    
    RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
    keytogo = Trim(Key)
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
    RegistryWriteValue Data, OpenKeyVal, Value, regType
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    'DisIPCConnection (Trim(Hostname))
        
End Function

Public Sub SetDistroTasksState()
  If Not (ConnState = csReady) Then
    If (Processes.Count = 0) Then
      'frmMain.txtState.Text = "Idle"
      WriteToLog 1, "Agent has stopped"
    Else
      'frmMain.txtState.Text = "Stopping"
      WriteToLog 2, "Agent still stopping (" & Processes.Count & " Task/s Outstanding)"
    End If
  Else
    'frmMain.txtState.Text = "Running"
  End If
End Sub

Public Function StopDistroTasks() As Boolean
  WriteToLog 2, "Agent stopping..."
  StopDistroTasks = False
  
  If ConnState <> csReady Then
    SendString "R105:NAK"
    Exit Function
  End If
  ConnState = csConfigured
  SetDistroTasksState
  
  SendString "R105:OK" & vbCrLf
  StopDistroTasks = True
End Function

Public Function StartDistroTasks() As Boolean
  Dim OldTime As Date
  
  WriteToLog 2, "Agent going active..."
  
  StartDistroTasks = False
  If (frmMain.tcpMain.State <> sckConnected) Then
    WriteToLog 4, "(StartAgent) Invalid Socket State - " & frmMain.tcpMain.State
    Exit Function
  End If
    
  SendString "R102:" + vbCrLf
  OldTime = Now
  
  Do While (ConnState <> csReady) And (DateDiff("s", OldTime, Now) < ttReady)
    DoEvents
  Loop
  
  StartDistroTasks = (ConnState = csReady)
  
  SetDistroTasksState

  WriteToLog 2, "Agent is active"

End Function

Public Function DoubleDigit(ByVal i As Integer) As String
  Dim s As String
  s = CStr(i)
  Do While Len(s) < 2
    s = "0" + s
  Loop
  DoubleDigit = s
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

Public Sub WriteToLog(ByVal Sev As Integer, ByVal xStr As String)
  On Error GoTo WriteToLogError
  Dim s As String
  If (Sev > LogSeverity) Then Exit Sub
  s = Year(Now) & DoubleDigit(Month(Now)) & DoubleDigit(Day(Now)) & " " & DoubleDigit(Hour(Now)) & ":" & DoubleDigit(Minute(Now)) & ":" & DoubleDigit(Second(Now))
  s = s + " (" & Sev & ") " + xStr
  Open ".\" & App.EXEName & ".txt" For Append As #1
  Print #1, s
  Close #1
  Exit Sub
WriteToLogError:
  Close #1
  Err.Clear
  Exit Sub
End Sub

Public Sub AddToHistory(ByVal Text As String)
  frmMain.lstHistory.AddItem Text, 0
  frmMain.lstHistory.Refresh
End Sub

Public Function Encrypt(ByVal Text As String, ByVal EncKey As String) As String
  Encrypt = Text
End Function

Public Function Decrypt(ByVal Text As String, ByVal EncKey As String) As String
  Decrypt = Text
End Function

Sub PlayWav(WaveFile As String)
    Dim n As Long
    If UCase(Right(WaveFile, 3)) <> "WAV" Then Exit Sub
    'Load and Play the sound.
    n = sndPlaySound(WaveFile, 0)
End Sub
