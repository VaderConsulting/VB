Attribute VB_Name = "modGlobal"
Option Explicit

'Connection State Constants
Public Const csNotAuthorised = 0
Public Const csAuthorised = 1
Public Const csConfigured = 2
Public Const csReady = 3

'Process State Constants
Public Const psIdle = 0
Public Const psPending = 1
Public Const psRunning = 2
Public Const psError = -1

'Default Agent Settings
Public Const DefaultMaxProcesses = 5

'Encryption Settings
Public DoTCPEncrypt As Boolean

Public SQLServer As String
Public SQLDatabase As String
Public SQLDSN As String
Public LogSeverity As Integer

Public Processes As Collection
Public TCPConns As Collection
Public Const MaxTCPConns = 15

Public DebugMode As Boolean

Public Function BroadcastTCPMessage(ByVal s As String) As Boolean
  Dim oTCP As clsTCP
  BroadcastTCPMessage = False
  For Each oTCP In TCPConns
    oTCP.SendBroadcast s
    DoEvents
    BroadcastTCPMessage = True
    Debug.Print "Sending broadcast to " & oTCP.ID
  Next
End Function

'****************************************************************
'* Change this function to return back an object reference or Item Number
'****************************************************************
Function FindTCPConnIDfromAgentName(AgentName As String) As Integer
  Dim i As Integer
  Dim oTCP As clsTCP
  FindTCPConnIDfromAgentName = 0
  For i = 1 To TCPConns.Count
    Set oTCP = TCPConns.Item(i)
    If oTCP.isConfigured Then
      If oTCP.AgentHostname = AgentName Then FindTCPConnIDfromAgentName = i
    End If
  Next
  Set oTCP = Nothing
End Function

Public Function FindTCPConnFromID(ByVal Index As Integer) As Integer
  Dim oTCP As clsTCP
  Dim i As Integer
  
  FindTCPConnFromID = -1
  
  For i = 1 To TCPConns.Count
    Set oTCP = TCPConns.Item(i)
    If (oTCP.ID = Index) Then FindTCPConnFromID = i
  Next
End Function

Public Sub CloseObjects()
  BroadcastTCPMessage "S1FD:SHUTDOWN"
  frmMain.tcpMain(0).Close
  CloseDatabase
End Sub

Public Function InitObjects() As Boolean
  Dim Host As String
  Dim isOK As Boolean
  
  InitObjects = False
  WriteToLog 1, "Starting " & App.EXEName & "..."
  
  Host = ""
  SQLDSN = ""
  DoTCPEncrypt = True
  
  Set DBConn = New ADODB.Connection
  Set Processes = New Collection
  Set TCPConns = New Collection
  
  App.StartLogging "", vbLogToNT
  
  '*****  SQL Server
  SQLServer = GetValue(Host, "Software\Alerter", "SQL Server", isOK)
  If Not isOK Or (SQLServer = "") Then
    App.LogEvent "Invalid SQL Server", vbLogEventTypeError
    WriteToLog 1, "Invalid SQL Server"
    Exit Function
  End If
  
  '*****  SQL Database
  SQLDatabase = GetValue(Host, "Software\Alerter", "SQL Database", isOK)
  If Not isOK Or (SQLDatabase = "") Then
    App.LogEvent "Invalid SQL Database", vbLogEventTypeError
    WriteToLog 1, "Invalid SQL Database"
    Exit Function
  End If
  SQLDSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & SQLDatabase & ";Data Source=" & SQLServer
  
  If Not OpenDatabase Then
    App.LogEvent "Unable to open Database"
    WriteToLog 1, "Unable to open Database"
    Exit Function
  End If
     
  If StartTCPServer Then
    WriteToLog 1, "TCP Server Online"
  Else
    WriteToLog 1, "TCP Server FAILED"
  End If
  
  InitObjects = True
  WriteToLog 1, App.EXEName & " Started"
End Function

Private Function StartTCPServer() As Boolean
  On Error GoTo StartTCPServerError
  StartTCPServer = False
  
  frmMain.tcpMain(0).LocalPort = 9998
  frmMain.tcpMain(0).Listen
  
  StartTCPServer = True
  Exit Function

StartTCPServerError:
  Exit Function
End Function

Private Function GetValue(ByVal Hostname As String, ByVal Key As String, ByVal Value As String, ByRef ECode As Boolean) As String
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
    InfoTextStr = RegistryQueryValue(OpenKeyVal, Value, REG_SZ)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    GetValue = InfoTextStr
    ECode = True
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
  
  If DebugMode = True Then
    Open "c:\temp\" & App.EXEName & ".txt" For Append As #1
  Else
    Open ".\" & App.EXEName & ".txt" For Append As #1
  End If
  
  Print #1, s
  Debug.Print s
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

Function Encrypt(ByVal Text As String, ByVal EncKey As String) As String
  Encrypt = Text
End Function

Function Decrypt(ByVal Text As String, ByVal EncKey As String) As String
  Decrypt = Text
End Function
