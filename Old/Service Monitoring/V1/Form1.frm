VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4410
   StartUpPosition =   1  'CenterOwner
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As Byte, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Type SERVICE_STATUS
  dwServiceType As Long
  dwCurrentState As Long
  dwControlsAccepted As Long
  dwWin32ExitCode As Long
  dwServiceSpecificExitCode As Long
  dwCheckPoint As Long
  dwWaitHint As Long
End Type

Private Type QUERY_SERVICE_CONFIG
  dwServiceType As Long
  dwStartType As Long
  dwErrorControl As Long
  lpBinaryPathName As Long 'String
  lpLoadOrderGroup As Long ' String
  dwTagId As Long
  lpDependencies As Long 'String
  lpServiceStartName As Long 'String
  lpDisplayName As Long  'String
End Type

Private Const SERVICE_STOPPED = &H1
Private Const SERVICE_START_PENDING = &H2
Private Const SERVICE_STOP_PENDING = &H3
Private Const SERVICE_RUNNING = &H4
Private Const SERVICE_CONTINUE_PENDING = &H5
Private Const SERVICE_PAUSE_PENDING = &H6
Private Const SERVICE_PAUSED = &H7
Private Const SERVICE_ACCEPT_STOP = &H1
Private Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Private Const SERVICE_ACCEPT_SHUTDOWN = &H4
Private Const SC_MANAGER_CONNECT = &H1
Private Const SERVICE_INTERROGATE = &H80
Private Const GENERIC_READ = &H80000000
Private Const ERROR_INSUFFICIENT_BUFFER = 122

Private Sub Form_Load()
  'On Error Resume Next
  Dim hSCM  As Long
  Dim hSVC As Long
  Dim pSTATUS As SERVICE_STATUS
  Dim udtConfig As QUERY_SERVICE_CONFIG
  Dim lRet As Long
  Dim lBytesNeeded As Long
  Dim sTemp As String
  Dim pFileName As Long
  Dim DSN As String
  Dim DB As ADODB.Recordset
  Dim Conn As ADODB.Connection
  Dim SQL As String, SQL2 As String, Hostname As String, Service As String
  Dim Ping As New XPing.XPing
  Dim Alert As New Monitoring.Alerter
  Dim Constants As New Monitoring.Constants, A As Boolean
  
  Set DB = CreateObject("adodb.recordset")
  Set Conn = CreateObject("adodb.connection")
  DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=CBDXAAI"
  SQL = "SELECT * FROM tblMonitor Where Type=2"
  
  Conn.Open DSN
  DB.Open SQL, Conn
  
  'List1.Clear
  
  Do Until DB.EOF
    ' Open The Service Control Manager
    Hostname = DB("Task")
    Service = DB("Var1")
    hSCM = OpenSCManager(Hostname, vbNullString, SC_MANAGER_CONNECT)
    If hSCM = 0 Then
      Ping.Hostname = Hostname
      Ping.Ping
      If Ping.LastError <> 0 Then
        Select Case Err.LastDllError
          Case 1722
            A = Alert.AddEvent(CStr(Date), CStr(Time), Constants.asCritical, "Service Monitoring", Constants.agNTSS, "Host not found - " & Hostname, "Host not found - " & Hostname, "Check " & Hostname)
            GoTo CloseHandles
          Case Else
            Debug.Print "Error - " & Err.LastDllError
        End Select
      End If
    End If
    
    ' Open the specific Service to obtain a handle
    '
    hSVC = OpenService(hSCM, Trim(Service), GENERIC_READ)
    If hSVC = 0 Then
      A = Alert.AddEvent(CStr(Date), CStr(Time), Constants.asCritical, "Service Monitoring", Constants.agNTSS, Service & " Service not found - " & Hostname, Service & " Service not found - " & Hostname, "Check " & Hostname)
      Debug.Print Hostname & "-" & Service & " does not exist"
      ' Ensure that the host is available
      Ping.Hostname = Hostname
      Ping.Ping
      If Ping.LastError = 0 Then
        ' If host is available, then the following is valid
        ' Add event explaining the next course of action
        SQL2 = "INSERT INTO tblEvents (Hostname, Action, Initiated_by,Description,Complete,StartDateTime,ActionType) "
        SQL2 = SQL2 & "VALUES ('" & Hostname & "','Service Monitoring','Automatic Monitoring','" & Service & " does not exist.  Monitoring of this service for " & Hostname & " has been removed.',0,GETDATE(),5)"
        Conn.Execute SQL2
        ' Remove this service/server from the list of those monitored
        SQL2 = "DELETE FROM tblMonitor WHERE Task = '" & Hostname & "' AND Type=2 AND Var1 = '" & Service & "'"
        Conn.Execute SQL2
      Else
        ' Server is not available
        A = Alert.AddEvent(CStr(Date), CStr(Time), Constants.asCritical, "Service Monitoring", Constants.agNTSS, "Host not found - " & Hostname, "Host not found - " & Hostname, "Check " & Hostname)
      End If
      GoTo CloseHandles
    End If
    
    ' Fill the Service Status Structure
    '
    lRet = QueryServiceStatus(hSVC, pSTATUS)
    If lRet = 0 Then
      'MsgBox "Error - " & Err.LastDllError
      GoTo CloseHandles
    End If
    
    ' Report the Current State
    '
    Select Case pSTATUS.dwCurrentState
    Case SERVICE_STOPPED
      sTemp = "Stopped"
    Case SERVICE_START_PENDING
      sTemp = "Being Started"
    Case SERVICE_STOP_PENDING
      sTemp = "Being Stopped"
    Case SERVICE_RUNNING
      sTemp = "Running"
    Case SERVICE_CONTINUE_PENDING
      sTemp = "Being Continued"
    Case SERVICE_PAUSE_PENDING
      sTemp = "Being Paused"
    Case SERVICE_PAUSED
      sTemp = "Paused"
    Case SERVICE_ACCEPT_STOP
      sTemp = "Stopped"
    Case SERVICE_ACCEPT_PAUSE_CONTINUE
      sTemp = "SERVICE_ACCEPT_PAUSE_CONTINUE"
    Case SERVICE_ACCEPT_SHUTDOWN
      sTemp = "Being Shutdown"
    Case Else
      sTemp = "UNKNOWN"
    End Select
    
    Debug.Print DB("Task") & "-" & DB("Var1") & " " & sTemp
    If sTemp <> "Running" Then
      SQL2 = "INSERT INTO tblEvents (Hostname, Action, Initiated_by,Description,Complete,StartDateTime,ActionType) "
      SQL2 = SQL2 & "VALUES ('" & Hostname & "','Service Monitoring','Automatic Monitoring','" & Service & " is " & sTemp & "',0,GETDATE(),5)"
      A = Alert.AddEvent(CStr(Date), CStr(Time), Constants.asCritical, "Service Monitoring", Constants.agNTSS, Service & " on " & Hostname & " is " & sTemp, Service & " on " & Hostname & " is " & sTemp, "Check " & Hostname)
      Conn.Execute SQL2
    End If
    'Debug.Print SQL2
    SQL2 = ""
        
CloseHandles:
     'Close the Handle to the Service
    CloseServiceHandle (hSVC)
    
    ' Close the Handle to the Service Control Manager
    CloseServiceHandle (hSCM)
    DB.MoveNext
  Loop
  Set DB = Nothing
  Set Conn = Nothing
  Unload Me
  End
End Sub
