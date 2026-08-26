VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SERVICES_ACTIVE_DATABASE = "ServicesActive"

' Controls
Const SERVICE_CONTROL_STOP = &H1
Const SERVICE_CONTROL_PAUSE = &H2

' Service State -- for CurrentState
Const SERVICE_STOPPED = &H1
Const SERVICE_START_PENDING = &H2
Const SERVICE_STOP_PENDING = &H3
Const SERVICE_RUNNING = &H4
Const SERVICE_CONTINUE_PENDING = &H5
Const SERVICE_PAUSE_PENDING = &H6
 Const SERVICE_PAUSED = &H7

' Service Control Manager object specific access types
 Const STANDARD_RIGHTS_REQUIRED = &HF0000
 Const SC_MANAGER_CONNECT = &H1
 Const SC_MANAGER_CREATE_SERVICE = &H2
 Const SC_MANAGER_ENUMERATE_SERVICE = &H4
 Const SC_MANAGER_LOCK = &H8
 Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
 Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20

 Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
'Service object specific access types
 Const SERVICE_QUERY_CONFIG = &H1
 Const SERVICE_CHANGE_CONFIG = &H2
 Const SERVICE_QUERY_STATUS = &H4
 Const SERVICE_ENUMERATE_DEPENDENTS = &H8
 Const SERVICE_START = &H10
 Const SERVICE_STOP = &H20
 Const SERVICE_PAUSE_CONTINUE = &H40
 Const SERVICE_INTERROGATE = &H80
 Const SERVICE_USER_DEFINED_CONTROL = &H100

 Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)

Type SERVICE_STATUS
        dwServiceType As Long
        dwCurrentState As Long
        dwControlsAccepted As Long
        dwWin32ExitCode As Long
        dwServiceSpecificExitCode As Long
        dwCheckPoint As Long
        dwWaitHint As Long
End Type

Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long

Public Function StatusofService(Computer As String, Service As String) As String
    Dim ServiceStatus As SERVICE_STATUS
    Dim lhSCManager As Long
    Dim lhService As Long
    Dim lhServiceStatus As Long

    lhSCManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
        
    If lhSCManager <> 0 Then
           
        lhService = OpenService(lhSCManager, Service, SERVICE_ALL_ACCESS)
        
        If lhService <> 0 Then
               
            lhServiceStatus = QueryServiceStatus(lhService, ServiceStatus)
                
            If lhServiceStatus <> 0 Then
                Select Case ServiceStatus.dwCurrentState
                    Case SERVICE_STOPPED
                        StatusofService = Service & " Service on " & Computer & " is stopped"
                    Case SERVICE_START_PENDING
                        StatusofService = Service & " Service on " & Computer & " is Start Pending"
                    Case SERVICE_STOP_PENDING
                        StatusofService = Service & " Service on " & Computer & " is Stop Pending"
                    Case SERVICE_RUNNING
                        StatusofService = Service & " Service on " & Computer & " is Running"
                    Case SERVICE_CONTINUE_PENDING
                        StatusofService = Service & " Service on " & Computer & "is Continue Pending"
                    Case SERVICE_PAUSE_PENDING
                        StatusofService = Service & " Service on " & Computer & "is Pause Pending"
                    Case SERVICE_PAUSED
                        StatusofService = Service & " Service on " & Computer & "is Paused"
                End Select
            End If
            CloseServiceHandle lhService
        End If
        CloseServiceHandle lhSCManager
    End If

End Function


Public Sub vbPauseService(Computer As String, Service As String)
    Dim ServiceStatus As SERVICE_STATUS
    Dim lhSCManager As Long
    Dim lhService As Long
    Dim lresult As Long
    
    lhSCManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If lhSCManager <> 0 Then
        lhService = OpenService(lhSCManager, Service, SERVICE_ALL_ACCESS)
        
        If lhService <> 0 Then
            lresult = ControlService(lhService, SERVICE_CONTROL_PAUSE, ServiceStatus)
            CloseServiceHandle lhService
        End If
        CloseServiceHandle lhSCManager
    End If
        
End Sub


Public Sub vbStartService(Computer As String, Service As String)
    Dim ServiceStatus As SERVICE_STATUS
    Dim lhSCManager As Long
    Dim lhService As Long
    Dim lresult As Long
    
    lhSCManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    
    If lhSCManager <> 0 Then
        lhService = OpenService(lhSCManager, Service, SERVICE_ALL_ACCESS)
        
        If lhService <> 0 Then
            lresult = StartService(lhService, 0, 0)
            CloseServiceHandle lhService
        End If
        CloseServiceHandle lhSCManager
    End If
        
End Sub


Public Sub vbStopService(Computer As String, Service As String)
    Dim ServiceStatus As SERVICE_STATUS
    Dim lhSCManager As Long
    Dim lhService As Long
    Dim lresult As Long
    
    lhSCManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    
    If lhSCManager <> 0 Then
        lhService = OpenService(lhSCManager, Service, SERVICE_ALL_ACCESS)
        
        If lhService <> 0 Then
            lresult = ControlService(lhService, SERVICE_CONTROL_STOP, ServiceStatus)
            CloseServiceHandle lhService
        End If
        CloseServiceHandle lhSCManager
    End If
End Sub


