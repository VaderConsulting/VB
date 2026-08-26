Attribute VB_Name = "gService"
Option Explicit

'Public Const MachineName = "s1"
'Public Const ServiceName = "Schedule"

Public Const SERVICES_ACTIVE_DATABASE = "ServicesActive"

' Controls
Public Const SERVICE_CONTROL_STOP = &H1
Public Const SERVICE_CONTROL_PAUSE = &H2

' Service State -- for CurrentState
Public Const SERVICE_STOPPED = &H1
Public Const SERVICE_START_PENDING = &H2
Public Const SERVICE_STOP_PENDING = &H3
Public Const SERVICE_RUNNING = &H4
Public Const SERVICE_CONTINUE_PENDING = &H5
Public Const SERVICE_PAUSE_PENDING = &H6
Public Const SERVICE_PAUSED = &H7

' Service Control Manager object specific access types
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SC_MANAGER_CONNECT = &H1
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const SC_MANAGER_LOCK = &H8
Public Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20

Public Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
'Service object specific access types
Public Const SERVICE_QUERY_CONFIG = &H1
Public Const SERVICE_CHANGE_CONFIG = &H2
Public Const SERVICE_QUERY_STATUS = &H4
Public Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Public Const SERVICE_START = &H10
Public Const SERVICE_STOP = &H20
Public Const SERVICE_PAUSE_CONTINUE = &H40
Public Const SERVICE_INTERROGATE = &H80
Public Const SERVICE_USER_DEFINED_CONTROL = &H100

Public Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)

Type SERVICE_STATUS
        dwServiceType As Long
        dwCurrentState As Long
        dwControlsAccepted As Long
        dwWin32ExitCode As Long
        dwServiceSpecificExitCode As Long
        dwCheckPoint As Long
        dwWaitHint As Long
End Type

Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long

Public Function StatusofService(Computer As String, service As String) As String
    Dim ServiceStatus As SERVICE_STATUS
    Dim lhSCManager As Long
    Dim lhService As Long
    Dim lhServiceStatus As Long

    lhSCManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
        
    If lhSCManager <> 0 Then
           
        lhService = OpenService(lhSCManager, service, SERVICE_ALL_ACCESS)
        
        If lhService <> 0 Then
               
            lhServiceStatus = QueryServiceStatus(lhService, ServiceStatus)
                
            If lhServiceStatus <> 0 Then
                Select Case ServiceStatus.dwCurrentState
                    Case SERVICE_STOPPED
                        StatusofService = service & " Service on " & Computer & " is stopped"
                    Case SERVICE_START_PENDING
                        StatusofService = service & " Service on " & Computer & " is Start Pending"
                    Case SERVICE_STOP_PENDING
                        StatusofService = service & " Service on " & Computer & " is Stop Pending"
                    Case SERVICE_RUNNING
                        StatusofService = service & " Service on " & Computer & " is Running"
                    Case SERVICE_CONTINUE_PENDING
                        StatusofService = service & " Service on " & Computer & "is Continue Pending"
                    Case SERVICE_PAUSE_PENDING
                        StatusofService = service & " Service on " & Computer & "is Pause Pending"
                    Case SERVICE_PAUSED
                        StatusofService = service & " Service on " & Computer & "is Paused"
                End Select
            End If
            CloseServiceHandle lhService
        End If
        CloseServiceHandle lhSCManager
    End If

End Function


Public Sub vbPauseService(Computer As String, service As String)
    Dim ServiceStatus As SERVICE_STATUS
    Dim lhSCManager As Long
    Dim lhService As Long
    Dim lresult As Long
    
    lhSCManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If lhSCManager <> 0 Then
        lhService = OpenService(lhSCManager, service, SERVICE_ALL_ACCESS)
        
        If lhService <> 0 Then
            lresult = ControlService(lhService, SERVICE_CONTROL_PAUSE, ServiceStatus)
            CloseServiceHandle lhService
        End If
        CloseServiceHandle lhSCManager
    End If
        
End Sub


Public Sub vbStartService(Computer As String, service As String)
    Dim ServiceStatus As SERVICE_STATUS
    Dim lhSCManager As Long
    Dim lhService As Long
    Dim lresult As Long
    
    lhSCManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    
    If lhSCManager <> 0 Then
        lhService = OpenService(lhSCManager, service, SERVICE_ALL_ACCESS)
        
        If lhService <> 0 Then
            lresult = StartService(lhService, 0, 0)
            CloseServiceHandle lhService
        End If
        CloseServiceHandle lhSCManager
    End If
        
End Sub


Public Sub vbStopService(Computer As String, service As String)
    Dim ServiceStatus As SERVICE_STATUS
    Dim lhSCManager As Long
    Dim lhService As Long
    Dim lresult As Long
    
    lhSCManager = OpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    
    If lhSCManager <> 0 Then
        lhService = OpenService(lhSCManager, service, SERVICE_ALL_ACCESS)
        
        If lhService <> 0 Then
            lresult = ControlService(lhService, SERVICE_CONTROL_STOP, ServiceStatus)
            CloseServiceHandle lhService
        End If
        CloseServiceHandle lhSCManager
    End If
End Sub

