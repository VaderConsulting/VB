Attribute VB_Name = "MServices"
Option Explicit

Public bServicesInit    As Boolean
Public vStatusStr       As Variant
Public vStartTypeStr    As Variant

Type SERVICE_STATUS
    dwServiceType       As Long
    dwCurrentState      As Long
    dwControlsAccepted  As Long
    dwWin32ExitCode     As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint        As Long
    dwWaitHint          As Long
End Type

'**********************************
'**  Constant Definitions:
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

'//
'// Service State -- for Enum Requests (Bit Mask)
'//
Public Const SERVICE_ACTIVE = 1
Public Const SERVICE_INACTIVE = 2
Public Const SERVICE_STATE_ALL = (SERVICE_ACTIVE Or SERVICE_INACTIVE)

'//
'// Service State -- for CurrentState
'//
Public Const SERVICE_STOPPED = 1
Public Const SERVICE_START_PENDING = 2
Public Const SERVICE_STOP_PENDING = 3
Public Const SERVICE_RUNNING = 4
Public Const SERVICE_CONTINUE_PENDING = 5
Public Const SERVICE_PAUSE_PENDING = 6
Public Const SERVICE_PAUSED = 7

Public Const SC_MANAGER_CONNECT             As Long = &H1
Public Const SC_MANAGER_CREATE_SERVICE      As Long = &H2
Public Const SC_MANAGER_ENUMERATE_SERVICE   As Long = &H4
Public Const SC_MANAGER_LOCK                As Long = &H8
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG  As Long = &H20
Public Const SC_MANAGER_QUERY_LOCK_STATUS   As Long = &H10

Public Const SC_MANAGER_ALL_ACCESS As Long = _
   (STANDARD_RIGHTS_REQUIRED _
    Or SC_MANAGER_CONNECT _
    Or SC_MANAGER_CREATE_SERVICE _
    Or SC_MANAGER_ENUMERATE_SERVICE _
    Or SC_MANAGER_LOCK _
    Or SC_MANAGER_QUERY_LOCK_STATUS _
    Or SC_MANAGER_MODIFY_BOOT_CONFIG)
    
'//
'// Service Types (Bit Mask)
'//
Public Const SERVICE_KERNEL_DRIVER = 1
Public Const SERVICE_FILE_SYSTEM_DRIVER = 2
Public Const SERVICE_ADAPTER = 4
Public Const SERVICE_RECOGNIZER_DRIVER = 8

Public Const SERVICE_DRIVER = (SERVICE_KERNEL_DRIVER _
    Or SERVICE_FILE_SYSTEM_DRIVER _
    Or SERVICE_RECOGNIZER_DRIVER)

Public Const SERVICE_WIN32_OWN_PROCESS = &H10
Public Const SERVICE_WIN32_SHARE_PROCESS = &H20
Public Const SERVICE_WIN32 = (SERVICE_WIN32_OWN_PROCESS _
    Or SERVICE_WIN32_SHARE_PROCESS)

Public Const SERVICE_INTERACTIVE_PROCESS = &H100

Public Const SERVICE_TYPE_ALL = (SERVICE_WIN32 _
    Or SERVICE_ADAPTER _
    Or SERVICE_DRIVER _
    Or SERVICE_INTERACTIVE_PROCESS)

'//
'// Start Type
'//
Public Const SERVICE_BOOT_START = 0
Public Const SERVICE_SYSTEM_START = 1
Public Const SERVICE_AUTO_START = 2
Public Const SERVICE_DEMAND_START = 3
Public Const SERVICE_DISABLED = 4

'//
'// Service object specific access type
'//
Public Const SERVICE_QUERY_CONFIG           As Long = &H1
Public Const SERVICE_CHANGE_CONFIG          As Long = &H2
Public Const SERVICE_QUERY_STATUS           As Long = &H4
Public Const SERVICE_ENUMERATE_DEPENDENTS   As Long = &H8
Public Const SERVICE_START                  As Long = &H10
Public Const SERVICE_STOP                   As Long = &H20
Public Const SERVICE_PAUSE_CONTINUE         As Long = &H40
Public Const SERVICE_INTERROGATE            As Long = &H80
Public Const SERVICE_USER_DEFINED_CONTROL   As Long = &H100

Public Const SERVICE_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED _
    Or SERVICE_QUERY_CONFIG _
    Or SERVICE_CHANGE_CONFIG _
    Or SERVICE_QUERY_STATUS _
    Or SERVICE_ENUMERATE_DEPENDENTS _
    Or SERVICE_START _
    Or SERVICE_STOP _
    Or SERVICE_PAUSE_CONTINUE _
    Or SERVICE_INTERROGATE _
    Or SERVICE_USER_DEFINED_CONTROL)

'**********************************
'**  Function Declarations:
Public Declare Function OpenSCManager& Lib "advapi32.dll" Alias "OpenSCManagerA" _
   (ByVal lpMachineName As String, _
    ByVal lpDatabaseName As String, _
    ByVal dwDesiredAccess As Long)

Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
   (ByVal lpString1 As String, ByVal lpString2 As Any) As Long

Declare Function CloseServiceHandle Lib "advapi32.dll" _
   (ByVal hSCObject As Long) As Long

Sub SetServiceStrings()
    vStatusStr = Array("Stopped", "Start Pending", "Stop Pending", _
        "Running", "Continue Pending", "Pause Pending", "Paused")
    vStartTypeStr = Array("Boot", "System", "Automatic", "Manual", "Disabled")
    bServicesInit = True
End Sub


