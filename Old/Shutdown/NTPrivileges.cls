VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
END
Attribute VB_Name = "clsNTPrivileges"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Option Explicit

Private Const ANYSIZE_ARRAY = 1

Private Type LUID
    lowpart As Long
    highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
        ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) _
        As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias _
        "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, _
        lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" ( _
        ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
        NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
        PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF&

Private Const TOKEN_ASSIGN_PRIMARY = &H1&
Private Const TOKEN_DUPLICATE = &H2&
Private Const TOKEN_IMPERSONATE = &H4&
Private Const TOKEN_QUERY = &H8&
Private Const TOKEN_QUERY_SOURCE = &H10&
Private Const TOKEN_ADJUST_PRIVILEGES = &H20&
Private Const TOKEN_ADJUST_GROUPS = &H40&
Private Const TOKEN_ADJUST_DEFAULT = &H80&
Private Const TOKEN_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or _
                                 TOKEN_ASSIGN_PRIMARY Or _
                                 TOKEN_DUPLICATE Or _
                                 TOKEN_IMPERSONATE Or _
                                 TOKEN_QUERY Or _
                                 TOKEN_QUERY_SOURCE Or _
                                 TOKEN_ADJUST_PRIVILEGES Or _
                                 TOKEN_ADJUST_GROUPS Or _
                                 TOKEN_ADJUST_DEFAULT
Private Const TOKEN_READ = STANDARD_RIGHTS_READ Or TOKEN_QUERY
Private Const TOKEN_WRITE = STANDARD_RIGHTS_WRITE Or _
                            TOKEN_ADJUST_PRIVILEGES Or _
                            TOKEN_ADJUST_GROUPS Or _
                            TOKEN_ADJUST_DEFAULT
Private Const TOKEN_EXECUTE = STANDARD_RIGHTS_EXECUTE

Private Const SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege"
Private Const SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege"
Private Const SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege"
Private Const SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"
Private Const SE_UNSOLICITED_INPUT_NAME = "SeUnsolicitedInputPrivilege"
Private Const SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege"
Private Const SE_TCB_NAME = "SeTcbPrivilege"
Private Const SE_SECURITY_NAME = "SeSecurityPrivilege"
Private Const SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege"
Private Const SE_LOAD_DRIVER_NAME = "SeLoadDriverPrivilege"
Private Const SE_SYSTEM_PROFILE_NAME = "SeSystemProfilePrivilege"
Private Const SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege"
Private Const SE_PROF_SINGLE_PROCESS_NAME = "SeProfileSingleProcessPrivilege"
Private Const SE_INC_BASE_PRIORITY_NAME = "SeIncreaseBasePriorityPrivilege"
Private Const SE_CREATE_PAGEFILE_NAME = "SeCreatePagefilePrivilege"
Private Const SE_CREATE_PERMANENT_NAME = "SeCreatePermanentPrivilege"
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
Private Const SE_RESTORE_NAME = "SeRestorePrivilege"
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Private Const SE_DEBUG_NAME = "SeDebugPrivilege"
Private Const SE_AUDIT_NAME = "SeAuditPrivilege"
Private Const SE_SYSTEM_ENVIRONMENT_NAME = "SeSystemEnvironmentPrivilege"
Private Const SE_CHANGE_NOTIFY_NAME = "SeChangeNotifyPrivilege"
Private Const SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege"

Private Const SE_PRIVILEGE_ENABLED_BY_DEFAULT = &H1&
Private Const SE_PRIVILEGE_ENABLED = &H2&
Private Const SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000

Private Const ERROR_SUCCESS = 0&

Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" _
                Alias "InitiateSystemShutdownA" ( _
                ByVal lpMachineName As String, ByVal lpMessage As String, _
                ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, _
                ByVal bRebootAfterShutdown As Long) As Boolean
Private Declare Function AbortSystemShutdown Lib "advapi32.dll" _
                Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long

Private tkpSaved As TOKEN_PRIVILEGES
Private hToken As Long
Public Sub AbortShutDown(ByVal cComputerName As String)

    Dim bDummy As Boolean

    If Len(cComputerName) = 0 Then
        bDummy = SetPrivilege(SE_SHUTDOWN_NAME)
    Else
        bDummy = SetPrivilege(SE_REMOTE_SHUTDOWN_NAME)
    End If
    If bDummy Then
        ' Display the shutdown dialog box and start the time-out countdown.
        If AbortSystemShutdown(cComputerName) Then
            ' Disable shutdown privilege.
            ResetPrivilege
        End If
    End If

End Sub


Public Sub ResetPrivilege()

    Dim tkp As TOKEN_PRIVILEGES
    Dim lDummy As Long
    Dim lRequired As Long

    lDummy = AdjustTokenPrivileges(hToken, False, tkpSaved, Len(tkp), tkp, lRequired)

End Sub
Public Function SetPrivilege(ByVal cPrivilegeName As String) As Boolean

    Dim tkp As TOKEN_PRIVILEGES
    Dim lDummy As Long
    Dim lRequired As Long
    Dim bRetVal As Boolean

    ' Get the current process token handle so we can get shutdown privilege.
    If OpenProcessToken(GetCurrentProcess(), _
                        TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) Then
        ' Get the LUID for shutdown privilege.
        If LookupPrivilegeValue(vbNullString, cPrivilegeName, tkp.Privileges(0).pLuid) > 0 Then
            tkp.PrivilegeCount = 1      ' one privilege to set
            tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    
            ' Get shutdown privilege for this process.
            lDummy = AdjustTokenPrivileges(hToken, False, tkp, Len(tkpSaved), tkpSaved, lRequired)
            ' Cannot test the return value of AdjustTokenPrivileges.
            lDummy = GetLastError
            If lDummy = ERROR_SUCCESS Then
                bRetVal = True
            Else
                Debug.Print lDummy
            End If
        End If
    End If
    SetPrivilege = bRetVal

End Function
Public Sub ShutDown(ByVal cComputerName As String, ByVal cMessage As String, _
                    ByVal lTimeout As Long, ByVal bForceAppsClosed As Boolean, _
                    ByVal bRebootAfterShutdown As Boolean)

    Dim bDummy As Boolean

    If Len(cComputerName) = 0 Then
        bDummy = SetPrivilege(SE_SHUTDOWN_NAME)
    Else
        bDummy = SetPrivilege(SE_REMOTE_SHUTDOWN_NAME)
    End If
    If bDummy Then
        ' Display the shutdown dialog box and start the time-out countdown.
        If InitiateSystemShutdown(cComputerName, cMessage, lTimeout, _
                                  bForceAppsClosed, bRebootAfterShutdown) Then
            ' Disable shutdown privilege.
            ResetPrivilege
        End If
    End If

End Sub
