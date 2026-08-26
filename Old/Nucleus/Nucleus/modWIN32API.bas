Attribute VB_Name = "modWIN32API"
Option Explicit

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const ANYSIZE_ARRAY = 1
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const SE_DEBUG_NAME = "SeDebugPrivilege"
Public Const SE_PRIVILEGE_ENABLED = &H2

Type LARGE_INTEGER
  lowpart As Long
  highpart As Long
End Type

Type Luid
  lowpart As Long
  highpart As Long
End Type

Type LUID_AND_ATTRIBUTES
  pLuid As Luid
  Attributes As Long
End Type

Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessID As Long
  dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As Luid) As Long
Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Function GetProcessExitCode(Proc As clsProcess) As Long
  Dim Misc As Long
  Call GetExitCodeProcess(Proc.hProcess, Misc)
  GetProcessExitCode = Misc
End Function

Public Function CleanupProcess(Proc As clsProcess) As Boolean
  CleanupProcess = True
  Call CloseHandle(Proc.hThread)
  Call CloseHandle(Proc.hProcess)
End Function

Public Function ProcessCompleted(Proc As clsProcess) As Boolean
  ProcessCompleted = (WaitForSingleObject(Proc.hProcess, 0) = 0)
End Function

Public Function ExecCmd(ByVal cmdline As String, oProc As clsProcess) As Boolean
  Dim Proc As PROCESS_INFORMATION
  Dim start As STARTUPINFO
  Dim Ret As Long

  ExecCmd = False
  ' Initialize the STARTUPINFO structure:
  start.cb = Len(start)

  ' Start the shelled application:
  Ret = CreateProcessA(0&, cmdline, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, Proc)

  If (Ret <> 1) Then Exit Function ' Launch Failed
    
  'Add Process to process monitor list
  oProc.SetProcessInfo Proc.hProcess, Proc.hThread, Proc.dwProcessID, Proc.dwThreadID
  
  ExecCmd = True
End Function

Public Sub KillProcess(ApplicationPID As Long)
  Dim hProcessID As Long         ' Handle to your process you are going to terminate.
  Dim hProcess As Long           ' Handle to your current process
  Dim hToken As Long             ' Handle to your process token.
  Dim lPrivilege As Long         ' Privilege to enable/disable
  Dim iPrivilegeflag As Boolean  ' Flag whether to enable/disable the privilege of concern.
  Dim lResult As Long            ' Result call of various APIs.
  
  ' set the incoming PID to our internal variable
  hProcessID = ApplicationPID
  
  ' get our current process handle
  hProcess = GetCurrentProcess
  
  ' show our handle just for fun
  'frmTerm02.List1.AddItem "Current (Pseudo) Process Handle : " & Hex(hProcess)
  
  ' open the tokens for this process (Term02.exe)
  lResult = OpenProcessToken(hProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
  
  ' if OpenProcessToken fails, the return result is zero, test for success here
  
  If (lResult = 0) Then
    'frmTerm02.List1.AddItem "Error: Unable To Open Process Token : " & Err.LastDllError
    CloseHandle (hToken)
    Exit Sub
  Else
    ' show success
    'frmTerm02.List1.AddItem "Opened Process Token : " & hToken
  End If
  
  ' Now that you have the token for this process, you want to set
  ' the SE_DEBUG_NAME privilege.
  
  lResult = SetPrivilege(hToken, SE_DEBUG_NAME, True)
  
  ' Make sure you could set the privilege on this token.
  
  If (lResult = False) Then
    'frmTerm02.List1.AddItem "Error : Could Not Set SeDebug Privilege on Token Handle"
    CloseHandle (hToken)
    Exit Sub
  Else
     'frmTerm02.List1.AddItem "Set SeDebug Privilege On Token Handle"
  End If
  
  ' Now that you have changed the privileges on the token,
  ' have some fun. You can now get a process handle to the
  ' process ID that you passed into this program, and
  ' demand whatever access you want on it!
  
  hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
  
  ' Make sure you opened the process so you can do stuff with it
  If (hProcess = Null) Then
    'frmTerm02.List1.AddItem "Error : Unable To Open Process : " & Err.LastDllError
    CloseHandle (hToken)
    Exit Sub
  Else
    'frmTerm02.List1.AddItem "Opened Process : " & hProcess
  End If
  
  ' Now turn the SE_DEBUG_PRIV back off,
  lResult = SetPrivilege(hToken, SE_DEBUG_NAME, False)
  
  ' Make sure you succeeded in reversing the privilege!
  If (lResult = False) Then
    'frmTerm02.List1.AddItem "Error : Unable To Disable SeDebug Privilege On Token Handle"
    CloseHandle (hProcess)
    CloseHandle (hToken)
    Exit Sub
  Else
     'frmTerm02.List1.AddItem "Disabled SeDebug Privilege On Token Handle"
  End If
  
  ' Now you want to kill the application, which you can do since
  ' your process handle to the application includes full access to
  ' romp and roam - you got the process handle when you had the
  ' SE_DEBUG_NAME privilege enabled!
  lResult = TerminateProcess(hProcess, 0)
  
  ' Let's see the result, and go from there.
  If (lResult = 0) Then
    'frmTerm02.List1.AddItem "Error : Unable To Terminate Application : " & Err.LastDllError
    CloseHandle (hProcess)
    CloseHandle (hToken)
    Exit Sub
  Else
    'frmTerm02.List1.AddItem "Terminated Application"
  End If
  
  ' Close our handles and get out of here.
  CloseHandle (hProcess)
  CloseHandle (hToken)
  
  ' Finally, let the user know that you have completed the
  ' SeDebug Sample.
  'frmTerm02.List1.AddItem "SeDebug Sample Completed"
  
End Sub

' The SetPrivilege function will accept a handle to a token, a
' privilege, and a flag to either enable/disable that privilege. The
' function will attempt to perform the desired action upon the token
' returning TRUE if it succeeded, or FALSE if it failed.

Private Function SetPrivilege(hToken As Long, Privilege As String, bSetFlag As Boolean) As Boolean

   Dim TP As TOKEN_PRIVILEGES          ' Used in getting the current
                                       ' token privileges
   Dim TPPrevious As TOKEN_PRIVILEGES  ' Used in setting the new
                                       ' token privileges
   Dim Luid As Luid                    ' Stores the Local Unique
                                       ' Identifier - refer to MSDN
   Dim cbPrevious As Long              ' Previous size of the
                                       ' TOKEN_PRIVILEGES structure
   Dim lResult As Long                 ' Result of various API calls

   ' Grab the size of the TOKEN_PRIVILEGES structure,
   ' used in making the API calls.
   cbPrevious = Len(TP)

   ' Grab the LUID for the request privilege.
   lResult = LookupPrivilegeValue("", Privilege, Luid)

   ' If LoopupPrivilegeValue fails, the return result will be zero.
   ' Test to make sure that the call succeeded.
   If (lResult = 0) Then
      SetPrivilege = False
   End If

   ' Set up basic information for a call.
   ' You want to retrieve the current privileges
   ' of the token under concern before you can modify them.
   TP.PrivilegeCount = 1
   TP.Privileges(0).pLuid = Luid
   TP.Privileges(0).Attributes = 0
   SetPrivilege = lResult

   ' You need to acquire the current privileges first
   lResult = AdjustTokenPrivileges(hToken, -1, TP, Len(TP), TPPrevious, cbPrevious)

   ' If AdjustTokenPrivileges fails, the return result is zero,
   ' test for success.
   If (lResult = 0) Then
      SetPrivilege = False
   End If

   ' Now you can set the token privilege information
   ' to what the user is requesting.
   TPPrevious.PrivilegeCount = 1
   TPPrevious.Privileges(0).pLuid = Luid

   ' either enable or disable the privilege,
   ' depending on what the user wants.
   Select Case bSetFlag
      Case True: TPPrevious.Privileges(0).Attributes = _
                 TPPrevious.Privileges(0).Attributes Or _
                 (SE_PRIVILEGE_ENABLED)
      Case False: TPPrevious.Privileges(0).Attributes = _
                  TPPrevious.Privileges(0).Attributes Xor _
                  (SE_PRIVILEGE_ENABLED And _
                  TPPrevious.Privileges(0).Attributes)
   End Select

   ' Call adjust the token privilege information.
   lResult = AdjustTokenPrivileges(hToken, -1, TPPrevious, cbPrevious, TP, cbPrevious)

   ' Determine your final result of this function.
   If (lResult = 0) Then
      ' You were not able to set the privilege on this token.
      SetPrivilege = False
   Else
      ' You managed to modify the token privilege
      SetPrivilege = True
   End If

End Function

