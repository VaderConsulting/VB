Attribute VB_Name = "Module1"
Option Explicit

Private Const DeskTop = "winsta0\default"

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
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

Public Type OSVERSIONINFO
    dwOSVersionInfoSize  As Long
    dwMajorVersion       As Long
    dwMinorVersion       As Long
    dwBuildNumber        As Long
    dwPlatformId         As Long
    szCSDVersion         As String * 128
End Type

Private Declare Function CreateProcessWithLogon Lib "advapi32.dll" _
                    Alias "CreateProcessWithLogonW" ( _
                        ByVal lpUsername As Long, _
                        ByVal lpDomain As Long, _
                        ByVal lpPassword As Long, _
                        ByVal dwLogonFlags As Long, _
                        ByVal lpApplicationName As Long, _
                        ByVal lpCommandLine As Long, _
                        ByVal dwCreationFlags As Long, _
                        ByVal lpEnvironment As Long, _
                        ByVal lpCurrentDirectory As Long, _
                        lpStartupInfo As STARTUPINFO, _
                        lpProcessInformation As PROCESS_INFORMATION) As Long


Private Declare Function CreateProcessAsUser Lib "advapi32.dll" _
                    Alias "CreateProcessAsUserA" ( _
                        ByVal hToken As Long, _
                        ByVal lpApplicationName As String, _
                        ByVal lpCommandLine As String, _
                        ByVal lpProcessAttributes As Long, _
                        ByVal lpThreadAttributes As Long, _
                        ByVal bInheritHandles As Long, _
                        ByVal dwCreationFlags As Long, _
                        ByVal lpEnvironment As Long, _
                        ByVal lpCurrentDirectory As String, _
                        lpStartupInfo As STARTUPINFO, _
                        lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function LogonUser Lib "advapi32" _
                    Alias "LogonUserA" ( _
                        ByVal lpszUsername As String, _
                        ByVal lpszDomain As String, _
                        ByVal lpszPassword As String, _
                        ByVal dwLogonType As Long, _
                        ByVal dwLogonProvider As Long, _
                        phToken As Long) As Long

Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Private Const LOGON_WITH_PROFILE = &H1
Private Const LOGON32_LOGON_INTERACTIVE = 2
Private Const LOGON32_PROVIDER_DEFAULT = 0
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const CREATE_NEW_CONSOLE = &H10
Private MyToken As Long
Private si As STARTUPINFO
Private pi As PROCESS_INFORMATION
Public TheApplication As String
Public Os As Integer

'********************************************************************
'                   Installation for Windows 2000
'********************************************************************
Public Function Win2KSetup(ByVal UserName As String, ByVal PassWord As String, _
    ByVal Domain As String, ByVal CommandLine As String) As Long
    
    Dim lpUName As Long, lpDomain As Long
    Dim lpPWD As Long, lpAppName As Long
    Dim lpCmdLine As Long, hRet As Long
    
    si.cb = LenB(si)
    si.lpDesktop = StrPtr(DeskTop)
    
    lpUName = StrPtr(UserName)
    lpDomain = StrPtr(Domain)
    lpPWD = StrPtr(PassWord)
    lpAppName = StrPtr(TheApplication)
    lpCmdLine = StrPtr(CommandLine)
    
    hRet = CreateProcessWithLogon(lpUName, lpDomain, lpPWD, LOGON_WITH_PROFILE, lpAppName, lpCmdLine, 0&, 0&, 0&, si, pi)
    Win2KSetup = hRet

End Function

'********************************************************************
'                   Installation for Windows NT4.0
'********************************************************************
Public Function WinNTSetup(ByVal UserName As String, ByVal PassWord, ByVal Domain, ByVal CommandLine, ByVal CurrentDirectory As String) As Long
    Dim retval As Long
    
    retval = LogonUser(UserName, Domain, PassWord, LOGON32_LOGON_INTERACTIVE, _
    LOGON32_PROVIDER_DEFAULT, MyToken)
    If retval = 0 Then
        WinNTSetup = Err.LastDllError
        Exit Function
    End If
    
    retval = CreateProcessAsUser(MyToken, TheApplication, CommandLine, 0&, 0&, False, _
    NORMAL_PRIORITY_CLASS + CREATE_NEW_CONSOLE, _
    0&, CurrentDirectory, si, pi)
    If retval = 0 Then
        WinNTSetup = Err.LastDllError
        Exit Function
    End If
    WinNTSetup = retval

End Function


