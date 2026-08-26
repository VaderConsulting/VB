Attribute VB_Name = "Pipe"
Option Explicit
' Copyright © 1997 by Desaware Inc. All Rights Reserved


Type STARTUPINFO
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

Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Declare Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, ByVal lpOverlapped As Long) As Long
Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long) As Long
Declare Function SetNamedPipeHandleState Lib "kernel32" (ByVal hNamedPipe As Long, lpMode As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long) As Long
Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lpFlags As Long, lpOutBufferSize As Long, lpInBufferSize As Long, lpMaxInstances As Long) As Long
Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Declare Function WriteFileAsync Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function ReadFileAsync Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" (ByVal lpNamedPipeName As String, ByVal nTimeOut As Long) As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, ByVal lpSecurityAttributes As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const STD_ERROR_HANDLE = -12&
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_ALL = &H10000000
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const INVALID_HANDLE_VALUE = -1
Public Const PIPE_ACCESS_INBOUND = &H1
Public Const PIPE_ACCESS_OUTBOUND = &H2
Public Const PIPE_ACCESS_DUPLEX = &H3
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const PIPE_TYPE_BYTE = &H0
Public Const PIPE_TYPE_MESSAGE = &H4
Public Const WAIT_FAILED = -1&
Public Const WAIT_OBJECT_0 = 0
Public Const WAIT_ABANDONED = &H80&
Public Const WAIT_ABANDONED_0 = &H80&
Public Const WAIT_TIMEOUT = &H102&
Public Const WAIT_IO_COMPLETION = &HC0&
Public Const STILL_ACTIVE = &H103&
Public Const INFINITE = -1&
Public Const ERROR_PIPE_CONNECTED = 535&

