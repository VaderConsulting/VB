Attribute VB_Name = "Module1"
Public Declare Function AllocConsole Lib "kernel32" () As Long
Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Public Declare Function FreeConsole Lib "kernel32" () As Long
Public Const STD_OUTPUT_HANDLE = -11&
Public Const STD_ERROR_HANDLE = -12&
Public Const STD_INPUT_HANDLE = -10&
Public Const INVALID_HANDLE_VALUE = -1

Public Declare Function CreateConsoleScreenBuffer Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes, ByVal dwFlags As Long, lpScreenBufferData As Any) As Long
Public Const GENERIC_ALL = &H10000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_SHARE_NONE = 0      ' Non-standard, but my 2 cents worth

Public Declare Function SetConsoleActiveScreenBuffer Lib "kernel32" (ByVal hConsoleOutput As Long) As Long
Public Const CONSOLE_TEXTMODE_BUFFER = 1

Public Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long
Public Const ENABLE_ECHO_INPUT = &H4
Public Const ENABLE_LINE_INPUT = &H2
Public Const ENABLE_PROCESSED_INPUT = &H1
Public Const ENABLE_WINDOW_INPUT = &H8
Public Const ENABLE_MOUSE_INPUT = &H10
Public Const ENABLE_PROCESSED_OUTPUT = &H1
Public Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

Public Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
Public Declare Function SetConsoleCP Lib "kernel32" (ByVal wCodePageID As Long) As Long

Public Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByRef lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, ByRef lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
