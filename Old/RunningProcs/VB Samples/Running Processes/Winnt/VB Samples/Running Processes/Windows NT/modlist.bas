Attribute VB_Name = "modList"
Option Explicit

Public Declare Function GetVersionExA Lib "kernel32" _
   (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type OSVERSIONINFO
   dwOSVersionInfoSize  As Long 'length in bytes of the structure.
   dwMajorVersion       As Long 'Major Version Number
   dwMinorVersion       As Long 'Minor Version Number
   dwBuildNumber        As Long 'Build Version Number
   dwPlatformId         As Long 'Operating System Running, see below
   szCSDVersion As String * 128 'Windows NT: Contains a null-terminated string,
                                'such as "Service Pack 3", that indicates the latest
                                'Service Pack installed on the system.
                                'If no Service Pack has been installed, the string is empty.
                                'Windows 9x: Contains a null-terminated string that provides
                                'arbitrary additional information about the operating system
End Type

Public Const hNull = 0

Public Const VER_PLATFORM_WIN32s = 0            'Win32s on Windows 3.1.
Public Const VER_PLATFORM_WIN32_WINDOWS = 1     'Win32 on Windows 95 or Windows 98.
                                                'For Windows 95, dwMinorVersion is 0.
                                                'For Windows 98, dwMinorVersion is 1.
Public Const VER_PLATFORM_WIN32_NT = 2          'Win32 on Windows NT.

Public Function GetVersion() As Long
'
' Returns the Operating System used
' 1 = Windows 9x
' 2 = Windows NT
'
Dim osinfo   As OSVERSIONINFO
Dim retvalue As Integer
    
With osinfo
    .dwOSVersionInfoSize = 148
    .szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    GetVersion = .dwPlatformId
End With
End Function



