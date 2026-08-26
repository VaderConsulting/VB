Attribute VB_Name = "modIdle"
Public Type LastInputInfo
    cbSize As Long
    dwTime As Long
End Type

Public Declare Function GetLastInputInfo Lib "User32" (lastinfo As LastInputInfo) As Long
Public Declare Function GetTickCount& Lib "kernel32" ()


