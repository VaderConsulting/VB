Attribute VB_Name = "ModScreenCapture"

Private Type POINTAPI
    x As Long
    y As Long
    End Type

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Global Const SRCCOPY = &HCC0020
Public Function GetCursorPos_X() As Long
    Dim ptAPI As POINTAPI
    GetCursorPos ptAPI
    GetCursorPos_X = ptAPI.x
End Function


Public Function GetCursorPos_Y() As Long
    Dim ptAPI As POINTAPI
    GetCursorPos ptAPI
    GetCursorPos_Y = ptAPI.y
End Function
