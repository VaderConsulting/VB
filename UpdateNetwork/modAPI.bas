Attribute VB_Name = "modAPI"
Option Explicit

Public Const WM_SETTINGCHANGE = &H1A
Public Const HWND_BROADCAST = &HFFFF
Public Const SMTO_ABORTIFHUNG = &H2

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Integer
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long



