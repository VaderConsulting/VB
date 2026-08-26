Attribute VB_Name = "mWindowPlay"
Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060
Private Const WM_SHOWWINDOW = &H18
Private Const WM_CANCELMODE = &H1F
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Function EnumWindowsProc( _
        ByVal hwnd As Long, _
        ByVal lParam As Long _
    ) As Long
Dim tR As RECT
   If IsWindowVisible(hwnd) <> 0 Then
      If IsIconic(hwnd) = 0 Then
         If WindowCaption(hwnd) <> "" Then
            GetWindowRect hwnd, tR
            MoveWindow hwnd, tR.Left + Rnd * 16 - 8, tR.Top + Rnd * 16 - 8, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
            If (lParam = 2) Then
               ' Watch it...
               EnumChildWindows hwnd, plAddressOf(AddressOf ChildWindowProc), 2
            End If
         End If
      End If
   End If
   EnumWindowsProc = 1
End Function
Private Function ChildWindowProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim tR As RECT
   GetWindowRect hwnd, tR
   OffsetRect tR, -tR.Left, -tR.Top
   MoveWindow hwnd, tR.Left + Rnd * 16 - 8, tR.Top + Rnd * 16 - 8, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
   ChildWindowProc = 1
End Function

Private Function WindowCaption(ByVal hwnd As Long) As String
Dim sBuf As String
Dim lLen As Long
   lLen = GetWindowTextLength(hwnd)
   If (lLen > 0) Then
      sBuf = String$(lLen + 1, 0)
      GetWindowText hwnd, sBuf, lLen + 1
      WindowCaption = Left$(sBuf, lLen)
   End If
End Function
Private Function ClassName(ByVal hwnd As Long) As String
Dim sBuf As String
Dim lLen As Long
Dim iPos As Long
   sBuf = String$(255, 0)
   lLen = GetClassName(hwnd, sBuf, 255)
   If (lLen > 0) Then
      iPos = InStr(sBuf, vbNullChar)
      If (iPos > 0) Then
         ClassName = Left$(sBuf, iPos - 1)
      Else
         ClassName = sBuf
      End If
   End If
End Function
Public Sub DoEarthQuake(ByVal iType As Long)
   EnumWindows plAddressOf(AddressOf EnumWindowsProc), iType
End Sub
Private Function plAddressOf(ByVal lPtr As Long) As Long
   plAddressOf = lPtr
End Function

Public Sub KillAssistant()
   EnumWindows plAddressOf(AddressOf KillAssistantProc), 0
End Sub
Private Function KillAssistantProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim tR As RECT
Dim i As Long
Dim lT As Long
   If (ClassName(hwnd) = "MSOASSISTANT") Then
      GetWindowRect hwnd, tR
      For i = tR.Left To Screen.Width \ Screen.TwipsPerPixelX + Abs(tR.Right - tR.Left) Step 16
         tR.Left = tR.Left + 16
         InflateRect tR, -1, -1
         lT = GetTickCount
         Do While GetTickCount < lT + 50
            DoEvents
         Loop
         MoveWindow hwnd, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
      Next i
      SendMessageByLong hwnd, WM_CANCELMODE, 0, 0
      SendMessageByLong hwnd, WM_SHOWWINDOW, 0, 0
      KillAssistantProc = 0
   Else
      KillAssistantProc = 1
   End If
End Function
