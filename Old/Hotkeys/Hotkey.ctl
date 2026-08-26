VERSION 5.00
Begin VB.UserControl HotKey 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   630
   ScaleHeight     =   600
   ScaleWidth      =   630
   ToolboxBitmap   =   "HotKey.ctx":0000
End
Attribute VB_Name = "HotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================
' cHotKey
' Steve McMahon
' 09 June 1998
'
' A simple implementation of the hot key control.
' ======================================================================================

' ======================================================================================
' API declares:
' ======================================================================================
' Memory functions:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
' Send message:
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' Creating new windows:
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
' General window styles:
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_CHILD = &H40000000
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_TABSTOP = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Const WM_SETFONT = &H30
Private Const WM_SETHOTKEY = &H32
Private Const WM_USER = &H400

' Font
Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Const BITSPIXEL = 12
    Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
    Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
' CommonControls function
Private Declare Sub InitCommonControls Lib "COMCTL32.DLL" ()
Private Const HOTKEY_CLASS = "msctls_hotkey32"
Public Enum echkModifierKeys
   HOTKEYF_SHIFT = &H1
   HOTKEYF_CONTROL = &H2
   HOTKEYF_ALT = &H4
   HOTKEYF_EXT = &H8
   HOTKEYF_SHIFTCONTROL = &H3
   HOTKEYF_ALTSHIFT = &H5
   HOTKEYF_CONTROLALT = &H6
   HOTKEYF_CONTROLALTSHIFT = &H7
End Enum
Public Enum echkInvalidHotKeyModifiers
   HKCOMB_NONE = &H1
   HKCOMB_S = &H2
   HKCOMB_C = &H4
   HKCOMB_A = &H8
   HKCOMB_SC = &H10
   HKCOMB_SA = &H20
   HKCOMB_CA = &H40
   HKCOMB_SCA = &H80
End Enum
Public Enum echkHotKeyErrors
    eHotKeyAlreadyAssigned = vbObjectError + 1048 + 40
    eHotKeyInvalidWindow
    eHotKeyInvalidHotKey
    eHotKeyUnknownError
End Enum

Private Const HKM_SETHOTKEY = (WM_USER + 1)
Private Const HKM_GETHOTKEY = (WM_USER + 2)
Private Const HKM_SETRULES = (WM_USER + 3)

Private m_hWnd As Long
Private m_lfwCombInv As Long
Private m_lfwModInv As Long

' Font support:
Private m_tULF As LOGFONT
Private m_hFnt As Long

Public Sub SetApplicationHotKey(ByVal hwnd As Long)
Attribute SetApplicationHotKey.VB_Description = "Sets the current hot key as the hot key for a window with handle  hWnd."
Dim iR As Long
    iR = SendMessageByLong(hwnd, WM_SETHOTKEY, HotKeyAndModifier(), 0)
    Select Case iR
    Case 2
        Err.Raise eHotKeyAlreadyAssigned, App.EXEName & ".cHotKey", "Hot key previously assigned"
    Case 1
        ' success
    Case 0
        Err.Raise eHotKeyAlreadyAssigned, App.EXEName & ".cHotKey", "Invalid window for Hot key"
    Case -1
        Err.Raise eHotKeyInvalidHotKey, App.EXEName & ".cHotKey", "Invalid Hot key"
    Case Else
        Err.Raise eHotKeyUnknownError, App.EXEName & ".cHotKey", "Failed to set Hot key"
    End Select
End Sub

Public Property Let InvalidHotKeyOperation(ByVal eInvalidModifier As echkInvalidHotKeyModifiers, ByVal eAlternateModifier As echkModifierKeys, ByVal bState As Boolean)
Attribute InvalidHotKeyOperation.VB_Description = "Sets rules determining which key combinations are valid as a hotkey."
   If (bState) Then
      m_lfwCombInv = m_lfwCombInv Or (eInvalidModifier And &HFF&)
      m_lfwModInv = m_lfwModInv Or (eAlternateModifier And &HFF&)
   Else
      m_lfwCombInv = m_lfwCombInv And Not (eInvalidModifier And &HFF&)
      m_lfwModInv = m_lfwModInv And Not (eAlternateModifier And &HFF&)
   End If
   SendMessageByLong m_hWnd, HKM_SETRULES, m_lfwCombInv, m_lfwModInv
End Property

Public Property Get HotKey() As Long
Attribute HotKey.VB_Description = "Gets/sets the virtual key code of the key used in the hotkey combination."
Dim lT As Long
   lT = HotKeyAndModifier()
   HotKey = (lT And &HFF&)
End Property
Public Property Let HotKey(ByVal lKey As Long)
Dim lT As Long
   lT = HotKeyAndModifier
   If (lKey <> (lT And &HFF&)) Then
      lT = (lT And &HFF00) Or (lKey And &HFF&)
      SendMessageByLong m_hWnd, HKM_SETHOTKEY, lT, 0
      PropertyChanged "HotKey"
   End If
End Property
Public Property Get HotKeyModifier() As echkModifierKeys
Attribute HotKeyModifier.VB_Description = "Gets/sets the modifier code (i.e. Ctrl, Alt etc) of the key used in the hotkey combination."
Dim lT As Long
   lT = HotKeyAndModifier
   HotKeyModifier = (lT And &HFF00&) \ &H100&
End Property
Public Property Let HotKeyModifier(ByVal eModifier As echkModifierKeys)
Dim lT As Long
   lT = HotKeyAndModifier
   If ((lT And &HFF00F) \ &HFF&) <> (eModifier And &HFF&) Then
      lT = (eModifier And &HFF&) * &H100& Or (lT And &HFF&)
      SendMessageByLong m_hWnd, HKM_SETHOTKEY, lT, 0
      PropertyChanged "HotKeyModifier"
   End If
End Property
Public Property Get HotKeyAndModifier() As Long
Attribute HotKeyAndModifier.VB_Description = "Gets a word containing the virtual key code in the lobyte and the modifier in the hibyte -  used in some API functions."
Dim lT As Long
   HotKeyAndModifier = SendMessageByLong(m_hWnd, HKM_GETHOTKEY, 0, 0)
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Gets/sets the font for the control."
    Set Font = UserControl.Font
End Property
Public Property Set Font(sFont As StdFont)
Dim hFnt As Long
    If Not (UserControl.Font Is sFont) Then
        Set UserControl.Font = sFont
        ' Store a log font structure for this font:
        pOLEFontToLogFont sFont, UserControl.hDC, m_tULF
        ' Store old font handle:
        hFnt = m_hFnt
        ' Create a new version of the font:
        m_hFnt = CreateFontIndirect(m_tULF)
        ' Ensure the edit portion has the correct font:
        If (m_hWnd <> 0) Then
            SendMessage m_hWnd, WM_SETFONT, m_hFnt, 1
        End If
        ' Delete previous version, if we had one:
        If (hFnt <> 0) Then
            DeleteObject hFnt
        End If
        PropertyChanged "Font"
    End If
End Property
Private Sub pOLEFontToLogFont(fntThis As StdFont, hDC As Long, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer

    ' Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        sFont = fntThis.Name
        ' There is a quicker way involving StrConv and CopyMemory, but
        ' this is simpler!:
        For iChar = 1 To Len(sFont)
            .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
        Next iChar
        ' Based on the Win32SDK documentation:
        .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
    End With

End Sub

Private Function pCreateHotKeyWindow()
Static bNotFirst As Boolean
   If Not (bNotFirst) Then
      InitCommonControls
      bNotFirst = True
   End If
   m_hWnd = CreateWindowEx( _
         0, _
         HOTKEY_CLASS, _
         "", _
         WS_CHILD Or WS_VISIBLE, _
         0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, _
         UserControl.hwnd, _
         0, _
         App.hInstance, _
         0)
   If (m_hWnd <> 0) Then
      SetFocusAPI m_hWnd
      
   End If
End Function

Private Sub UserControl_GotFocus()
   If (m_hWnd <> 0) Then
      SetFocusAPI m_hWnd
   End If
End Sub

Private Sub UserControl_InitProperties()
   Set Font = UserControl.Ambient.Font
   pCreateHotKeyWindow
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   pCreateHotKeyWindow
    Dim sFnt As New StdFont
    sFnt.Name = "MS Sans Serif"
    sFnt.Size = 8
   Set Font = PropBag.ReadProperty("Font", sFnt)
   HotKey = PropBag.ReadProperty("HotKey", 0)
   HotKeyModifier = PropBag.ReadProperty("HotKeyModifier", HOTKEYF_ALT)
End Sub

Private Sub UserControl_Resize()
    If (m_hWnd <> 0) Then
        MoveWindow m_hWnd, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 1
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim sFnt As New StdFont
    sFnt.Name = "MS Sans Serif"
    sFnt.Size = 8
    PropBag.WriteProperty "Font", Font, sFnt
    PropBag.WriteProperty "HotKey", HotKey, 0
    PropBag.WriteProperty "HotKeyModifier", HotKeyModifier, HOTKEYF_ALT
End Sub
