VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form and Menu Effects"
   ClientHeight    =   2175
   ClientLeft      =   6735
   ClientTop       =   3060
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   2910
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1200
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1440
      Picture         =   "Form1.frx":00EA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdAddIcon 
      Caption         =   "Add Images"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.OptionButton optColor 
      BackColor       =   &H00404080&
      Caption         =   "Green"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.OptionButton optColor 
      BackColor       =   &H00FF0000&
      Caption         =   "Blue"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optColor 
      BackColor       =   &H000040C0&
      Caption         =   "Red"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdGradient 
      Caption         =   "Make Gradient"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Make Transparent"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Menu mnu1 
      Caption         =   "Menu1"
      Begin VB.Menu smnu1 
         Caption         =   "SubMenu1"
         Checked         =   -1  'True
      End
      Begin VB.Menu smnu2 
         Caption         =   "SubMenu2"
      End
      Begin VB.Menu smnu3 
         Caption         =   "SubMenu3"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnu2 
      Caption         =   "Menu2"
      Begin VB.Menu smnu4 
         Caption         =   "SubMenu4"
      End
      Begin VB.Menu smnu5 
         Caption         =   "SubMenu5"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'---------------------------------------------------
' Transparent Form Related.
'---------------------------------------------------
'
Const GWL_STYLE = (-16)
Const GWL_EXSTYLE = (-20)
Const WS_EX_TRANSPARENT = &H20&
Const SWP_FRAMECHANGED = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "User32" _
    Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal lIndex As Long) As Long

Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
'
'---------------------------------------------------
' Used to Remove the Close Button.
'---------------------------------------------------
'
Const MAXIMIXE_BUTTON = &HFFFEFFFF
Const MINIMIZE_BUTTON = &HFFFDFFFF
Const MF_BYPOSITION = &H400
'
' Enumeration used when calling pRemoveMenu.
'
Enum RemoveMenuEnum
    rmMove = 1
    rmSize = 2
    rmMinimize = 3
    rmMaximize = 4
    rmClose = 6
End Enum

Private Declare Function GetSystemMenu Lib "User32" _
    (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "User32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'
'---------------------------------------------------
' Menu Image Related.
'---------------------------------------------------
'
Private Const MF_BITMAP = &H4&
'
' Gets the handle to the menu assigned to the given window.
Private Declare Function GetMenu Lib "User32" _
    (ByVal hWnd As Long) As Long
'
' Get the handle to the drop-down menu or submenu
' activated by the specified menu item.
Private Declare Function GetSubMenu Lib "User32" _
    (ByVal hMenu As Long, ByVal nPos As Long) As Long
'
' Get the menu item identifier of a menu item located at
' the specified position in a menu.
Private Declare Function GetMenuItemID Lib "User32" _
    (ByVal hMenu As Long, ByVal nPos As Long) As Long
'
' Associates the specified bitmap with a menu item. Whether the
' menu item is checked or unchecked, the system displays the
' appropriate bitmap next to the menu item.
Private Declare Function SetMenuItemBitmaps Lib "User32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
    ByVal hBitmapChecked As Long) As Long
'
' Determines the number of items in the specified menu.
Private Declare Function GetMenuItemCount Lib "User32" _
    (ByVal hMenu As Long) As Long

Private Sub cmdAddIcon_Click()
Dim l           As Long
Dim j           As Long
Dim lMenuHnd    As Long
Dim lSubMenuHnd As Long
Dim lMenuCnt    As Long
Dim lSubMenuCnt As Long
Dim lSubMenuID  As Long
'
' Get the menu handle for the current form.
'
lMenuHnd = GetMenu(Me.hWnd)
'
' Find out how many menus there are.
'
lMenuCnt = GetMenuItemCount(lMenuHnd)
'
' Process each menu entry.
'
For l = 0 To lMenuCnt - 1
    '
    ' Get the next submenu handle for this menu.
    '
    lSubMenuHnd = GetSubMenu(lMenuHnd, l)
    '
    ' Find out how many entries are in this submenu.
    '
    lSubMenuCnt = GetMenuItemCount(lSubMenuHnd)
    '
    ' Process each submenu entry.
    '
    For j = 0 To lSubMenuCnt - 1
        '
        ' Get each entry ID for the current submenu
        '
        lSubMenuID = GetMenuItemID(lSubMenuHnd, j)
        '
        ' Add two pictures - one for checked and one for unchecked
        '
        ' If either the hBitmapUnchecked or hBitmapChecked parameter is NULL,
        ' the system displays nothing next to the menu item for the corresponding
        ' check state. If both parameters are NULL, the system displays the
        ' default check-mark bitmap when the item is checked, and removes the
        ' bitmap when the item is not checked.
        '
        ' When the menu is destroyed, these bitmaps are not destroyed; it is
        ' up to the application to destroy them.
        '
        ' The checked and unchecked bitmaps should be monochrome. The system
        ' uses the Boolean AND operator to combine bitmaps with the menu so that
        ' the white part becomes transparent and the black part becomes the menu-item
        ' color. If you use color bitmaps, the results may be undesirable.
        '
        ' Use the GetSystemMetrics function with the CXMENUCHECK and CYMENUCHECK
        ' values to retrieve the bitmap dimensions.
        '
        Call SetMenuItemBitmaps(lMenuHnd, lSubMenuID, MF_BITMAP, Picture2.Picture, Picture1.Picture)
    Next
Next
End Sub
Private Sub cmdClear_Click()
'
' Set the extended style bits of the form.
' Re-show the form.
'
Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
End Sub
Private Sub cmdGradient_Click()
Dim intLoop As Integer

With Me
    .AutoRedraw = True
    .DrawStyle = vbInsideSolid
    .DrawMode = vbCopyPen
    .ScaleMode = vbPixels
    .DrawWidth = 2
    .ScaleHeight = 256
End With

For intLoop = 0 To 255
    If optColor(0) Then
        'Red
        Me.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    ElseIf optColor(1) Then
        'Blue
        Me.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Else
        'Green
        Me.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    End If
Next
End Sub
Private Sub Form_Load()
Dim lMenu As Long
Dim lRes  As Long

Picture1.Visible = False
Picture2.Visible = False

Call pRemoveMenu(Me, rmClose)
Call pRemoveMenu(Me, rmMaximize)
Call pRemoveMenu(Me, rmMinimize)
Call pRemoveMenu(Me, rmSize)
Call pRemoveMenu(Me, rmMove)

End Sub
Private Sub pRemoveMenu(ByVal TargetForm _
    As Form, ByVal MenuToRemove As RemoveMenuEnum)
'
' Remove the specified menu item from the control menu
' and the corresponding functionality from the form.
'
' Parameters
'   TargetForm   - Form to perform the operation on.
'   MenuToRemove - Enum specifying which menu to remove.
'
Dim lSysMenu As Long
Dim lStyle   As Long
'
' GetSystemMenu allows access to the window menu
' (also known as the System Menu or the Control
' Menu) for copying and modifying.
'
lSysMenu = GetSystemMenu(TargetForm.hWnd, 0&)
Call RemoveMenu(lSysMenu, MenuToRemove, MF_BYPOSITION)

Select Case MenuToRemove
    Case rmClose
        '
        ' Remove the Close menu and the separator over it.
        ' RemoveMenu deletes a menu item from the specified menu.
        '
        Call RemoveMenu(lSysMenu, MenuToRemove - 1, MF_BYPOSITION)
    Case rmMinimize, rmMaximize
        '
        ' Get the current window style and set the new style.
        '
        lStyle = GetWindowLong(TargetForm.hWnd, GWL_STYLE)

        If MenuToRemove = rmMaximize Then
            lStyle = lStyle And MAXIMIXE_BUTTON
        Else
            lStyle = lStyle And MINIMIZE_BUTTON
        End If

        Call SetWindowLong(TargetForm.hWnd, GWL_STYLE, lStyle)
End Select
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub
