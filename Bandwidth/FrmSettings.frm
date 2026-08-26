VERSION 5.00
Begin VB.Form FrmSettings 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "FrmSettings.frx":0000
   ScaleHeight     =   5325
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "General Options"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   10
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   480
      Width           =   1785
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   7080
      Picture         =   "FrmSettings.frx":110FF
      Tag             =   "Close"
      ToolTipText     =   "Close"
      Top             =   50
      Width           =   315
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Colors"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3600
      TabIndex        =   9
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   5640
      TabIndex        =   8
      Tag             =   "TitleColor"
      Top             =   4005
      Width           =   1485
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Apply"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   5640
      TabIndex        =   7
      Tag             =   "TitleColor"
      Top             =   3645
      Width           =   1485
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Colors"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3600
      TabIndex        =   6
      Tag             =   "TitleColor"
      Top             =   3900
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "System Tray Options"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3720
      TabIndex        =   5
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   480
      Width           =   1785
   End
   Begin VB.Line Line2 
      X1              =   5400
      X2              =   5640
      Y1              =   1920
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   5400
      X2              =   5640
      Y1              =   2400
      Y2              =   2160
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "The Non-White Areas are transparent"
      Height          =   735
      Left            =   5760
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   4
      Left            =   3720
      Picture         =   "FrmSettings.frx":11621
      Top             =   2760
      Width           =   180
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   3
      Left            =   3720
      Picture         =   "FrmSettings.frx":11A0E
      Top             =   2280
      Width           =   180
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   2
      Left            =   3720
      Picture         =   "FrmSettings.frx":11DFB
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   1
      Left            =   3720
      Picture         =   "FrmSettings.frx":121E8
      Top             =   1320
      Width           =   180
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   0
      Left            =   3720
      Picture         =   "FrmSettings.frx":125D5
      Top             =   840
      Width           =   180
   End
   Begin VB.Image Image18 
      Height          =   240
      Index           =   3
      Left            =   5040
      Picture         =   "FrmSettings.frx":129C2
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image Image18 
      Height          =   240
      Index           =   2
      Left            =   4680
      Picture         =   "FrmSettings.frx":12F4C
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image Image18 
      Height          =   240
      Index           =   1
      Left            =   4320
      Picture         =   "FrmSettings.frx":134D6
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image Image18 
      Height          =   240
      Index           =   0
      Left            =   3960
      Picture         =   "FrmSettings.frx":13A60
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image Image17 
      Height          =   240
      Index           =   3
      Left            =   5040
      Picture         =   "FrmSettings.frx":13FEA
      Top             =   2280
      Width           =   240
   End
   Begin VB.Image Image17 
      Height          =   240
      Index           =   2
      Left            =   4680
      Picture         =   "FrmSettings.frx":14574
      Top             =   2280
      Width           =   240
   End
   Begin VB.Image Image17 
      Height          =   240
      Index           =   1
      Left            =   4320
      Picture         =   "FrmSettings.frx":14AFE
      Top             =   2280
      Width           =   240
   End
   Begin VB.Image Image17 
      Height          =   240
      Index           =   0
      Left            =   3960
      Picture         =   "FrmSettings.frx":15088
      Top             =   2280
      Width           =   240
   End
   Begin VB.Image Image16 
      Height          =   240
      Index           =   3
      Left            =   5040
      Picture         =   "FrmSettings.frx":15612
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image Image16 
      Height          =   240
      Index           =   2
      Left            =   4680
      Picture         =   "FrmSettings.frx":1575C
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image Image16 
      Height          =   240
      Index           =   1
      Left            =   4320
      Picture         =   "FrmSettings.frx":158A6
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image Image16 
      Height          =   240
      Index           =   0
      Left            =   3960
      Picture         =   "FrmSettings.frx":159F0
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image Image15 
      Height          =   240
      Index           =   3
      Left            =   5040
      Picture         =   "FrmSettings.frx":15B3A
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image Image15 
      Height          =   240
      Index           =   2
      Left            =   4680
      Picture         =   "FrmSettings.frx":15C84
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image Image15 
      Height          =   240
      Index           =   1
      Left            =   4320
      Picture         =   "FrmSettings.frx":15DD0
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image Image15 
      Height          =   240
      Index           =   0
      Left            =   3960
      Picture         =   "FrmSettings.frx":15F1C
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image Image14 
      Height          =   240
      Index           =   3
      Left            =   5040
      Picture         =   "FrmSettings.frx":16068
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image14 
      Height          =   240
      Index           =   2
      Left            =   4680
      Picture         =   "FrmSettings.frx":16A6A
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image14 
      Height          =   240
      Index           =   1
      Left            =   4320
      Picture         =   "FrmSettings.frx":1746C
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image14 
      Height          =   240
      Index           =   0
      Left            =   3960
      Picture         =   "FrmSettings.frx":17E6E
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Launch At Startup"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   600
      TabIndex        =   3
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Desktop Form At Startup"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   600
      TabIndex        =   2
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   1200
      Width           =   2505
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Main Form At Startup"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   600
      TabIndex        =   1
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   2520
      Width           =   2250
   End
   Begin VB.Image Image13 
      Height          =   1380
      Left            =   600
      Picture         =   "FrmSettings.frx":18870
      Top             =   2880
      Width           =   2280
   End
   Begin VB.Image Image12 
      Height          =   795
      Left            =   600
      Picture         =   "FrmSettings.frx":1A201
      Top             =   1560
      Width           =   2010
   End
   Begin VB.Image Image11 
      Height          =   180
      Left            =   360
      Picture         =   "FrmSettings.frx":1B258
      Top             =   1200
      Width           =   180
   End
   Begin VB.Image Image10 
      Height          =   180
      Left            =   360
      Picture         =   "FrmSettings.frx":1B613
      Top             =   2520
      Width           =   180
   End
   Begin VB.Image Image3 
      Height          =   180
      Left            =   360
      Picture         =   "FrmSettings.frx":1B9CE
      Top             =   840
      Width           =   180
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Tag             =   "TitleColor"
      Top             =   0
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   180
      Left            =   2160
      Picture         =   "FrmSettings.frx":1BD89
      Top             =   4920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   2160
      Picture         =   "FrmSettings.frx":1C162
      Top             =   4560
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image9 
      Height          =   180
      Left            =   1920
      Picture         =   "FrmSettings.frx":1C51D
      Top             =   4920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image6 
      Height          =   180
      Left            =   1920
      Picture         =   "FrmSettings.frx":1C915
      Top             =   4560
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   120
      Picture         =   "FrmSettings.frx":1CD02
      Top             =   4560
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   1560
      Picture         =   "FrmSettings.frx":1D402
      Top             =   4560
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image7 
      Height          =   330
      Left            =   120
      Picture         =   "FrmSettings.frx":1D8F7
      Top             =   4920
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image8 
      Height          =   315
      Left            =   1560
      Picture         =   "FrmSettings.frx":1E006
      Top             =   4920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image19 
      Height          =   330
      Left            =   3600
      Picture         =   "FrmSettings.frx":1E528
      ToolTipText     =   "Settings"
      Top             =   3850
      Width           =   1425
   End
   Begin VB.Image Image20 
      Height          =   330
      Left            =   5640
      Picture         =   "FrmSettings.frx":1EC37
      ToolTipText     =   "Settings"
      Top             =   3600
      Width           =   1425
   End
   Begin VB.Image Image21 
      Height          =   330
      Left            =   5640
      Picture         =   "FrmSettings.frx":1F346
      ToolTipText     =   "Settings"
      Top             =   3960
      Width           =   1425
   End
   Begin VB.Shape Shape3 
      Height          =   3615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      Height          =   2415
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   1695
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Const ERROR_SUCCESS = 0&
    Const REG_SZ = 1 ' Unicode nul terminated String
    Const REG_DWORD = 4 ' 32-bit number


Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

'For Dragging Borderless Forms...
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Dim sControlSelected As String

Private Sub Form_Load()
Dim CheckMe1 As String
Dim CheckMe2 As String
Dim CheckMe3 As String

CheckMe1 = ReadINI("settings", "launchatstartup", App.Path & "\settings.ini")
CheckMe2 = ReadINI("settings", "showdesktopform", App.Path & "\settings.ini")
CheckMe3 = ReadINI("settings", "showmainform", App.Path & "\settings.ini")
IconToUse = ReadINI("settings", "icon", App.Path & "\settings.ini")

If CheckMe1 = "unchecked" Then
Image3.Picture = Image1.Picture
End If
If CheckMe1 = "checked" Then
Image3.Picture = Image2.Picture
End If

If CheckMe2 = "unchecked" Then
Image11.Picture = Image1.Picture
End If
If CheckMe2 = "checked" Then
Image11.Picture = Image2.Picture
End If

If CheckMe3 = "unchecked" Then
Image10.Picture = Image1.Picture
End If
If CheckMe3 = "checked" Then
Image10.Picture = Image2.Picture
End If

If IconToUse = "icon1" Then
sControlSelected = "icon1"
imgSelected_Click (0)
End If

If IconToUse = "icon2" Then
sControlSelected = "icon2"
imgSelected_Click (1)
End If

If IconToUse = "icon3" Then
sControlSelected = "icon3"
imgSelected_Click (2)
End If

If IconToUse = "icon4" Then
sControlSelected = "icon4"
imgSelected_Click (3)
End If

If IconToUse = "icon5" Then
sControlSelected = "icon5"
imgSelected_Click (4)
End If
Me.Height = 4485
Me.Width = 7500
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Image10.Picture = Image1.Picture Then
    Image10.Picture = Image2.Picture
    Exit Sub
    End If
    If Image10.Picture = Image2.Picture Then
    Image10.Picture = Image1.Picture
    End If
End If
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Image11.Picture = Image1.Picture Then
    Image11.Picture = Image2.Picture
    Exit Sub
    End If
    If Image11.Picture = Image2.Picture Then
    Image11.Picture = Image1.Picture
    End If
End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Image3.Picture = Image1.Picture Then
    Image3.Picture = Image2.Picture
    Exit Sub
    End If
    If Image3.Picture = Image2.Picture Then
    Image3.Picture = Image1.Picture
    End If
End If
End Sub

Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgClose.Picture = Image5.Picture
End If

End Sub
Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgClose.Picture = Image8.Picture
End If

End Sub

Private Sub imgSelected_Click(Index As Integer)
Dim X As Byte

'Clear the radio buttons...
For X = 0 To 4
    imgSelected(X).Picture = Image6.Picture
Next X

'Update the radio buttons...
imgSelected(Index).Picture = Image9.Picture

'Remember the control selected...
Select Case Index
    Case 0
        sControlSelected = "icon1"
    Case 1
        sControlSelected = "icon2"
    Case 2
        sControlSelected = "icon3"
    Case 3
        sControlSelected = "icon4"
    Case 4
        sControlSelected = "icon5"
End Select
End Sub

Private Sub Label15_Click()
FrmColors.Show
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
    DoEvents
End If
End Sub
Public Sub DragForm(Frm As Form)

On Local Error Resume Next

'Move the borderless form...
Call ReleaseCapture
Call SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub

Private Sub Label5_Click()

If Image3.Picture = Image2.Picture Then
Call AddToRun("CS Bandwidth Monitor", App.Path & "\" & App.EXEName & ".exe")
End If
If Image3.Picture = Image1.Picture Then
Call RemoveFromRun("CS Bandwidth Monitor")
End If

Call SaveChanges
IconToUse = sControlSelected

End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image20.Picture = Image4.Picture
End If
End Sub
Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Image20.Picture = Image7.Picture
End If
End Sub

Private Sub Label7_Click()
Unload Me
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image21.Picture = Image4.Picture
End If
End Sub
Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Image21.Picture = Image7.Picture
End If
End Sub
Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image19.Picture = Image4.Picture
End If
End Sub
Private Sub Label15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Image19.Picture = Image7.Picture
End If
End Sub
Sub SaveChanges()

'On Local Error Resume Next

'Save the color settings to the skin scheme ini file...
    If Image3.Picture = Image1.Picture Then
    Call WriteINI("settings", "launchatstartup", "unchecked", App.Path & "\settings.ini")
    End If
    If Image3.Picture = Image2.Picture Then
    Call WriteINI("settings", "launchatstartup", "checked", App.Path & "\settings.ini")
    End If
    
    If Image11.Picture = Image1.Picture Then
    Call WriteINI("settings", "showdesktopform", "unchecked", App.Path & "\settings.ini")
    End If
    If Image11.Picture = Image2.Picture Then
    Call WriteINI("settings", "showdesktopform", "checked", App.Path & "\settings.ini")
    End If

    If Image10.Picture = Image1.Picture Then
    Call WriteINI("settings", "showmainform", "unchecked", App.Path & "\settings.ini")
    End If
    If Image10.Picture = Image2.Picture Then
    Call WriteINI("settings", "showmainform", "checked", App.Path & "\settings.ini")
    End If
    
Call WriteINI("settings", "icon", sControlSelected, App.Path & "\settings.ini")

End Sub
Public Sub AddToRun(ProgramName As String, FileToRun As String)
    'Add a program to the 'Run at Startup' r
    '     egistry keys
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, FileToRun)
End Sub


Public Sub RemoveFromRun(ProgramName As String)
    'Remove a program from the 'Run at Start
    '     up' registry keys
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName)
End Sub
Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strData As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    r = RegCloseKey(keyhand)
End Sub


Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function


Public Function DeleteKey(ByVal hKey As HKeyTypes, ByVal strPath As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry")
    '
    Dim keyhand As Long
    r = RegDeleteKey(hKey, strPath)
End Function
