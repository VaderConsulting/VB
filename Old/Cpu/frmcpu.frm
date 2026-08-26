VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   9
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   9
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   8
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   7
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   240
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   6
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   5
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   480
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   4
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   600
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   3
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   720
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   2
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   840
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   1
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   960
      Width           =   135
   End
   Begin VB.PictureBox picCPU 
      Height          =   135
      Index           =   0
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   1080
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type LARGE_INTEGER
                lowpart As Long
                highpart As Long
End Type
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Const REG_DWORD = 4 ' 32-bit number
Private Const HKEY_DYN_DATA = &H80000006
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Form_Activate()
    Height = 1305
    Top = Screen.Height - Height
    Left = Screen.Width - Width
End Sub
Private Sub Form_DblClick()
    'If the mouse is double clicked then end the
    'program.
    End
End Sub
Private Sub Form_Load()
    'Positions the form at the top right corner of
    'the screen regardless of monitor size.
    Form1.Top = 1
    Form1.Left = Screen.Width - Form1.Width

    Call InitCPU
    Call OnTop
End Sub

Private Sub SSPanel1_DblClick(Index As Integer)
    Call Form_DblClick
End Sub

Private Sub picCPU_DblClick(Index As Integer)
    Call Form_DblClick
End Sub
Private Sub Timer1_Timer()
    Dim lData As Long, lType As Long, lSize As Long
    Dim hKey As Long

    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)

    'If there's a problem accessing the registry
    If Qry <> 0 Then
         MsgBox "Can't Open Statistics Key"
         End
    End If

    lType = REG_DWORD
    lSize = 4

    'Querying the registry for CPUUsage
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)

    'clears the SSPanel boxes
    Do Until picCPU(x).BackColor = &HC0C0C0
       picCPU(x).BackColor = &HC0C0C0
       x = x + 1
       If x >= 10 Then Exit Do
    Loop

    'statbar is the variable that holds the CPU
    'usage divided by 10.
    '(ex. if 79% of the CPU is being used then
    ' statbar will hold the int(7.9) = 8)
    statbar = Int(lData / 10)
    If statbar >= 1 Then statbar = statbar - 1

    'used to fill the SSPanel with the color green
    'beginning with 0 and ending with the value of
    'statbar.
    For fillall = 0 To statbar
        picCPU(fillall).BackColor = &HFF00&
    Next fillall

    If Int(lData / 10) = 0 Then picCPU(0).BackColor = &HC0C0C0
    'Print lData
    'Label2.Caption = lData & "%"
    Qry = RegCloseKey(hKey)

End Sub
Private Sub InitCPU()
    Dim lData As Long, lType As Long, lSize As Long
    Dim hKey As Long
    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
    If Qry <> 0 Then
          MsgBox "Can't Open Statistics Key"
          End
    End If

    lType = REG_DWORD
    lSize = 4
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
    Qry = RegCloseKey(hKey)
End Sub
Private Sub OnTop()
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    If SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS) = True Then
        success% = SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

