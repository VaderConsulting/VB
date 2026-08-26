VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   2745
   ClientLeft      =   2295
   ClientTop       =   1500
   ClientWidth     =   4605
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0A02
   ScaleHeight     =   1894.648
   ScaleMode       =   0  'User
   ScaleWidth      =   4324.334
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H0097948F&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   240
      Left            =   600
      Picture         =   "frmAbout.frx":7222
      ScaleHeight     =   168.56
      ScaleMode       =   0  'User
      ScaleWidth      =   168.56
      TabIndex        =   0
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.croftssoftware.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   480
      MouseIcon       =   "frmAbout.frx":7C24
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   1560
      Picture         =   "frmAbout.frx":7F2E
      Top             =   1560
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image8 
      Height          =   315
      Left            =   1560
      Picture         =   "frmAbout.frx":8423
      Top             =   1920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   4200
      Picture         =   "frmAbout.frx":8945
      Tag             =   "Close"
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   315
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1950
      TabIndex        =   5
      Tag             =   "TitleColor"
      Top             =   2010
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   0
      Picture         =   "frmAbout.frx":8E67
      Top             =   1560
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image7 
      Height          =   330
      Left            =   0
      Picture         =   "frmAbout.frx":9567
      Top             =   1920
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About CS Bandwidth Monitor"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Tag             =   "TitleColor"
      Top             =   120
      Width           =   2445
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CS Bandwidth Monitor By Shane M. Croft Of Crofts Software."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   360
      TabIndex        =   1
      Top             =   1410
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   360
      TabIndex        =   2
      Top             =   525
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   3
      Top             =   1065
      Width           =   3885
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   1590
      Picture         =   "frmAbout.frx":9C76
      ToolTipText     =   "Settings"
      Top             =   1965
      Width           =   1425
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For Dragging Borderless Forms...
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblCaption.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

Private Sub Image3_Click()
  Unload Me
End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image3.Picture = Image4.Picture
End If
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image3.Picture = Image7.Picture
End If
End Sub

Private Sub Label1_Click()
On Error Resume Next
Call ShellExecute(hwnd, "Open", "http://www.croftssoftware.com", "", App.Path, 1)

End Sub

Private Sub Label11_Click()
Unload Me
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image3.Picture = Image4.Picture
End If
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image3.Picture = Image7.Picture
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

