VERSION 5.00
Begin VB.Form FrmDesktop 
   BorderStyle     =   0  'None
   Caption         =   "CS Bandwidth Monitor"
   ClientHeight    =   1440
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   3900
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "form1"
   ScaleHeight     =   1440
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   195
      Left            =   3720
      TabIndex        =   14
      ToolTipText     =   "Reset Background Image To Match Current Desktop"
      Top             =   0
      Width           =   135
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3480
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3480
      Top             =   960
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   0
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   13
      Top             =   1680
      Width           =   3900
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000C0C0&
      X1              =   3840
      X2              =   3840
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 Kb"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 Kb"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Upload Speed"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Download Speed"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblRecv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblSent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sent"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000C0C0&
      X1              =   0
      X2              =   2040
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000C0C0&
      X1              =   0
      X2              =   2280
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000C0C0&
      X1              =   2040
      X2              =   2040
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000C0C0&
      X1              =   2040
      X2              =   3840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000C0C0&
      X1              =   0
      X2              =   3840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0000C0C0&
      X1              =   0
      X2              =   3840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000C0C0&
      X1              =   840
      X2              =   840
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kb"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kb"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "FrmDesktop"
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

Private m_objIpHelper As CIpHelper
Private TransferRate                    As Single
Private TransferRate2                   As Single
Public Sub DragForm(Frm As Form)

On Local Error Resume Next

'Move the borderless form...
Call ReleaseCapture
Call SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub
Public Sub UpdatePicture()
Dim DeskhWnd As Long
Dim DeskhDC As Long
Dim XX As Long
Dim YY As Long

Set m_objIpHelper = New CIpHelper

YY = Me.ScaleY(Me.Left, vbTwips, vbPixels)
XX = Me.ScaleX(Me.Top, vbTwips, vbPixels)
Me.WindowState = 1
DoEvents
    DeskhWnd = GetDesktopWindow
    DeskhDC = GetDC(DeskhWnd)
    
    Call BitBlt(picTemp.hdc, 0, 0, Me.Width, Me.Height, _
                DeskhDC, YY, XX, SRCCOPY)
    
    Call ReleaseDC(DeskhWnd, DeskhDC)
    
    Me.Picture = Me.picTemp.Image
    
    Me.WindowState = 0

    Dim xFile As String
    xFile = App.Path & "\background.bmp"

    SavePicture picTemp.Image, xFile
    
End Sub

Private Sub Command1_Click()
Call UpdatePicture
End Sub

Private Sub Form_Load()
Set m_objIpHelper = New CIpHelper
Me.Top = Val(ReadINI("formposition", "desktoptop", QuickRef.UserINIFileName))
Me.Left = Val(ReadINI("formposition", "desktopleft", QuickRef.UserINIFileName))
DoEvents

    Me.Picture = LoadPicture(App.Path & "\background.bmp")
    
Call LoadColors
Call SetColors(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Me.Top <= 0 Then
Me.Top = 0
End If

If Me.Top >= Screen.Height - Me.Height - 500 Then
Me.Top = Screen.Height - Me.Height - 500
End If

If Me.Left <= 0 Then
Me.Left = 0
End If

If Me.Left >= Screen.Width - Me.Width Then
Me.Left = Screen.Width - Me.Width
End If

If Button = vbLeftButton Then
    Call DragForm(Me)
    DoEvents
Call WriteINI("formposition", "desktoptop", Me.Top, App.Path & "\settings.ini")
Call WriteINI("formposition", "desktopleft", Me.Left, App.Path & "\settings.ini")
End If
End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
Call UpdateInterfaceInfo
End Sub
Private Sub UpdateInterfaceInfo()

'On Error Resume Next
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)
Select Case objInterface.InterfaceType
Case MIB_IF_TYPE_ETHERNET: lblType.Caption = "Ethernet"
Case MIB_IF_TYPE_FDDI: lblType.Caption = "FDDI"
Case MIB_IF_TYPE_LOOPBACK: lblType.Caption = "Loopback"
Case MIB_IF_TYPE_OTHER: lblType.Caption = "Other"
Case MIB_IF_TYPE_PPP: lblType.Caption = "PPP"
Case MIB_IF_TYPE_SLIP: lblType.Caption = "SLIP"
Case MIB_IF_TYPE_TOKENRING: lblType.Caption = "TokenRing"
End Select
lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived / 1024, "###,###,###,###,##0"))
lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent / 1024, "###,###,###,###,##0"))
Set st_objInterface = objInterface

End Sub

Private Sub Timer2_Timer()

On Error Resume Next
DoEvents
Dim XX As Long
Dim YY As Long
Dim XXX As Long
Dim YYY As Long
YYY = Label6.Caption
YY = Label5.Caption
DoEvents
XX = Me.lblRecv.Caption - YY
XXX = Me.lblSent.Caption - YYY
DoEvents
TransferRate = Format(Int(XX), "00.00")
DoEvents
TransferRate2 = Format(Int(XXX), "00.00")
DoEvents

                Label10.Caption = TransferRate2 & " KB"
                DoEvents

                Label9.Caption = TransferRate & " KB"
                DoEvents
    DoEvents
    Label5.Caption = Me.lblRecv.Caption
    Label6.Caption = Me.lblSent.Caption
    DoEvents

End Sub
