VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "CS Bandwidth Monitor"
   ClientHeight    =   4155
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":0A02
   ScaleHeight     =   4155
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic1 
      Height          =   375
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   3480
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":7222
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":7C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":8646
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9058
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   2760
      Top             =   2880
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3240
      Top             =   2880
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4440
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9E7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   3240
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A134
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A28E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A3E8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   3840
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A542
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B076
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B610
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   4440
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BBAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":C144
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":C6DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":CC78
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image btnBottomDown 
      Height          =   165
      Index           =   0
      Left            =   2280
      Picture         =   "FrmMain.frx":D212
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image btnBottomDown 
      Height          =   165
      Index           =   2
      Left            =   2280
      Picture         =   "FrmMain.frx":D570
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image btnBottomDown 
      Height          =   165
      Index           =   1
      Left            =   240
      Picture         =   "FrmMain.frx":D8CE
      Top             =   2400
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   1920
      Picture         =   "FrmMain.frx":DC2C
      Top             =   3240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image8 
      Height          =   315
      Left            =   1560
      Picture         =   "FrmMain.frx":E145
      Top             =   3240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image7 
      Height          =   330
      Left            =   120
      Picture         =   "FrmMain.frx":E667
      Top             =   3240
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   1920
      Picture         =   "FrmMain.frx":ED76
      Top             =   2880
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   1560
      Picture         =   "FrmMain.frx":F26A
      Top             =   2880
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   120
      Picture         =   "FrmMain.frx":F75F
      Top             =   2880
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2400
      TabIndex        =   15
      Tag             =   "TitleColor"
      Top             =   2325
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   2400
      Picture         =   "FrmMain.frx":FE5F
      ToolTipText     =   "Settings"
      Top             =   2280
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4080
      Picture         =   "FrmMain.frx":1056E
      Top             =   2400
      Width           =   240
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
      Left            =   960
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
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
      Left            =   1320
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   600
      Y2              =   1320
   End
   Begin VB.Line Line6 
      X1              =   4200
      X2              =   4200
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "KB"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "KB"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   840
      Width           =   255
   End
   Begin VB.Line Line9 
      X1              =   1200
      X2              =   1200
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Line Line8 
      X1              =   360
      X2              =   4200
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line7 
      X1              =   360
      X2              =   4200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line5 
      X1              =   2400
      X2              =   4200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line4 
      X1              =   2400
      X2              =   2400
      Y1              =   600
      Y2              =   840
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   2640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   2400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sent"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   405
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   1980
   End
   Begin VB.Label lblSent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblRecv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Download Speed"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Upload Speed"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 KB"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 KB"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CS Bandwidth Monitor v1.0.0"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Tag             =   "TitleColor"
      Top             =   120
      Width           =   2385
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   4200
      Picture         =   "FrmMain.frx":10F70
      Tag             =   "Close"
      ToolTipText     =   "Close"
      Top             =   45
      Width           =   315
   End
   Begin VB.Image imgMinimize 
      Height          =   300
      Left            =   3840
      Picture         =   "FrmMain.frx":11492
      Tag             =   "Minimize"
      ToolTipText     =   "Minimize"
      Top             =   45
      Width           =   315
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   720
      TabIndex        =   14
      Tag             =   "TitleColor"
      Top             =   2325
      Width           =   1485
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   720
      Picture         =   "FrmMain.frx":119AB
      ToolTipText     =   "Settings"
      Top             =   2280
      Width           =   1425
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For Dragging Borderless Forms...
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long


Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Type NOTIFYICONDATA
    cbSize As Long
    mhWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Dim TheForm As NOTIFYICONDATA

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private m_objIpHelper As CIpHelper
Private TransferRate                    As Single
Private TransferRate2                   As Single


Private Sub Form_Unload(Cancel As Integer)
    'If UnloadMode = 0 Then
    Cancel = True
    'End If
    Me.Hide
    Unload FrmStats
End Sub

Private Sub Image2_Click()
    FrmSettings.Show
End Sub

Private Sub Image3_Click()
    frmAbout.Show
End Sub

Private Sub Label11_Click()
    frmAbout.Show
End Sub

Private Sub Label12_Click()
    FrmSettings.Show
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Rec As Boolean, Msg As Long
    
    Msg = X / Screen.TwipsPerPixelX
    If Rec = False Then
        Rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                FrmMain.Show
            Case WM_LBUTTONDOWN:

            Case WM_LBUTTONUP:
                
            Case WM_RBUTTONDBLCLK:
                'PopupMenu mnufile
            Case WM_RBUTTONDOWN:
                
            Case WM_RBUTTONUP:
                PopupMenu FrmMenu.MenuFile
        End Select
        Rec = False
    End If
End Sub

Private Sub Timer1_Timer()
    
    On Error Resume Next
    Call UpdateInterfaceInfo
    
End Sub

Private Sub UpdateInterfaceInfo()
    
    On Error Resume Next
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
    '---------------
    blnIsRecv = (m_objIpHelper.BytesReceived / 1024 > lngBytesRecv / 1024)
    blnIsSent = (m_objIpHelper.BytesSent / 1024 > lngBytesSent / 1024)
    
    If IconToUse = "icon1" Then
        If blnIsRecv And blnIsSent Then
            Pic1.Picture = ImageList1.ListImages(4).Picture
            Me.Icon = ImageList1.ListImages(4).Picture
            Image1.Picture = ImageList1.ListImages(4).Picture
        ElseIf (Not blnIsRecv) And blnIsSent Then
            Pic1.Picture = ImageList1.ListImages(2).Picture
            Me.Icon = ImageList1.ListImages(2).Picture
            Image1.Picture = ImageList1.ListImages(2).Picture
        ElseIf blnIsRecv And (Not blnIsSent) Then
            Pic1.Picture = ImageList1.ListImages(3).Picture
            Me.Icon = ImageList1.ListImages(3).Picture
            Image1.Picture = ImageList1.ListImages(3).Picture
        ElseIf Not (blnIsRecv And blnIsSent) Then
            Pic1.Picture = ImageList1.ListImages(1).Picture
            Me.Icon = ImageList1.ListImages(1).Picture
            Image1.Picture = ImageList1.ListImages(1).Picture
        End If
    End If
    
    If IconToUse = "icon2" Then
        If blnIsRecv And blnIsSent Then
            Pic1.Picture = ImageList2.ListImages(4).Picture
            Me.Icon = ImageList2.ListImages(4).Picture
            Image1.Picture = ImageList2.ListImages(4).Picture
        ElseIf (Not blnIsRecv) And blnIsSent Then
            Pic1.Picture = ImageList2.ListImages(2).Picture
            Me.Icon = ImageList2.ListImages(2).Picture
            Image1.Picture = ImageList2.ListImages(2).Picture
        ElseIf blnIsRecv And (Not blnIsSent) Then
            Pic1.Picture = ImageList2.ListImages(3).Picture
            Me.Icon = ImageList2.ListImages(3).Picture
            Image1.Picture = ImageList2.ListImages(3).Picture
        ElseIf Not (blnIsRecv And blnIsSent) Then
            Pic1.Picture = ImageList2.ListImages(1).Picture
            Me.Icon = ImageList2.ListImages(1).Picture
            Image1.Picture = ImageList2.ListImages(1).Picture
        End If
    End If
    
    If IconToUse = "icon3" Then
        If blnIsRecv And blnIsSent Then
            Pic1.Picture = ImageList3.ListImages(4).Picture
            Me.Icon = ImageList3.ListImages(4).Picture
            Image1.Picture = ImageList3.ListImages(4).Picture
        ElseIf (Not blnIsRecv) And blnIsSent Then
            Pic1.Picture = ImageList3.ListImages(2).Picture
            Me.Icon = ImageList3.ListImages(2).Picture
            Image1.Picture = ImageList3.ListImages(2).Picture
        ElseIf blnIsRecv And (Not blnIsSent) Then
            Pic1.Picture = ImageList3.ListImages(3).Picture
            Me.Icon = ImageList3.ListImages(3).Picture
            Image1.Picture = ImageList3.ListImages(3).Picture
        ElseIf Not (blnIsRecv And blnIsSent) Then
            Pic1.Picture = ImageList3.ListImages(1).Picture
            Me.Icon = ImageList3.ListImages(1).Picture
            Image1.Picture = ImageList3.ListImages(1).Picture
        End If
    End If
    
    If IconToUse = "icon4" Then
        If blnIsRecv And blnIsSent Then
            Pic1.Picture = ImageList4.ListImages(4).Picture
            Me.Icon = ImageList4.ListImages(4).Picture
            Image1.Picture = ImageList4.ListImages(4).Picture
        ElseIf (Not blnIsRecv) And blnIsSent Then
            Pic1.Picture = ImageList4.ListImages(2).Picture
            Me.Icon = ImageList4.ListImages(2).Picture
            Image1.Picture = ImageList4.ListImages(2).Picture
        ElseIf blnIsRecv And (Not blnIsSent) Then
            Pic1.Picture = ImageList4.ListImages(3).Picture
            Me.Icon = ImageList4.ListImages(3).Picture
            Image1.Picture = ImageList4.ListImages(3).Picture
        ElseIf Not (blnIsRecv And blnIsSent) Then
            Pic1.Picture = ImageList4.ListImages(1).Picture
            Me.Icon = ImageList4.ListImages(1).Picture
            Image1.Picture = ImageList4.ListImages(1).Picture
        End If
    End If
    
    If IconToUse = "icon5" Then
        If blnIsRecv And blnIsSent Then
            Pic1.Picture = ImageList5.ListImages(4).Picture
            Me.Icon = ImageList5.ListImages(4).Picture
            Image1.Picture = ImageList5.ListImages(4).Picture
        ElseIf (Not blnIsRecv) And blnIsSent Then
            Pic1.Picture = ImageList5.ListImages(2).Picture
            Me.Icon = ImageList5.ListImages(2).Picture
            Image1.Picture = ImageList5.ListImages(2).Picture
        ElseIf blnIsRecv And (Not blnIsSent) Then
            Pic1.Picture = ImageList5.ListImages(3).Picture
            Me.Icon = ImageList5.ListImages(3).Picture
            Image1.Picture = ImageList5.ListImages(3).Picture
        ElseIf Not (blnIsRecv And blnIsSent) Then
            Pic1.Picture = ImageList5.ListImages(1).Picture
            Me.Icon = ImageList5.ListImages(1).Picture
            Image1.Picture = ImageList5.ListImages(1).Picture
        End If
    End If
    
    ModifyIcon
    lngBytesRecv = m_objIpHelper.BytesReceived
    lngBytesSent = m_objIpHelper.BytesSent
    DoEvents

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

Public Sub DragForm(Frm As Form)

    On Local Error Resume Next
    
    'Move the borderless form...
    Call ReleaseCapture
    Call SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub

Private Sub btnBottomDown_Click(Index As Integer)
    FrmStats.WindowState = 0
    ' only do this if we are slid out
    If bBottomOut = True Then
        ' reset the target offest ...
        iRelBottomTrayOffset = FrmMain.Top
        ' and reel her in
        Do While FrmStats.Top > iRelBottomTrayOffset
            ' in two pixels
            FrmStats.Top = FrmStats.Top - 30
            ' and make sure the main form stays on top
            FrmMain.ZOrder
            DoEvents
        Loop
        ' now hide the tray and set the bBottomOut flag false to allow sliding down
        FrmStats.Hide
        bBottomOut = False
        Exit Sub
    End If
    DoEvents
    ' don't do anything if it's already slid down
    If bBottomOut = True Then Exit Sub
    ' postition the tray ready to slide down
    FrmStats.Left = FrmMain.Left ' + 150
    FrmStats.Top = FrmMain.Top
    FrmStats.Show
    Me.SetFocus
    DoEvents
    iRelBottomTrayOffset = FrmMain.Top + FrmMain.Height ' - 75
    'Do Until FrmStats.Height >= 4485
    'FrmStats.Height = FrmStats.Height + 15
    'Loop
    Do While FrmStats.Top < iRelBottomTrayOffset
        'down 1 more pixel
        FrmStats.Top = FrmStats.Top + 15
        ' make sure the main form stays on top
        FrmMain.ZOrder
        DoEvents
    Loop
    FrmStats.Top = FrmMain.Top + FrmMain.Height
    bBottomOut = True
     
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
        End
    End If
    DoEvents
    
    Set m_objIpHelper = New CIpHelper
    
    bBottomOut = False
    
    Me.Height = 2745
    Me.Width = 4560
    
    IconToUse = ReadINI("settings", "icon", App.Path & "\settings.ini")
    
    Me.Top = Val(ReadINI("formposition", "maintop", QuickRef.UserINIFileName))
    Me.Left = Val(ReadINI("formposition", "mainleft", QuickRef.UserINIFileName))
    
    Me.Caption = "Bandwidth v" & App.Major & "." & App.Minor & "." & App.Revision & " Beta 1"
    lblCaption.Caption = "Bandwidth v" & App.Major & "." & App.Minor & "." & App.Revision & " Beta 1"
    
    Call LoadColors
    Call SetColors(Me)
    
    SysTray
    DoEvents
    If ReadINI("settings", "showmainform", App.Path & "\settings.ini") = "unchecked" Then
        Timer3.Enabled = True
    End If
    
    If ReadINI("settings", "showdesktopform", App.Path & "\settings.ini") = "checked" Then
        FrmDesktop.Show
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    
    btnBottomDown(1).Picture = btnBottomDown(0).Picture
    'Move the form if the user is pressing and holding the mouse button...
    If Button = vbLeftButton Then
        Call DragForm(Me)
        If bBottomOut = True Then
            FrmStats.Left = FrmMain.Left
            FrmStats.Top = FrmMain.Top + FrmMain.Height
        End If
    Call WriteINI("formposition", "maintop", Me.Top, App.Path & "\settings.ini")
    Call WriteINI("formposition", "mainleft", Me.Left, App.Path & "\settings.ini")
    End If


End Sub

Private Sub Form_Resize()
    If FrmStats.WindowState = vbMinimized Then
        FrmStats.WindowState = 0
    End If
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

Private Sub imgMinimize_Click()
    On Error Resume Next
    Me.WindowState = vbMinimized
    FrmStats.WindowState = vbMinimized
End Sub

Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        imgMinimize.Picture = Image6.Picture
    End If
    
End Sub

Private Sub imgMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        imgMinimize.Picture = Image9.Picture
    End If

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
Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Image2.Picture = Image4.Picture
    End If
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Image2.Picture = Image7.Picture
    End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Image2.Picture = Image4.Picture
    End If
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Image2.Picture = Image7.Picture
    End If
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    btnBottomDown(1).Picture = btnBottomDown(0).Picture
    'Move the form if the user is pressing and holding the mouse button...
    If Button = vbLeftButton Then
        Call DragForm(Me)
        If bBottomOut = True Then
            FrmStats.Left = FrmMain.Left
            FrmStats.Top = FrmMain.Top + FrmMain.Height
        End If
    End If

End Sub

Private Sub btnBottomDown_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' hilite the little green down arrow and unhilite the up arrow by swapping pics
    btnBottomDown(1).Picture = btnBottomDown(2).Picture
    'btnBottomUp(0) = btnBottomUp(1)
End Sub

Public Function SysTray()
    TheForm.cbSize = Len(TheForm)
    
    TheForm.mhWnd = Pic1.hwnd
    TheForm.hIcon = Pic1.Picture
    TheForm.uId = 1&
    
    TheForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    
    TheForm.ucallbackMessage = WM_MOUSEMOVE
    
    TheForm.szTip = "CS Bandwidth Monitor v" & App.Major & "." & App.Minor & "." & App.Revision
    
    Shell_NotifyIcon NIM_ADD, TheForm
End Function

Function ModifyIcon()
    TheForm.cbSize = Len(TheForm)
    
    TheForm.mhWnd = Pic1.hwnd
    TheForm.hIcon = Pic1.Picture
    TheForm.uId = 1&
    
    TheForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    
    TheForm.ucallbackMessage = WM_MOUSEMOVE
    
    Shell_NotifyIcon NIM_MODIFY, TheForm
End Function

Public Sub CleanUpSystray()
    Shell_NotifyIcon NIM_DELETE, TheForm
End Sub

Private Sub Timer3_Timer()
    Me.Hide
    Timer3.Enabled = False
End Sub
