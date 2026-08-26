VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send SMTP email without mswinsck.ocx - by Joseph Ninan <josephninan@crosswinds.net>"
   ClientHeight    =   6405
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Form1.frx":0442
   ScaleHeight     =   6405
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCLS 
      Caption         =   "Clear Log"
      Height          =   495
      Left            =   7800
      TabIndex        =   26
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load todays Log file"
      Height          =   495
      Left            =   6360
      TabIndex        =   24
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveLog 
      Caption         =   "Save Log"
      Height          =   495
      Left            =   5160
      TabIndex        =   18
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtm_subject 
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Text            =   "API's testing"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtm_ReplyTo 
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Text            =   "someone@jofu.8m.com"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtname_rcpt 
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Text            =   "joseph"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtm_rcpt 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Text            =   "jofu@crosswinds.net"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtname_from 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Text            =   "Someone"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtm_from 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Text            =   "anyone@jofu.8m.com"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "Send Mail"
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtm_data 
      Height          =   915
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form1.frx":074C
      Top             =   2520
      Width           =   8775
   End
   Begin VB.TextBox txthost 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtport 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "25"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtStatus 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4560
      Width           =   8775
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   4320
      Width           =   7095
   End
   Begin VB.Label Label12 
      Caption         =   "Rate this program. Search for Joseph Ninan at Planetsourcecode.Other codes like Permutations - Get all possible passwords"
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   5520
      TabIndex        =   23
      ToolTipText     =   "http://www.planetsourcecode.com"
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label11 
      Caption         =   "Visit my site"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7080
      MouseIcon       =   "Form1.frx":0890
      TabIndex        =   22
      ToolTipText     =   "http://www.jofu.8m.com"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Rate my code"
      Height          =   255
      Left            =   5520
      TabIndex        =   21
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Will be saved under file c:\smtplog*.*"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label8 
      Caption         =   $"Form1.frx":0B9A
      Height          =   855
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   8535
   End
   Begin VB.Label Label7 
      Caption         =   "Debug Info......"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Subject"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Reply to :"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Rcpt"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "&Mail Server:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu menFile 
      Caption         =   "&File"
      Begin VB.Menu miExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This was programmed by  Joseph Ninan
' email - josephninan@crosswinds.net
' S4-Computer Science and engineering
' SCT College of engineering
' Trivandrum, Kerala, India
' Phone - 0091-471-594477
' www.jofu.8m.com

Private Sub cmdCLS_Click()
    Me.txtStatus.Text = ""
End Sub

Private Sub cmdLoad_Click()
    Dim afile, aLog, bLog As Variant
    Call cmdSaveLog_Click
    On Error Resume Next
    afile = "c:\smtplog" & Format(Date, "dddd_mmm_d_yyyy") & ".txt"
    Open afile For Input As #5
    While Not EOF(5)
        Line Input #5, aLog
        bLog = bLog & vbCrLf & aLog
    Wend
    Me.txtStatus.Text = bLog
    Close

End Sub

Private Sub cmdSaveLog_Click()
    Dim afile As Variant
    On Error Resume Next
    afile = "c:\smtplog" & Format(Date, "dddd_mmm_d_yyyy") & ".txt"
    Open afile For Append As #5
        Print #5, txtStatus.Text
        txtStatus.Text = ""
    Close
End Sub

Private Sub cmdSendMail_Click()
Dim aRes As Integer
    aRes = smtp(txthost, txtport, txtm_from, txtm_rcpt, txtname_from, txtname_rcpt, txtm_ReplyTo, txtm_subject, txtm_data)
    MsgBox ("The result of smtp is " & aRes)
End Sub
Private Sub Form_Load()
    Dim MyString, MyHost As String, MyReg As New cReadWriteEasyReg, i As Integer
    'ok, we have to start winsock, DUH!
    Call StartWinsock("")
    'lets subclassing the handle
    'for the connection we are going to make
    Call Hook(Form1.hWnd)
    Me.Label13.Caption = "c:\smtplog" & Format(Date, "dddd_mmm_d_yyyy") & ".txt"

    '
    'This function will return a specific value from the registry

    If Not MyReg.OpenRegistry(HKEY_CURRENT_USER, "Software\Microsoft\Internet Account Manager\Accounts\00000001") Then
        MsgBox "Couldn't open the registry.....Proceeding with default values.... Change them as you wish...."
        Exit Sub
    End If
    MyString = MyReg.GetValue("SMTP Email Address")
    MyHost = MyReg.GetValue("SMTP Server")
    Me.txtm_from.Text = MyString
    Me.txtname_from = Me.txtname_from & "-" & MyString
    Me.txtm_ReplyTo = MyString
    Me.txthost.Text = MyHost
    MyReg.CloseRegistry
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePointer = 1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'lets close the connection
    Call closesocket(mysock)
    'lets unhook the hwnd so we dont
    'get an error
    Call UnHook(Form1.hWnd)
End Sub
Private Sub Label11_Click()
    Dim a
    On Error Resume Next
    a = Shell("explorer http://www.jofu.8m.com", vbNormalFocus)
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePointer = 99
End Sub
Private Sub Label12_Click()
    Dim a
    On Error Resume Next
    a = Shell("explorer http://www.planetsourcecode.com", vbNormalFocus)
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePointer = 99
End Sub
Private Sub miExit_Click()
    'lets close the connection
    Call closesocket(mysock)
    'lets unhook the hwnd so we dont
    'get an error
    Call UnHook(Form1.hWnd)
    End
End Sub

Private Sub txtm_data_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePointer = 1

End Sub

Private Sub txtStatus_Change()
    'keep the txtbox at the very bottom at all times
    txtStatus.SelStart = Len(txtStatus)
End Sub
