VERSION 5.00
Object = "{CAF15339-E0F8-11D4-8999-204C4F4F5020}#1.0#0"; "akled.ocx"
Object = "{0468C941-83E2-11D3-BE51-00C0DFC2E32C}#1.0#0"; "PingX.Ocx"
Object = "{B186D399-9B91-41CB-9242-E12B96CA99E1}#1.0#0"; "DRPing.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Real Time Monitor Server"
   ClientHeight    =   6525
   ClientLeft      =   3135
   ClientTop       =   3360
   ClientWidth     =   10695
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   Begin Ping.drPing Ping 
      Left            =   720
      Top             =   2280
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin MabryPingXCtl.PingX pngRemote 
      Left            =   240
      Top             =   3240
      RemoteAddress   =   ""
      RemotePort      =   0
      Timeout         =   5
      DebugMode       =   0   'False
      BlockingMode    =   0
      Blocking        =   -1  'True
      RequestLength   =   64
      RequestData     =   "ECHO"
      LibraryName     =   "WSOCK32.DLL"
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   120
      Top             =   2280
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   10695
      TabIndex        =   1
      Top             =   0
      Width           =   10695
      Begin VB.TextBox txtAdminPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   13
         Text            =   "25"
         Top             =   135
         Width           =   555
      End
      Begin VB.TextBox txtMonitored 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "0.0.0.0"
         Top             =   480
         Width           =   1425
      End
      Begin VB.ListBox listStatus 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   6000
      End
      Begin VB.TextBox textLocalPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3870
         TabIndex        =   3
         Text            =   "21"
         Top             =   135
         Width           =   555
      End
      Begin VB.TextBox textLocalAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "0.0.0.0"
         Top             =   120
         Width           =   1425
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   7650
         Top             =   780
      End
      Begin akLEDProject.akLED ledRx 
         Height          =   255
         Left            =   6600
         Top             =   165
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ScaleMode       =   3
         LED_Colour      =   11467422
      End
      Begin akLEDProject.akLED ledTx 
         Height          =   255
         Left            =   7200
         Top             =   160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ScaleMode       =   3
         LED_Colour      =   11467422
      End
      Begin akLEDProject.akLED ledStatus 
         Height          =   255
         Left            =   3000
         Top             =   510
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ScaleMode       =   3
         LED_Colour      =   8421504
      End
      Begin VB.Label lblMonitored 
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   520
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Admin Port:"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   4560
         TabIndex        =   12
         Top             =   150
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "Monitored Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "Status:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   435
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Local Port:"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   2940
         TabIndex        =   8
         Top             =   150
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Local Address:"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   165
         TabIndex        =   7
         Top             =   150
         Width           =   1155
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   0
         X2              =   20000
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   8415
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Label Label3 
         Caption         =   "Rx"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   6360
         TabIndex        =   6
         Top             =   210
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "Tx"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   6960
         TabIndex        =   5
         Top             =   210
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10695
      TabIndex        =   0
      Top             =   1875
      Width           =   10695
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu WindowMenu 
      Caption         =   "&Window"
      Visible         =   0   'False
      WindowList      =   -1  'True
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu HelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents SocketX As SocketXObj
Attribute SocketX.VB_VarHelpID = -1
Dim WithEvents AdminSocket As SocketXObj
Attribute AdminSocket.VB_VarHelpID = -1
Public boolTx As Boolean, boolRx As Boolean
Public MonitoredIP As String

Private Sub FileExit_Click()
    Unload Me
    Set SocketX = Nothing
    Set AdminSocket = Nothing
    End
End Sub

Private Sub SocketX_Accept(ByVal SocketHandle As Long, ByVal ErrorCode As Integer)
    Dim foo As New frmSession

    foo.textSocketHandle = CStr(SocketHandle)
    'listStatus.AddItem "OnAccept" & Str(ErrorCode)
    listStatus.AddItem "Connected to " & foo.SocketX.SocketAddress
    If Trim$(foo.SocketX.RemoteAddress) <> "" Then
        foo.lblClientIP = foo.SocketX.RemoteAddress
    Else
        foo.lblClientIP = "Unknown IP"
    End If
    foo.lblClientPort = CStr(foo.SocketX.RemotePort)
End Sub

Private Sub AdminSocket_Accept(ByVal SocketHandle As Long, ByVal ErrorCode As Integer)
    Dim foo As New frmAdmin

    foo.textSocketHandle = CStr(SocketHandle)
    'listStatus.AddItem "OnAccept" & Str(ErrorCode)
    listStatus.AddItem "Admin Connection from " & foo.SocketX.SocketAddress
    If Trim$(foo.SocketX.RemoteAddress) <> "" Then
        foo.lblClientIP = foo.SocketX.RemoteAddress
    Else
        foo.lblClientIP = "Unknown IP"
    End If
    foo.lblClientPort = CStr(foo.SocketX.RemotePort)
End Sub

Private Sub SocketX_Close(ByVal ErrorCode As Integer)
    listStatus.AddItem "Close" & Str(ErrorCode)
    SocketX.Close
    SocketX.Create
    SocketX.Listen
End Sub

Private Sub AdminSocket_Close(ByVal ErrorCode As Integer)
    listStatus.AddItem "Closed Admin Connection " & Str(ErrorCode)
    AdminSocket.Close
    AdminSocket.Create
    AdminSocket.Listen
End Sub

Private Sub SocketX_Connect(ByVal ErrorCode As Integer)
    If ErrorCode <> 10065 And ErrorCode <> 10061 Then
        listStatus.AddItem "Connect " & Str(ErrorCode)
    End If
    '
    ' Remove all of the following lines if you have a known
    ' address.
    '
    RXStart
    SocketX.LocalPort = 21
    SocketX.LocalAddress = SocketX.SocketAddress
    textLocalAddress = SocketX.LocalAddress
    SocketX.Close
    SocketX.Create
    SocketX.Listen
    SocketX.KeepAliveEnabled = True
    RXEnd
    'listStatus.AddItem "Done determining local address"
End Sub

Private Sub AdminSocket_Connect(ByVal ErrorCode As Integer)
    If ErrorCode <> 10065 And ErrorCode <> 10061 Then
        listStatus.AddItem "Admin Socket Connected " & Str(ErrorCode)
    End If
    '
    ' Remove all of the following lines if you have a known
    ' address.
    '
    RXStart
    AdminSocket.LocalPort = 25
    AdminSocket.LocalAddress = AdminSocket.SocketAddress
    textLocalAddress = AdminSocket.LocalAddress
    AdminSocket.Close
    AdminSocket.Create
    AdminSocket.Listen
    RXEnd
    'listStatus.AddItem "Done determining local address"
End Sub

Private Sub MDIForm_Load()
    'Me.Show
    'listStatus.AddItem "Determining local address, please wait"
    DoEvents
    '
    ' note:  When running as a client on a service which
    ' provides you with an ip address when you connect the
    ' program must go through some hoops to get the assigned
    ' ip address to use with this server example.
    '
    ' This sample program was tested on such a server and
    ' handles the problem.  For a _real_ server this is
    ' somewhat unrealistic since a real server would
    ' be expected to have a known address and presumably
    ' you would know what that address is.
    '
    Set SocketX = New SocketXObj
    SocketX.Blocking = False
    SocketX.EventMask = -1
    SocketX.LocalPort = 0
    'SocketX.LocalAddress = "0.0.0.0"
    SocketX.LocalAddress = GetIPFromHostName(Environ$("computername"))
    
    Set AdminSocket = New SocketXObj
    AdminSocket.Blocking = False
    AdminSocket.EventMask = -1
    AdminSocket.LocalPort = 0
    'AdminSocket.LocalAddress = "0.0.0.0"
    AdminSocket.LocalAddress = GetIPFromHostName(Environ$("computername"))
    
    '
    ' Winsock won't tell us what our ip address is until
    ' after we connect to some remote socket.  Here I've
    ' chosen to use mit.edu.
    '
    SocketX.RemotePort = 13
    AdminSocket.RemotePort = 13
    '
    ' If you can't do name lookups comment out the following
    ' three lines and uncomment the one following them.
    '
    ' Also, if you have a known server address then remove
    ' the following three lines, uncomment the fourth, and
    ' supply the appropriate address there.  Be sure to remove
    ' the indicated lines in OnConnect as well in this case.
    '
    SocketX.RemoteNameAddrXlate = True
    SocketX.RemoteName = "10.1.1.2"
    SocketX.RemoteNameAddrXlate = False
    'SocketX.RemoteAddress = GetIPFromHostName("10.1.1.2")
    SocketX.RemoteAddress = "10.1.1.2"
    
    AdminSocket.RemoteNameAddrXlate = True
    AdminSocket.RemoteName = "10.1.1.2"
    AdminSocket.RemoteNameAddrXlate = False
    AdminSocket.RemoteAddress = "10.1.1.2"
    '    SocketX.RemoteAddress = "18.72.2.1"

    SocketX.Create
    AdminSocket.Create
    On Error Resume Next
    '
    ' Remove the following 5 lines if you have a known address
    '
    SocketX.Connect
    If (Err <> 0 And Err <> 10035) Then
        'MsgBox Error
    End If
    
    AdminSocket.Connect
    If (Err <> 0 And Err <> 10035) Then
        'MsgBox Error
    End If
    On Error GoTo 0
    '
    ' Uncomment the following two lines if you have a known
    ' address.
    '
    'SocketX.Action = ASocketCreate
    'SocketX.Action = ASocketListen
End Sub

Private Sub MDIForm_Resize()
    Line1.Y1 = Picture1.Height - Screen.TwipsPerPixelY
    Line1.Y2 = Line1.Y1
    Line2.Y1 = 0
    Line2.Y2 = 0
    Line1.X1 = 0
    Line1.X2 = Picture1.Width
    Line2.X1 = 0
    Line2.X2 = Picture1.Width
    listStatus.Width = Me.ScaleWidth - (listStatus.Left + 5 * Screen.TwipsPerPixelX)
End Sub

Public Sub RXStart()
    boolRx = True
    ledRx.LED_Colour = &HFF00&
    ledRx.Refresh
End Sub

Public Sub RXEnd()
    boolRx = False
    ledRx.LED_Colour = &HAEFA9E
    ledRx.Refresh
End Sub

Public Sub TXStart()
    boolTx = True
    ledTx.LED_Colour = &HFF00&
    ledTx.Refresh
End Sub

Public Sub TXEnd()
    boolTx = False
    ledTx.LED_Colour = &HAEFA9E
    ledTx.Refresh
End Sub

Private Sub Timer2_Timer()
    Dim ReplyTime As String, PingError As String
    Dim pingResult As String
    On Error Resume Next
    pngRemote.PingType = PingTypeICMP
    TXStart ' Transmit LED on
    
    'MonitoredIP = "192.168.0.16"
    MonitoredIP = "10.1.3.1"
    
    txtMonitored = MonitoredIP
    'pngRemote.Ping MonitoredIP, "ECHO" '    10.1.1.2
    pingResult = Ping.Ping(MonitoredIP, 2000)
    'Debug.Print Ping.PingTime & " " & Ping.ReturnCode & " " & p
    'PingError = pngRemote.LastErrorString
    
    Select Case pingResult
        Case "1048596"
            ledStatus.LED_Colour = &HFF&   ' Red
            lblMonitored = "Timed out"
        Case Else
            RXStart ' Receive LED on
            ledStatus.LED_Colour = &HFF00& ' Green
            lblMonitored = "Host OK"
    End Select
    
    
    
    'PingError = Err.Number
    'ReplyTime = pngRemote.ResponseTime
    'Debug.Print pngRemote.LastErrorString
    'Debug.Print Time & " " & PingError & " (" & Error(PingError) & ")" & " " & ReplyTime
    'lblMonitored = PingError
    'Select Case PingError
    '    Case "No error."
    '        RXStart ' Receive LED on
    '        ledStatus.LED_Colour = &HFF00& ' Green
    '    Case "Remote host is currently unreachable. "
    '        ledStatus.LED_Colour = &HFF&   ' Red
    '    Case Else
    '        ledStatus.LED_Colour = &H80FF& ' Orange
    'End Select
    
        
    'If pngRemote.LastErrorString <> "No error." Then
    '    ledStatus.LED_Colour = &HFF&   ' Red
    'ElseIf ReplyTime > "200" Then
    '    RXStart ' Receive LED on
    '    ledStatus.LED_Colour = &H80FF& ' Orange
    'Else
    '    RXStart ' Receive LED on
    '    ledStatus.LED_Colour = &HFF00& ' Green
    'End If
    TXEnd ' Transmit LED off
    RXEnd ' Receive LED off
End Sub


