VERSION 5.00
Object = "{2B2EA379-3B26-42D5-873B-F645B57B5C56}#1.0#0"; "GetHostX.ocx"
Object = "{0468C941-83E2-11D3-BE51-00C0DFC2E32C}#1.0#0"; "PingX.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resolver"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5490
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   285
      Left            =   3960
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtMaxtime 
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   0
      Left            =   5160
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Timer tmrPing 
      Interval        =   750
      Left            =   3240
      Top             =   480
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   3120
      Width           =   855
   End
   Begin VB.CheckBox chkContinuous 
      Caption         =   "Continuous Ping"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox chkSyncTime 
      Caption         =   "Sync time with Host"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox chkPing 
      Caption         =   "Ping Host"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox chkGetTime 
      Caption         =   "Get time from Host"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin MabryPingXCtl.PingX pngHost 
      Left            =   3720
      Top             =   480
      RemoteAddress   =   ""
      RemotePort      =   0
      Timeout         =   5
      DebugMode       =   0   'False
      BlockingMode    =   0
      Blocking        =   0   'False
      RequestLength   =   64
      RequestData     =   "ECHO"
      LibraryName     =   "WSOCK32.DLL"
   End
   Begin GetHostOCXLibCtl.GetHostX nmeHost 
      Left            =   2640
      Top             =   480
      Host            =   ""
      Blocking        =   0   'False
      BlockingMode    =   0
      DebugMode       =   0   'False
      Timeout         =   60
      LicenseKey      =   ""
      Type            =   3
   End
   Begin VB.CommandButton cmdResolve 
      Caption         =   "Resolve"
      Default         =   -1  'True
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtHostname 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   1
      Left            =   4920
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   2
      Left            =   4680
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   3
      Left            =   4440
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   4
      Left            =   4200
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   5
      Left            =   3960
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   6
      Left            =   3720
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   7
      Left            =   3480
      TabIndex        =   23
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   8
      Left            =   3240
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1455
      Index           =   9
      Left            =   3000
      TabIndex        =   25
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblPeriodDown 
      Height          =   255
      Left            =   840
      TabIndex        =   32
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblPeriod 
      Caption         =   "Period"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblTimeDown 
      Height          =   255
      Left            =   840
      TabIndex        =   30
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblTimeup 
      Height          =   255
      Left            =   840
      TabIndex        =   29
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblDownat 
      Caption         =   "Down at:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblLastUp 
      Caption         =   "Last up:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblHostInfo 
      Alignment       =   2  'Center
      Caption         =   "Host time info"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblMaxtime 
      Caption         =   "Max:"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPingtime 
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblHosttime 
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblHostname 
      Caption         =   "Hostname"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lngMaxTime As Long
Private timeDown As Date
Private timeUp As Date
Private timePeriod As Long

Private Sub chkContinuous_Click()
  If chkContinuous.Value = vbChecked Then
    tmrPing.Enabled = True
    For lp = pbrPing.LBound To pbrPing.UBound
      pbrPing(lp).Visible = True
    Next lp
    lblMaxtime.Visible = True
    txtMaxtime.Visible = True
    cmdSet.Visible = True
  Else
    tmrPing.Enabled = False
    For lp = pbrPing.LBound To pbrPing.UBound
      pbrPing(lp).Visible = False
    Next lp
    lblMaxtime.Visible = False
    txtMaxtime.Visible = False
    cmdSet.Visible = False
  End If
End Sub

Private Sub chkGetTime_Click()
  If chkGetTime.Value = vbChecked Then
    chkSyncTime.Enabled = True
  Else
    chkSyncTime.Enabled = False
  End If
End Sub

Private Sub chkPing_Click()
If chkPing.Value = vbChecked Then
    chkContinuous.Enabled = True
  Else
    chkContinuous.Enabled = False
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
  End
End Sub

Private Sub cmdResolve_Click()
  txtHostname.SetFocus
  ' Reset values
  timeDown = "0:00:00"
  timeUp = "0:00:00"
  timePeriod = 0
  lblTimeup = ""
  lblTimeDown = ""
  lblPeriodDown = ""
  For lp = pbrPing.LBound To pbrPing.UBound
    pbrPing(lp).Value = 0
  Next lp
  If txtHostname <> "" Then
    lblStatus = "Resolving " & txtHostname
    lblStatus.Refresh
    nmeHost.Host = txtHostname
    nmeHost.Resolve
  ElseIf txtAddress <> "" Then
    lblStatus = "Resolving " & txtAddress
    lblStatus.Refresh
    nmeHost.Host = txtAddress
    nmeHost.Resolve
  Else
    lblStatus = "Cannot resolve NULL"
    lblStatus.Refresh
    Exit Sub
  End If
End Sub

Private Sub cmdSet_Click()
  If IsNumeric(txtMaxtime) Then
    lngMaxTime = CLng(txtMaxtime.Text)
    For lp = pbrPing.LBound To pbrPing.UBound
      pbrPing(lp).Max = lngMaxTime
    Next lp
  Else
    txtMaxtime = ""
  End If
End Sub

Private Sub Form_Load()
  lngMaxTime = 500
  txtMaxtime = lngMaxTime
End Sub

Private Sub nmeHost_Done(ByVal Hostname As String, ByVal IPAddress As String, ByVal ErrorCode As Long, ByVal ErrorString As String)
  If IPAddress <> "" And Hostname <> "" Then
    txtHostname = Hostname
    txtAddress = IPAddress
    If chkGetTime.Value = vbChecked Then
      Dim timHost As NetHostTime.HostTime
      Set timHost = New NetHostTime.HostTime
      lblStatus = "Retrieving time from " & txtHostname
      lblStatus.Refresh
      lblHosttime = timHost.GetTime(txtHostname)
      If IsDate(lblHosttime) And chkSyncTime.Value = vbChecked Then
        Time = CDate(lblHosttime)
      End If
      If lblHosttime <> "" Then
        lblHosttime = "Time: " & lblHosttime
        lblStatus.Refresh
      Else
        lblHosttime = "Time unknown"
        lblStatus.Refresh
      End If
      Set timHost = Nothing
    End If
    lblStatus = "Idle"
    If chkPing.Value = vbChecked Then
      pngHost.RemoteAddress = txtAddress
      lblStatus = "Pinging " & txtAddress
      lblStatus.Refresh
      pngHost.Ping
    End If
  Else
    lblStatus = ErrorString
    lblStatus.Refresh
  End If
End Sub

Private Sub pngHost_Done(ByVal RemoteAddress As String, ByVal ResponseTime As Long, ByVal ErrorCode As Integer, ByVal ErrorString As String)
  Static Stack As Integer
  lblPingtime = "Time: " & ResponseTime
  lblStatus = "Idle"
  lblStatus.Refresh
  If ResponseTime > lngMaxTime Then ResponseTime = lngMaxTime
  If chkContinuous.Value = vbChecked Then
    If Stack >= pbrPing.UBound Then Stack = pbrPing.UBound
    For lp = Stack To pbrPing.LBound + 1 Step -1
      If pbrPing(lp - 1).Value > lngMaxTime Then
        pbrPing(lp).Value = lngMaxTime
      Else
        pbrPing(lp).Value = pbrPing(lp - 1).Value
      End If
    Next lp
    pbrPing(0).Value = ResponseTime
    Stack = Stack + 1
  Else
    Stack = 0
  End If
  If ErrorString = "The current connection has timedout." Then ' Timeout
    lblStatus = "Host down at " & Time
    lblStatus.Refresh
    If CStr(timeDown) = "0:00:00" Then
      timeDown = Now
      lblTimeDown = timeDown
    End If
  ElseIf ErrorString <> "No error." Then                       ' Error
    lblStatus = ErrorString
    lblStatus.Refresh
  Else
    lblStatus = "Response OK"                                  ' OK
    lblStatus.Refresh
    timeUp = Now
    lblTimeup = timeUp
  End If
  If lblTimeDown <> "" Then
    If timeDown < timeUp Then
      timePeriod = DateDiff("s", timeDown, timeUp)
      lblPeriodDown = timePeriod & " s"
    Else
      timePeriod = DateDiff("s", timeDown, Now)
      lblPeriodDown = timePeriod & " s"
    End If
  Else
      
  End If
End Sub

Private Sub tmrPing_Timer()
  If txtAddress = "" Then Exit Sub
  If chkContinuous.Value = vbChecked Then
    pngHost.RemoteAddress = txtAddress
    lblStatus = "Pinging " & txtAddress
    lblStatus.Refresh
    tmrPing.Enabled = False
    If pngHost.State <> WaitingResponse Then
      pngHost.Ping
    End If
    tmrPing.Enabled = True
  End If
End Sub

Private Sub txtMaxtime_GotFocus()
  txtMaxtime.SelStart = 0
  txtMaxtime.SelLength = Len(txtMaxtime)
End Sub
