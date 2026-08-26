VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Begin VB.Form frmShortcut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Shortcuts"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8760
   Icon            =   "frmShortcut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   5760
      Top             =   6840
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
            Picture         =   "frmShortcut.frx":0742
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcut.frx":0B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcut.frx":0FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShortcut.frx":1438
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fmeInformation 
      Caption         =   "Information"
      Height          =   3495
      Left            =   3720
      TabIndex        =   18
      Top             =   2400
      Width           =   4935
      Begin VB.CommandButton cmdPingAll 
         Caption         =   "Ping All"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdAllSpace 
         Caption         =   "Space (all)"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3000
         Width           =   975
      End
      Begin VB.ListBox lstSpace 
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton cmdSpace 
         Caption         =   "Disk Space"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pbrPing 
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   1560
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   10
      End
      Begin VB.CommandButton cmdTime 
         Caption         =   "Time"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "Ping"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin MSChart20Lib.MSChart chtPing 
         Height          =   1695
         Left            =   1080
         OleObjectBlob   =   "frmShortcut.frx":188A
         TabIndex        =   24
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblUsed 
         Caption         =   "used (E)"
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image imgStatus 
         Height          =   495
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblAverage 
         Alignment       =   2  'Center
         Caption         =   "Average:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblPing 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   2040
         Width           =   3495
      End
   End
   Begin VB.Frame fmeServices 
      Caption         =   "Services"
      Height          =   2175
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton cmdService 
         Caption         =   "Pause"
         Height          =   375
         Index           =   3
         Left            =   3600
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdService 
         Caption         =   "Start"
         Height          =   375
         Index           =   2
         Left            =   3600
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdService 
         Caption         =   "Stop"
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.ListBox lstServices 
         Height          =   1815
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdService 
         Caption         =   "Query"
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fmeOpen 
      Caption         =   "Open"
      Height          =   5775
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdIP 
         Caption         =   "Get IP"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "HTTP"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   29
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Server"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "C$"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "D$"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "E$"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Repl$"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Admin$"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Netlogon"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "VNC"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   4
         Top             =   5280
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc adoServers 
      Height          =   375
      Left            =   3000
      Top             =   6480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=CBDXAAI"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=CBDXAAI"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmShortcut.frx":3311
      Caption         =   "Servers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ListBox lstServers 
      Height          =   5325
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   11760
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11760
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   7215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
   End
End
Attribute VB_Name = "frmShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Const SW_ERASE = &H4
    Const SW_HIDE = 0
    Const SW_INVALIDATE = &H2
    Const SW_MAX = 10
    Const SW_MAXIMIZE = 3
    Const SW_MINIMIZE = 6
    Const SW_NORMAL = 1
    Const SW_OTHERUNZOOM = 4
    Const SW_OTHERZOOM = 2
    Const SW_PARENTCLOSING = 1
    Const SW_PARENTOPENING = 3
    Const SW_RESTORE = 9
    Const SW_SCROLLCHILDREN = &H1
    Const SW_SHOW = 5
    Const SW_SHOWDEFAULT = 10
    Const SW_SHOWMAXIMIZED = 3
    Const SW_SHOWMINIMIZED = 2
    Const SW_SHOWMINNOACTIVE = 7
    Const SW_SHOWNA = 8
    Const SW_SHOWNOACTIVATE = 4
    Const SW_SHOWNORMAL = 1
    
    Private Type TIME_OF_DAY
        t_elapsedt As Long
        t_msecs As Long
        t_hours As Long
        t_mins As Long
        t_secs As Long
        t_hunds As Long
        t_timezone As Long
        t_tinterval As Long
        t_day As Long
        t_month As Long
        t_year As Long
        t_weekday As Long
    End Type


    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function NetRemoteTOD Lib "netapi32.dll" (ByVal server As String, buffer As Any) As Long
    Private Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal buffer As Long) As Long
    Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
    '
    '
    'Initialize Public Collection
    Public Drives As New Collection
    
    ' Define GB and MB values
    Const lGByte As Long = 1073741824
    Const lMByte As Long = 1048576

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdIP_Click()
    Dim IP As New GetIP
    Dim IP2 As String
    adoServers.Refresh
    lblStatus = "Retrieving IP Address for selected Server"
    lblStatus.Refresh
    adoServers.Recordset.Move lstServers.ListIndex, 1
    IP.HostName = lstServers.Text
    IP2 = IP.IPAddress
    lblIP = "Db :" & adoServers.Recordset.Fields("IP") & vbCrLf & " Actual : " & IP2
    lblIP.Refresh
    lblStatus = "Ready"
    frmShortcut.Refresh
    Set IP = Nothing
End Sub

Private Sub cmdOpen_Click(Index As Integer)
    Dim server As String
    Dim res As Long
    Dim cmd As String
    
    server = lstServers.List(lstServers.ListIndex)
    If server <> "" Then
        Select Case Index
        
        Case 0
            cmd = "\\" & server & "\c$"
        Case 1
            cmd = "\\" & server & "\d$"
        Case 2
            cmd = "\\" & server & "\e$"
        Case 3
            cmd = "\\" & server & "\repl$"
        Case 4
            cmd = "\\" & server & "\admin$"
        Case 5
            cmd = "\\" & server & "\netlogon"
        Case 6
            cmd = "http://" & server & ":5800"
        Case 7
            cmd = "\\" & server
        Case 8
            cmd = "http://" & server
        End Select
        lblStatus = "Executing command"
        lblStatus.Refresh
        res = ShellExecute(hwnd, "open", cmd, vbNullString, CurDir$, SW_SHOW)
        lblStatus = "Ready"
        lblStatus.Refresh
    End If
End Sub

Private Sub cmdPing_Click()
    Dim fMatch As Boolean, sRTT As String, sHost As String
    Dim PingAverage As Long
    Dim IP As New GetIP

    cmdIP_Click
    PING_TIMEOUT = 3000
    IP.HostName = lstServers.Text
    sHost = IP.IPAddress
    'sHost = lblIP
    PingAverage = 0
    If lblIP <> "" Then
        imgStatus.Picture = imlStatus.ListImages(4).Picture
        lblStatus = "Pinging server with timeout of 3000 mS ..."
        pbrPing.Visible = True
        For lp = 1 To 10
            chtPing.Row = lp
            pbrPing = lp
            pingresult = ping(sHost, sRTT, True, 32, PING_TIMEOUT)
            If Val(pingresult) > 3000 Then pingresult = ""
            If pingresult = "" Then
                lblPing = ">3000 mS"
                chtPing.Data = 3000
                pingresult = 3000
            Else
                ' Found a host
               lblPing = pingresult & " mS"
               chtPing.Data = pingresult
            End If
            PingAverage = PingAverage + pingresult
            frmShortcut.Refresh
            DoEvents
        Next lp
        PingAverage = PingAverage / 10
        'chtPing
        lblPing = PingAverage & " mS"
        If PingAverage < 301 Then imgStatus.Picture = imlStatus.ListImages(1).Picture
        If PingAverage > 300 And PingAverage < 3000 Then imgStatus.Picture = imlStatus.ListImages(3).Picture
        If PingAverage = 3000 Then imgStatus.Picture = imlStatus.ListImages(2).Picture
        pbrPing.Visible = False
        lblStatus = "Ready"
        DoEvents
    End If
    Set IP = Nothing
End Sub

Private Sub cmdPingAll_Click()
    If Dir("c:\temp\pingtimes.csv") <> "" Then
        Kill "c:\temp\pingtimes.csv"
    End If
    
    For lp = 0 To lstServers.ListCount - 1
        DoEvents
        lstServers.ListIndex = lp
        cmdPing_Click
        frmShortcut.Refresh
        msg = lblPing.Caption
        
        
        Open "c:\temp\pingtimes.csv" For Append As #2
            Print #2, lstServers.List(lp) & "," & msg
        Close 2
    Next lp
End Sub

Private Sub cmdService_Click(Index As Integer)
    Dim server As String
    Dim service As String
    Dim retval As Long
    server = lstServers.List(lstServers.ListIndex)
    service = lstServices.List(lstServices.ListIndex)
    If server <> "" And service <> "" Then
        lblStatus = "Executing command... please wait"
        lblStatus.Refresh
        Screen.MousePointer = vbHourglass
        fmeServices.Enabled = False
        Select Case Index
            Case 0
                lblStatus = StatusofService(server, service)
                If lblStatus = "" Then
                    lblStatus = service & " Service on " & server & " not found, or error querying service status."
                End If
            Case 1
                msg = "Are you sure you want to stop the " & service & " Service on " & server & " ?"
                retval = MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation required")
                If retval = vbYes Then
                    vbStopService server, service
                    lblStatus = service & " Service on " & server & " has been stopped"
                End If
                
            Case 2
                vbStartService server, service
                lblStatus = service & " Service on " & server & " has been started"
            Case 3
                msg = "Are you sure you want to pause the " & service & " Service on " & server & " ?"
                retval = MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation required")
                If retval = vbYes Then
                    vbPauseService server, service
                    lblStatus = service & " Service on " & server & " has been paused"
                End If
        End Select
        Screen.MousePointer = vbDefault
        fmeServices.Enabled = True
    End If
End Sub

Private Sub cmdSpace_Click()
    Dim vDrive As Variant

    lstSpace.Clear
    ' Add the name of the Server you are checking to
    ' the top of the list of drives
    
    'lstSpace.AddItem "Server Name: " & lstServers.Text
    'lstSpace.AddItem "-----"
    
    ' Walk through the collection of possible default shares
    
    For Each vDrive In Drives
        DoEvents
        FindDriveInfo (CStr(vDrive))
        DoEvents
    Next
End Sub

Private Sub cmdTime_Click()
    Dim t As TIME_OF_DAY, tPtr As Long, res As Long, szServer As String, days As Date, todays As Date
    If lstServers.List(lstServers.ListIndex) <> "" Then
        lblTime = "Retrieving time..."
        lblTime.Refresh
        server = lstServers.List(lstServers.ListIndex)
        szServer = StrConv("\\" & server, vbUnicode)    'Convert the server name to unicode
        res = NetRemoteTOD(szServer, tPtr)              'You could also pass vbNullString for the server name
        If res = 0 Then
            CopyMemory t, ByVal tPtr, Len(t)            'Copy the pointer returned to a TIME_OF_DAY structure
            days = DateSerial(70, 1, 1) + (t.t_elapsedt / 60 / 60 / 24)  'Convert the elapsed time since 1/1/70 to a date
            days = days - (t.t_timezone / 60 / 24)      'Adjust for TimeZone differences
            lblTime = days
                                                        'Get local computer information for comparison
            todays = DateSerial(70, 1, 1) + (DateDiff("s", DateSerial(70, 1, 1), Now()) / 60 / 60 / 24)
                                                        'Print DateDiff("s", DateSerial(70, 1, 1), Now()), todays, t.t_elapsedt, days
            lblTime = days & " Difference: " & DateDiff("s", days, todays) & " Seconds"
            NetApiBufferFree (tPtr)                     'Free the memory at the pointer
        Else
            MsgBox "Error occurred call NetRemoteTOD: " & res, vbOKOnly, "NetRemoteTOD"
                                                        'Error 53: cannot find server
        End If
    End If
End Sub

Private Sub cmdAllSpace_Click()
    If Dir("c:\temp\drivespace.csv") <> "" Then
        Kill "c:\temp\drivespace.csv"
    End If
    
    For lp = 0 To lstServers.ListCount - 1
        DoEvents
        lstServers.ListIndex = lp
        cmdSpace_Click
        frmShortcut.Refresh
        msg = ""
        For lp2 = 0 To lstSpace.ListCount - 2
            msg = lstSpace.List(lp2)
            msg = msg & ","
        Next lp2
        msg = msg & lstSpace.List(lstSpace.ListCount - 1)
        Open "c:\temp\drivespace.csv" For Append As #2
            Print #2, lstServers.List(lp) & "," & msg
        Close 2
    Next lp
End Sub

Private Sub Form_Load()
    frmShortcut.Show
    lblStatus = "Retrieving servers from database"
    frmShortcut.Refresh
    adoServers.Refresh
    adoServers.Recordset.MoveFirst
    Do Until adoServers.Recordset.EOF
        server = adoServers.Recordset.Fields("Hostname")
        If server <> "" Then
            lstServers.AddItem server
        End If
        adoServers.Recordset.MoveNext
    Loop
    lblStatus = "Ready"
    frmShortcut.Refresh
    lstServices.AddItem "SNMP"
    lstServices.AddItem "VncSnapshot"
    lstServices.AddItem "Spooler"
    lstServices.AddItem "SMS_EXECUTIVE"
    lstServices.AddItem "LexBceS"
    lstServices.AddItem "MSExchangeDS"
    lstServices.AddItem "MSExchangeIS"
    lstServices.AddItem "MSExchangeIMS"
    lstServices.AddItem "MSExchangeMTA"
    lstServices.AddItem "MSExchangeSA"
    lstServices.AddItem "Replicator"
    lstServices.AddItem "clisvc"
    
    ' get the Domain/Workgroup info, fill Combo2 list
    'GetDomain (SV_TYPE_DOMAIN_ENUM)
                       
    ' Retrieve stored default Domain/Workgroup setting, if it exists
    'Combo2.Text = GetSetting("ShowDrives", "Configuration", "DefaultDomain", "Pick One")
    
    ' Add the names of possible default Administrative
    ' shares to the Public Collection 'Drives'
            
    'Drives.Add "C$"
    'Drives.Add "D$"
    Drives.Add "E$"
    'Drives.Add "F$"
    'Drives.Add "G$"
    'Drives.Add "H$"
    'Drives.Add "I$"
    'Drives.Add "J$"
    'Drives.Add "K$"
    'Drives.Add "L$"
    'Drives.Add "M$"
    'Drives.Add "N$"
    'Drives.Add "O$"
    'Drives.Add "P$"
    ' If you have more physical drives in your system than this, and
    ' they are not in a RAID configuration, then I think you need help ;-)
    'Drives.Add "Q$"
    'Drives.Add "R$"
    'Drives.Add "S$"
    'Drives.Add "T$"
    'Drives.Add "U$"
    'Drives.Add "V$"
    ' For the real sickos
    'Drives.Add "W$"
    'Drives.Add "X$"
    'Drives.Add "Y$"
    'Drives.Add "Z$"
End Sub

Private Sub lstServers_Click()
    Dim IP As New GetIP
    Dim IP2 As String
    'adoServers.Refresh
    'lblStatus = "Retrieving IP Address for selected Server"
    'lblStatus.Refresh
    'adoServers.Recordset.Move lstServers.ListIndex, 1
    'IP.HostName = lstServers.Text
    'IP2 = IP.IPAddress
    'lblIP = "Db :" & adoServers.Recordset.Fields("IP") & vbCrLf & " Actual : " & IP2
    'lblIP.Refresh
    lblStatus = "Ready"
    lblStatus.Refresh
    Set IP = Nothing
End Sub

Private Sub GetDomain(lType As Long)

    Dim lReturn As Long
    Dim Server_Info As Long
    Dim lEntries As Long
    Dim lTotal As Long
    Dim lMax As Long
    Dim vResume As Variant
    Dim tServer_info_101 As SERVER_INFO_101
    Dim sServer As String
    Dim sDomain As String
    Dim lServerInfo101StructPtr As Long
    Dim X As Long, i As Long
    Dim bBuffer(512) As Byte
    
    lReturn = NetServerEnum(ByVal 0&, 101, Server_Info, lMax, lEntries, lTotal, ByVal lType, sDomain, vResume)

    If lReturn <> 0 Then
        lblStatus.Caption = "Error: " & Err.LastDllError & "  has occurred"
        Exit Sub
    End If

    X = 1
    lServerInfo101StructPtr = Server_Info

    Do While X <= lTotal
        RtlMoveMemory tServer_info_101, ByVal lServerInfo101StructPtr, Len(tServer_info_101)
        lstrcpyW bBuffer(0), tServer_info_101.ptr_name
        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & Chr$(bBuffer(i))
            i = i + 2
        Loop
        'Combo2.AddItem sServer
        DoEvents
        X = X + 1
            sServer = ""
        lServerInfo101StructPtr = lServerInfo101StructPtr + Len(tServer_info_101)
    Loop
    lReturn = NetApiBufferFree(Server_Info)
End Sub

Public Sub FindDriveInfo(sDrive As String)
    Dim lReturn As Long
    Dim cBytesToCall As Currency
    Dim cBytesOnDrive As Currency
    Dim cFreeBytes As Currency
    Dim cUsedBytes As Currency
    Dim sServer As String
    Dim Output1 As String
    Dim Output2 As String
    Dim Output3 As String

    sServer = "\\" & lstServers.Text & "\" & sDrive
     
    ' Use in place of GetDiskFreeSpaceEx to bypass the
    ' 2 GB limit.  Use the Currency Data Type multiplied by
    ' 10000 so that 9,223,372,036,854,775,807 is the limit
    ' of the return value (see "Harcore Visual Basic", Chpt. 2
    ' for more information
            
    lReturn = GetDiskFreeSpaceEx(sServer, cBytesToCall, cBytesOnDrive, cFreeBytes)
    
    DoEvents
    
    If lReturn = 0 Then
        Exit Sub
    End If

    If lReturn <> 0 Then
    ' Format the output based on the returned value so that
    ' the string output is in MBs or GBs.

        ' Total Free Bytes on Drive
        Output1 = Format((cBytesOnDrive * 10000) / IIf(cBytesOnDrive * 10000 >= lGByte, lGByte, lMByte), IIf(cBytesOnDrive * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
                  
        ' Total Available Bytes on Drive
        Output2 = Format((cFreeBytes * 10000) / IIf(cFreeBytes * 10000 >= lGByte, lGByte, lMByte), IIf(cFreeBytes * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
                  
        ' Calculate Used Bytes
        cUsedBytes = cBytesOnDrive - cFreeBytes
        
        ' Total Used Bytes on Drive
        Output3 = Format((cUsedBytes * 10000) / IIf(cUsedBytes * 10000 >= lGByte, lGByte, lMByte), IIf(cUsedBytes * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
    
    End If
    
    ' If Output1 (Total Free Bytes on Drive) is 0, then
    ' you know the drive does not exist.  Otherwise, port
    ' the output to lstSpace
    
    If Output1 <> "0.MB" Then
        freepercent = Int((cFreeBytes / cBytesOnDrive) * 100)
        usedpercent = Int((cUsedBytes / cBytesOnDrive) * 100)
        'lstSpace.AddItem sDrive & " - " & freepercent & "% free, " & usedpercent & "% used"
        lstSpace.AddItem usedpercent & "%"
        
        'lstSpace.AddItem " Total Bytes on Drive: " & Output1
        'lstSpace.AddItem " Total Bytes Free: " & Output2
        'lstSpace.AddItem " Total Bytes Used: " & Output3
        'lstSpace.AddItem " "
    Else
        
    End If
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub
