VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStats 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Statistics"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5115
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
   Picture         =   "FrmStats.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Interval        =   250
      Left            =   4680
      Top             =   2880
   End
   Begin VB.Timer Timer4 
      Interval        =   250
      Left            =   4680
      Top             =   2400
   End
   Begin VB.Timer Timer3 
      Interval        =   250
      Left            =   4680
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   4680
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4680
      Top             =   960
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   5442
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   5442
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   5442
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ListView ListView4 
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   5442
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ListView ListView5 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   5442
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ICMP (out)"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   4
      Left            =   3360
      Picture         =   "FrmStats.frx":35B8
      Top             =   495
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ICMP(in)"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   3
      Left            =   2280
      Picture         =   "FrmStats.frx":39A5
      Top             =   495
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UDP"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   2
      Left            =   1440
      Picture         =   "FrmStats.frx":3D92
      Top             =   495
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IP"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   1
      Left            =   840
      Picture         =   "FrmStats.frx":417F
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TCP"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   0
      Left            =   120
      Picture         =   "FrmStats.frx":456C
      Top             =   495
      Width           =   180
   End
   Begin VB.Image Image9 
      Height          =   180
      Left            =   4800
      Picture         =   "FrmStats.frx":4959
      Top             =   3840
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image6 
      Height          =   180
      Left            =   4800
      Picture         =   "FrmStats.frx":4D51
      Top             =   3480
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   4680
      Picture         =   "FrmStats.frx":513E
      Top             =   240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image8 
      Height          =   315
      Left            =   4680
      Picture         =   "FrmStats.frx":5633
      Top             =   600
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   4200
      Picture         =   "FrmStats.frx":5B55
      Tag             =   "Close"
      ToolTipText     =   "Close"
      Top             =   45
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statistics"
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
      TabIndex        =   1
      Tag             =   "TitleColor"
      Top             =   0
      Width           =   1155
   End
End
Attribute VB_Name = "FrmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Dim IP As MIB_IPSTATS
Dim tcp As MIB_TCPSTATS
Dim udp As MIB_UDPSTATS
Dim icmp As MIBICMPINFO
Dim tStats As MIB_TCPSTATS

Dim sControlSelected As String

Private Sub Form_Load()

sControlSelected = "tcp"
Call imgSelected_Click(0)

Call LoadColors
Call SetColors(Me)

Me.Width = 4560

    With ListView1.ListItems
        '
        .Add , , "Timeout algorithm"
        .Add , , "Minimum timeout"
        .Add , , "Maximum timeout"
        .Add , , "Maximum connections"
        .Add , , "Active opens"
        .Add , , "Passive opens"
        .Add , , "Failed attempts"
        .Add , , "Establised connections reset"
        .Add , , "Established connections"
        .Add , , "Segments received"
        .Add , , "Segment sent"
        .Add , , "Segments retransmitted"
        .Add , , "Incoming errors"
        .Add , , "Outgoing resets"
        .Add , , "Cumulative connections"
        '
    End With
    
    With ListView2.ListItems
    .Add , , "IP forwarding enabled or disabled"
    .Add , , "Default time-to-live"
    .Add , , "Datagrams received"
    .Add , , "Received header errors"
    .Add , , "Received address errors"
    .Add , , "datagrams forwarded"
    .Add , , "datagrams with unknown protocol"
    .Add , , "received datagrams discarded"
    .Add , , "received datagrams delivered"
    .Add , , "outgoing datagrams requested"
    .Add , , "outgoing datagrams discarded"
    .Add , , "sent datagrams discarded"
    .Add , , "datagrams for which no route"
    .Add , , "datagrams for which all frags didn't arrive"
    .Add , , "datagrams requiring reassembly"
    .Add , , "successful reassemblies"
    .Add , , "failed reassemblies"
    .Add , , "successful fragmentations"
    .Add , , "failed fragmentations"
    .Add , , "datagrams fragmented"
    .Add , , "number of interfaces on computer"
    .Add , , "number of IP address on computer"
    .Add , , "number of routes in routing table"
    End With

    With ListView3.ListItems
    .Add , , "received datagrams"
    .Add , , "datagrams for which no port"
    .Add , , "errors on received datagrams"
    .Add , , "sent datagrams"
    .Add , , "number of entries in UDP listener table"
    End With
    
    With ListView4.ListItems
    .Add , , "number of messages"
    .Add , , "number of errors"
    .Add , , "destination unreachable messages"
    .Add , , "time-to-live exceeded messages"
    .Add , , "parameter problem messages"
    .Add , , "source quench messages"
    .Add , , "redirection messages"
    .Add , , "echo requests"
    .Add , , "echo replies"
    .Add , , "timestamp requests"
    .Add , , "timestamp replies"
    .Add , , "address mask requests"
    .Add , , "address mask replies"
    End With
    
    With ListView5.ListItems
    .Add , , "number of messages"
    .Add , , "number of errors"
    .Add , , "destination unreachable messages"
    .Add , , "time-to-live exceeded messages"
    .Add , , "parameter problem messages"
    .Add , , "source quench messages"
    .Add , , "redirection messages"
    .Add , , "echo requests"
    .Add , , "echo replies"
    .Add , , "timestamp requests"
    .Add , , "timestamp replies"
    .Add , , "address mask requests"
    .Add , , "address mask replies"
    End With

Call GetTcpStatistics(tStats)

With tStats
ListView1.ListItems(1).SubItems(1) = .dwRtoAlgorithm
ListView1.ListItems(2).SubItems(1) = .dwRtoMin
ListView1.ListItems(3).SubItems(1) = .dwRtoMax
ListView1.ListItems(4).SubItems(1) = .dwMaxConn
ListView1.ListItems(5).SubItems(1) = .dwActiveOpens
ListView1.ListItems(6).SubItems(1) = .dwPassiveOpens
ListView1.ListItems(7).SubItems(1) = .dwAttemptFails
ListView1.ListItems(8).SubItems(1) = .dwEstabResets
ListView1.ListItems(9).SubItems(1) = .dwCurrEstab
ListView1.ListItems(10).SubItems(1) = .dwInSegs
ListView1.ListItems(11).SubItems(1) = .dwOutSegs
ListView1.ListItems(12).SubItems(1) = .dwRetransSegs
ListView1.ListItems(13).SubItems(1) = .dwInErrs
ListView1.ListItems(14).SubItems(1) = .dwOutRsts
ListView1.ListItems(15).SubItems(1) = .dwNumConns
End With
DoEvents

Call GetIpStatistics(IP)

With IP
ListView2.ListItems(1).SubItems(1) = .dwForwarding
ListView2.ListItems(2).SubItems(1) = .dwDefaultTTL
ListView2.ListItems(3).SubItems(1) = .dwInReceives
ListView2.ListItems(4).SubItems(1) = .dwInHdrErrors
ListView2.ListItems(5).SubItems(1) = .dwInAddrErrors
ListView2.ListItems(6).SubItems(1) = .dwForwDatagrams
ListView2.ListItems(7).SubItems(1) = .dwInUnknownProtos
ListView2.ListItems(8).SubItems(1) = .dwInDiscards
ListView2.ListItems(9).SubItems(1) = .dwInDelivers
ListView2.ListItems(10).SubItems(1) = .dwOutRequests
ListView2.ListItems(11).SubItems(1) = .dwRoutingDiscards
ListView2.ListItems(12).SubItems(1) = .dwOutDiscards
ListView2.ListItems(13).SubItems(1) = .dwOutNoRoutes
ListView2.ListItems(14).SubItems(1) = .dwReasmTimeout
ListView2.ListItems(15).SubItems(1) = .dwReasmReqds
ListView2.ListItems(16).SubItems(1) = .dwReasmOks
ListView2.ListItems(17).SubItems(1) = .dwReasmFails
ListView2.ListItems(18).SubItems(1) = .dwFragOks
ListView2.ListItems(19).SubItems(1) = .dwFragFails
ListView2.ListItems(20).SubItems(1) = .dwFragCreates
ListView2.ListItems(21).SubItems(1) = .dwNumIf
ListView2.ListItems(22).SubItems(1) = .dwNumAddr
ListView2.ListItems(23).SubItems(1) = .dwNumRoutes
End With
DoEvents

Call GetUdpStatistics(udp)

With udp
ListView3.ListItems(1).SubItems(1) = .dwInDatagrams
ListView3.ListItems(2).SubItems(1) = .dwNoPorts
ListView3.ListItems(3).SubItems(1) = .dwInErrors
ListView3.ListItems(4).SubItems(1) = .dwOutDatagrams
ListView3.ListItems(5).SubItems(1) = .dwNumAddrs
End With
DoEvents

Call GetIcmpStatistics(icmp)

With icmp
ListView4.ListItems(1).SubItems(1) = .icmpInStats.dwMsgs
ListView4.ListItems(2).SubItems(1) = .icmpInStats.dwErrors
ListView4.ListItems(3).SubItems(1) = .icmpInStats.dwDestUnreachs
ListView4.ListItems(4).SubItems(1) = .icmpInStats.dwTimeExcds
ListView4.ListItems(5).SubItems(1) = .icmpInStats.dwParmProbs
ListView4.ListItems(6).SubItems(1) = .icmpInStats.dwSrcQuenchs
ListView4.ListItems(7).SubItems(1) = .icmpInStats.dwRedirects
ListView4.ListItems(8).SubItems(1) = .icmpInStats.dwEchos
ListView4.ListItems(9).SubItems(1) = .icmpInStats.dwEchoReps
ListView4.ListItems(10).SubItems(1) = .icmpInStats.dwTimestamps
ListView4.ListItems(11).SubItems(1) = .icmpInStats.dwTimestampReps
ListView4.ListItems(12).SubItems(1) = .icmpInStats.dwAddrMasks
ListView4.ListItems(13).SubItems(1) = .icmpInStats.dwAddrMaskReps
DoEvents
ListView5.ListItems(1).SubItems(1) = .icmpOutStats.dwMsgs
ListView5.ListItems(2).SubItems(1) = .icmpOutStats.dwErrors
ListView5.ListItems(3).SubItems(1) = .icmpOutStats.dwDestUnreachs
ListView5.ListItems(4).SubItems(1) = .icmpOutStats.dwTimeExcds
ListView5.ListItems(5).SubItems(1) = .icmpOutStats.dwParmProbs
ListView5.ListItems(6).SubItems(1) = .icmpOutStats.dwSrcQuenchs
ListView5.ListItems(7).SubItems(1) = .icmpOutStats.dwRedirects
ListView5.ListItems(8).SubItems(1) = .icmpOutStats.dwEchos
ListView5.ListItems(9).SubItems(1) = .icmpOutStats.dwEchoReps
ListView5.ListItems(10).SubItems(1) = .icmpOutStats.dwTimestamps
ListView5.ListItems(11).SubItems(1) = .icmpOutStats.dwTimestampReps
ListView5.ListItems(12).SubItems(1) = .icmpOutStats.dwAddrMasks
ListView5.ListItems(13).SubItems(1) = .icmpOutStats.dwAddrMaskReps
End With
DoEvents
End Sub
Private Sub imgClose_Click()
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
        Unload FrmStats
        bBottomOut = False
    End If

End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgClose.Picture = Image5.Picture
End If

End Sub
Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgClose.Picture = Image8.Picture
End If

End Sub
Private Sub UpdateStats1()
On Error Resume Next

    Dim tStats          As MIB_TCPSTATS
    Static tStaticStats As MIB_TCPSTATS
    '
    Dim lRetValue       As Long
    '
    Dim blnIsSent       As Boolean
    Dim blnIsRecv       As Boolean
    '
    lRetValue = GetTcpStatistics(tStats)
    '
    With tStats
        '
        If Not tStaticStats.dwRtoAlgorithm = .dwRtoAlgorithm Then _
            ListView1.ListItems(1).SubItems(1) = .dwRtoAlgorithm
        If Not tStaticStats.dwRtoMin = .dwRtoMin Then _
            ListView1.ListItems(2).SubItems(1) = .dwRtoMin
        If Not tStaticStats.dwRtoMax = .dwRtoMax Then _
            ListView1.ListItems(3).SubItems(1) = .dwRtoMax
        If Not tStaticStats.dwMaxConn = .dwMaxConn Then _
            ListView1.ListItems(4).SubItems(1) = .dwMaxConn
        If Not tStaticStats.dwActiveOpens = .dwActiveOpens Then _
            ListView1.ListItems(5).SubItems(1) = .dwActiveOpens
        If Not tStaticStats.dwPassiveOpens = .dwPassiveOpens Then _
            ListView1.ListItems(6).SubItems(1) = .dwPassiveOpens
        If Not tStaticStats.dwAttemptFails = .dwAttemptFails Then _
            ListView1.ListItems(7).SubItems(1) = .dwAttemptFails
        If Not tStaticStats.dwEstabResets = .dwEstabResets Then _
            ListView1.ListItems(8).SubItems(1) = .dwEstabResets
        If Not tStaticStats.dwCurrEstab = .dwCurrEstab Then _
            ListView1.ListItems(9).SubItems(1) = .dwCurrEstab
        If Not tStaticStats.dwInSegs = .dwInSegs Then _
            ListView1.ListItems(10).SubItems(1) = .dwInSegs
        If Not tStaticStats.dwOutSegs = .dwOutSegs Then _
            ListView1.ListItems(11).SubItems(1) = .dwOutSegs
        If Not tStaticStats.dwRetransSegs = .dwRetransSegs Then _
            ListView1.ListItems(12).SubItems(1) = .dwRetransSegs
        If Not tStaticStats.dwInErrs = .dwInErrs Then _
            ListView1.ListItems(13).SubItems(1) = .dwInErrs
        If Not tStaticStats.dwOutRsts = .dwOutRsts Then _
            ListView1.ListItems(14).SubItems(1) = .dwOutRsts
        If Not tStaticStats.dwNumConns = .dwNumConns Then _
            ListView1.ListItems(15).SubItems(1) = .dwNumConns
        '
    End With

    tStaticStats = tStats

End Sub
Private Sub UpdateStats2()

On Error Resume Next
Static ip2 As MIB_IPSTATS
Dim lRetValue       As Long

lRetValue = GetIpStatistics(IP)

With IP
If Not ip2.dwForwarding = .dwForwarding Then _
ListView2.ListItems(1).SubItems(1) = .dwForwarding
If Not ip2.dwDefaultTTL = .dwDefaultTTL Then _
ListView2.ListItems(2).SubItems(1) = .dwDefaultTTL
If Not ip2.dwInReceives = .dwInReceives Then _
ListView2.ListItems(3).SubItems(1) = .dwInReceives
If Not ip2.dwInHdrErrors = .dwInHdrErrors Then _
ListView2.ListItems(4).SubItems(1) = .dwInHdrErrors
If Not ip2.dwInAddrErrors = .dwInAddrErrors Then _
ListView2.ListItems(5).SubItems(1) = .dwInAddrErrors
If Not ip2.dwForwDatagrams = .dwForwDatagrams Then _
ListView2.ListItems(6).SubItems(1) = .dwForwDatagrams
If Not ip2.dwInUnknownProtos = .dwInUnknownProtos Then _
ListView2.ListItems(7).SubItems(1) = .dwInUnknownProtos
If Not ip2.dwInDiscards = .dwInDiscards Then _
ListView2.ListItems(8).SubItems(1) = .dwInDiscards
If Not ip2.dwInDelivers = .dwInDelivers Then _
ListView2.ListItems(9).SubItems(1) = .dwInDelivers
If Not ip2.dwOutRequests = .dwOutRequests Then _
ListView2.ListItems(10).SubItems(1) = .dwOutRequests
If Not ip2.dwRoutingDiscards = .dwRoutingDiscards Then _
ListView2.ListItems(11).SubItems(1) = .dwRoutingDiscards
If Not ip2.dwOutDiscards = .dwOutDiscards Then _
ListView2.ListItems(12).SubItems(1) = .dwOutDiscards
If Not ip2.dwOutNoRoutes = .dwOutNoRoutes Then _
ListView2.ListItems(13).SubItems(1) = .dwOutNoRoutes
If Not ip2.dwReasmTimeout = .dwReasmTimeout Then _
ListView2.ListItems(14).SubItems(1) = .dwReasmTimeout
If Not ip2.dwReasmReqds = .dwReasmReqds Then _
ListView2.ListItems(15).SubItems(1) = .dwReasmReqds
If Not ip2.dwReasmOks = .dwReasmOks Then _
ListView2.ListItems(16).SubItems(1) = .dwReasmOks
If Not ip2.dwReasmFails = .dwReasmFails Then _
ListView2.ListItems(17).SubItems(1) = .dwReasmFails
If Not ip2.dwFragOks = .dwFragOks Then _
ListView2.ListItems(18).SubItems(1) = .dwFragOks
If Not ip2.dwFragFails = .dwFragFails Then _
ListView2.ListItems(19).SubItems(1) = .dwFragFails
If Not ip2.dwFragCreates = .dwFragCreates Then _
ListView2.ListItems(20).SubItems(1) = .dwFragCreates
If Not ip2.dwNumIf = .dwNumIf Then _
ListView2.ListItems(21).SubItems(1) = .dwNumIf
If Not ip2.dwNumAddr = .dwNumAddr Then _
ListView2.ListItems(22).SubItems(1) = .dwNumAddr
If Not ip2.dwNumRoutes = .dwNumRoutes Then _
ListView2.ListItems(23).SubItems(1) = .dwNumRoutes
End With

ip2 = IP


End Sub

Private Sub UpdateStats3()

On Error Resume Next
Dim lRetValue       As Long
Static udp2 As MIB_UDPSTATS

lRetValue = GetUdpStatistics(udp)

With udp
If Not udp2.dwInDatagrams = .dwInDatagrams Then _
ListView3.ListItems(1).SubItems(1) = .dwInDatagrams

If Not udp2.dwNoPorts = .dwNoPorts Then _
ListView3.ListItems(2).SubItems(1) = .dwNoPorts

If Not udp2.dwInErrors = .dwInErrors Then _
ListView3.ListItems(3).SubItems(1) = .dwInErrors

If Not udp2.dwOutDatagrams = .dwOutDatagrams Then _
ListView3.ListItems(4).SubItems(1) = .dwOutDatagrams

If Not udp2.dwNumAddrs = .dwNumAddrs Then _
ListView3.ListItems(5).SubItems(1) = .dwNumAddrs

End With

udp2 = udp

End Sub
Private Sub UpdateStats4()
On Error Resume Next

Dim lRetValue       As Long
Static icmp2 As MIBICMPINFO

lRetValue = GetIcmpStatistics(icmp)

With icmp
If Not icmp2.icmpOutStats.dwMsgs = .icmpOutStats.dwMsgs Then _
ListView4.ListItems(1).SubItems(1) = .icmpOutStats.dwMsgs
If Not icmp2.icmpOutStats.dwErrors = .icmpOutStats.dwErrors Then _
ListView4.ListItems(2).SubItems(1) = .icmpOutStats.dwErrors
If Not icmp2.icmpOutStats.dwDestUnreachs = .icmpOutStats.dwDestUnreachs Then _
ListView4.ListItems(3).SubItems(1) = .icmpOutStats.dwDestUnreachs
If Not icmp2.icmpOutStats.dwTimeExcds = .icmpOutStats.dwTimeExcds Then _
ListView4.ListItems(4).SubItems(1) = .icmpOutStats.dwTimeExcds
If Not icmp2.icmpOutStats.dwParmProbs = .icmpOutStats.dwParmProbs Then _
ListView4.ListItems(5).SubItems(1) = .icmpOutStats.dwParmProbs
If Not icmp2.icmpOutStats.dwSrcQuenchs = .icmpOutStats.dwSrcQuenchs Then _
ListView4.ListItems(6).SubItems(1) = .icmpOutStats.dwSrcQuenchs
If Not icmp2.icmpOutStats.dwRedirects = .icmpOutStats.dwRedirects Then _
ListView4.ListItems(7).SubItems(1) = .icmpOutStats.dwRedirects
If Not icmp2.icmpOutStats.dwEchos = .icmpOutStats.dwEchos Then _
ListView4.ListItems(8).SubItems(1) = .icmpOutStats.dwEchos
If Not icmp2.icmpOutStats.dwEchoReps = .icmpOutStats.dwEchoReps Then _
ListView4.ListItems(9).SubItems(1) = .icmpOutStats.dwEchoReps
If Not icmp2.icmpOutStats.dwTimestamps = .icmpOutStats.dwTimestamps Then _
ListView4.ListItems(10).SubItems(1) = .icmpOutStats.dwTimestamps
If Not icmp2.icmpOutStats.dwTimestampReps = .icmpOutStats.dwTimestampReps Then _
ListView4.ListItems(11).SubItems(1) = .icmpOutStats.dwTimestampReps
If Not icmp2.icmpOutStats.dwAddrMasks = .icmpOutStats.dwAddrMasks Then _
ListView4.ListItems(12).SubItems(1) = .icmpOutStats.dwAddrMasks
If Not icmp2.icmpOutStats.dwAddrMaskReps = .icmpOutStats.dwAddrMaskReps Then _
ListView4.ListItems(13).SubItems(1) = .icmpOutStats.dwAddrMaskReps
End With

icmp2 = icmp

End Sub
Private Sub UpdateStats5()

On Error Resume Next
Dim lRetValue       As Long
Static icmp2 As MIBICMPINFO

lRetValue = GetIcmpStatistics(icmp)

With icmp
If Not icmp2.icmpInStats.dwMsgs = .icmpInStats.dwMsgs Then _
ListView4.ListItems(1).SubItems(1) = .icmpInStats.dwMsgs
If Not icmp2.icmpInStats.dwErrors = .icmpInStats.dwErrors Then _
ListView4.ListItems(2).SubItems(1) = .icmpInStats.dwErrors
If Not icmp2.icmpInStats.dwDestUnreachs = .icmpInStats.dwDestUnreachs Then _
ListView4.ListItems(3).SubItems(1) = .icmpInStats.dwDestUnreachs
If Not icmp2.icmpInStats.dwTimeExcds = .icmpInStats.dwTimeExcds Then _
ListView4.ListItems(4).SubItems(1) = .icmpInStats.dwTimeExcds
If Not icmp2.icmpInStats.dwParmProbs = .icmpInStats.dwParmProbs Then _
ListView4.ListItems(5).SubItems(1) = .icmpInStats.dwParmProbs
If Not icmp2.icmpInStats.dwSrcQuenchs = .icmpInStats.dwSrcQuenchs Then _
ListView4.ListItems(6).SubItems(1) = .icmpInStats.dwSrcQuenchs
If Not icmp2.icmpInStats.dwRedirects = .icmpInStats.dwRedirects Then _
ListView4.ListItems(7).SubItems(1) = .icmpInStats.dwRedirects
If Not icmp2.icmpInStats.dwEchos = .icmpInStats.dwEchos Then _
ListView4.ListItems(8).SubItems(1) = .icmpInStats.dwEchos
If Not icmp2.icmpInStats.dwEchoReps = .icmpInStats.dwEchoReps Then _
ListView4.ListItems(9).SubItems(1) = .icmpInStats.dwEchoReps
If Not icmp2.icmpInStats.dwTimestamps = .icmpInStats.dwTimestamps Then _
ListView4.ListItems(10).SubItems(1) = .icmpInStats.dwTimestamps
If Not icmp2.icmpInStats.dwTimestampReps = .icmpInStats.dwTimestampReps Then _
ListView4.ListItems(11).SubItems(1) = .icmpInStats.dwTimestampReps
If Not icmp2.icmpInStats.dwAddrMasks = .icmpInStats.dwAddrMasks Then _
ListView4.ListItems(12).SubItems(1) = .icmpInStats.dwAddrMasks
If Not icmp2.icmpInStats.dwAddrMaskReps = .icmpInStats.dwAddrMaskReps Then _
ListView4.ListItems(13).SubItems(1) = .icmpInStats.dwAddrMaskReps
End With

icmp2 = icmp

End Sub

Private Sub imgSelected_Click(Index As Integer)
Dim x As Byte

'Clear the radio buttons...
For x = 0 To 4
    imgSelected(x).Picture = Image6.Picture
Next x

'Update the radio buttons...
imgSelected(Index).Picture = Image9.Picture

'Remember the control selected...
Select Case Index
    Case 0
        sControlSelected = "tcp"
    Case 1
        sControlSelected = "ip"
    Case 2
        sControlSelected = "udp"
    Case 3
        sControlSelected = "in"
    Case 4
        sControlSelected = "out"
End Select

If sControlSelected = "tcp" Then
ListView1.Visible = True
ListView2.Visible = False
ListView3.Visible = False
ListView4.Visible = False
ListView5.Visible = False
End If

If sControlSelected = "ip" Then
ListView1.Visible = False
ListView2.Visible = True
ListView3.Visible = False
ListView4.Visible = False
ListView5.Visible = False
End If

If sControlSelected = "udp" Then
ListView1.Visible = False
ListView2.Visible = False
ListView3.Visible = True
ListView4.Visible = False
ListView5.Visible = False
End If

If sControlSelected = "in" Then
ListView1.Visible = False
ListView2.Visible = False
ListView3.Visible = False
ListView4.Visible = True
ListView5.Visible = False
End If

If sControlSelected = "out" Then
ListView1.Visible = False
ListView2.Visible = False
ListView3.Visible = False
ListView4.Visible = False
ListView5.Visible = True
End If
End Sub

Private Sub Label1_Click(Index As Integer)
Dim x As Byte

'Clear the radio buttons...
For x = 0 To 4
    imgSelected(x).Picture = Image6.Picture
Next x

'Update the radio buttons...
imgSelected(Index).Picture = Image9.Picture

'Remember the control selected...
Select Case Index
    Case 0
        sControlSelected = "tcp"
    Case 1
        sControlSelected = "ip"
    Case 2
        sControlSelected = "udp"
    Case 3
        sControlSelected = "in"
    Case 4
        sControlSelected = "out"
End Select

If sControlSelected = "tcp" Then
ListView1.Visible = True
ListView2.Visible = False
ListView3.Visible = False
ListView4.Visible = False
ListView5.Visible = False
End If

If sControlSelected = "ip" Then
ListView1.Visible = False
ListView2.Visible = True
ListView3.Visible = False
ListView4.Visible = False
ListView5.Visible = False
End If

If sControlSelected = "udp" Then
ListView1.Visible = False
ListView2.Visible = False
ListView3.Visible = True
ListView4.Visible = False
ListView5.Visible = False
End If

If sControlSelected = "in" Then
ListView1.Visible = False
ListView2.Visible = False
ListView3.Visible = False
ListView4.Visible = True
ListView5.Visible = False
End If

If sControlSelected = "out" Then
ListView1.Visible = False
ListView2.Visible = False
ListView3.Visible = False
ListView4.Visible = False
ListView5.Visible = True
End If
End Sub

Private Sub Timer1_Timer()

    UpdateStats1

End Sub

Private Sub Timer2_Timer()
UpdateStats2
End Sub

Private Sub Timer3_Timer()
UpdateStats3
End Sub

Private Sub Timer4_Timer()
UpdateStats4
End Sub

Private Sub Timer5_Timer()
UpdateStats5
End Sub
