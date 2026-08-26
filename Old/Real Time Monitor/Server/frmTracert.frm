VERSION 5.00
Begin VB.Form frmTracert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tracert"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   2445
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "32"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Resolve IP Addresses"
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Trace"
      Default         =   -1  'True
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Chars per Packet"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "No of Packets"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Target IP"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmTracert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_Load()

    With Combo1
      .AddItem "www.tusk.com.au"
      .AddItem "www.gov.on.ca"
      .AddItem "www.microsoft.com"
      .AddItem "www.yahoo.com"
      .ListIndex = 1
   End With

   Text1.Text = ""
   Text4.Text = ""
   
   ReDim TabArray(0 To 3) As Long
   
   TabArray(0) = 30
   TabArray(1) = 54
   TabArray(2) = 105
   TabArray(3) = 182
   
  'Clear existing tabs
  'and set the text tabstops
   Call SendMessage(Text4.hwnd, EM_SETTABSTOPS, 0&, ByVal 0&)
   Call SendMessage(Text4.hwnd, EM_SETTABSTOPS, 4&, TabArray(0))
   Text4.Refresh

End Sub


Public Sub Command1_Click()
   
   Command1.Enabled = False
   'Call TraceRT
   Command1.Enabled = True
   
End Sub


Public Function TraceRT(strIPAddress As String, strRoute As String)
    Dim ipo As ICMP_OPTIONS
    Dim echo As ICMP_ECHO_REPLY
    Dim ttl As Integer
    Dim ttlAdjust As Integer
    Dim hPort As Long
    Dim nChrsPerPacket As Long
    Dim dwAddress As Long
    Dim sAddress As String
    Dim sHostIP As String
    
    
    'set up
    'Text1.Text = ""  'the target IP
    'Text2.Text = "1" 'force the no of packets = 1 for a tracert
    'Text4.Text = ""  'clear the output window
    'List1.Clear      'for info/debuging only
    
    nChrsPerPacket = 32
    
    If SocketsInitialize() Then
    
        'returns the IP Address for the Host in Combo 1
        'ie returns 209.68.48.118 for www.mvps.org
        
        'sAddress = GetIPFromHostName(strHostname)
        sAddress = strIPAddress
        
        If sAddress <> "" Then
            'convert the address into an internet address.
            'ie returns 1982874833 when passed 209.68.48.118
            dwAddress = inet_addr(sAddress)
            
            'open an internet file handle
            hPort = IcmpCreateFile()
            
            If hPort <> 0 Then
            
                'update the textboxes
                'Text1.Text = sAddress
                strRoute = "Tracing Route to " + strIPAddress + ":" & vbCrLf & vbCrLf
                
                'The heart of the call. See the VBnet
                'page description of the TraceRt TTL
                'member and its use in performing a
                'Trace Route.
                'frmTracert.Refresh
                For ttl = 1 To 255
                    
                    '--------------------------------
                    'for demo/dedbugging only. The
                    'list will show each TTL passed
                    'to the calls. Duplicate TTL's
                    'mean the request timed out, and
                    'additional attempts to obtain
                    'the route were tried.
                    'List1.AddItem ttl
                    '--------------------------------
                    
                    'set the IPO time to live
                    'value to the current hop
                    ipo.ttl = ttl
                    
                    'Call the API.
                    '
                    'Two items of consequence happen here.
                    'First, the return value of the call is
                    'assigned to an 'adjustment' variable. If
                    'the call was successful, the adjustment
                    'is 0, and the Next will increment the TTL
                    'to obtain the next hop. If the return value
                    'is 1, 1 is subtacted from the TTL value, so
                    'when the next increments the TTL counter it
                    'will be the same value as the last pass. In
                    'doing this, routers that time out are retried
                    'to ensure a completed route is determined.
                    '(The values in the List1 show the actual
                    ' hops/tries that the method made.)
                    'i.e. if the TTL = 3 and it times out,
                    '     adjust = 1 so ttl - 1 = 2. On
                    '     encountering the Next, TTL is
                    '     reset to 3 and the route is tried again.
                    
                    'The second thing happening concerns the
                    'sHostIP member of the call. When the call
                    'returns, sHostIP will contain the name
                    'of the traced host IP.  If it matches the
                    'string initially used to create the address
                    '(above) were at the target, so end.
                    ttlAdjust = TraceRTSendEcho(hPort, _
                                                        dwAddress, _
                                                        nChrsPerPacket, _
                                                        sHostIP, _
                                                        echo, _
                                                        ipo)
                    
                    ttl = ttl - ttlAdjust
                    'need some processing time
                    DoEvents
                    
                    If sHostIP = strIPAddress Then
                        
                        'we're done
                        strRoute = strRoute & vbCrLf + "Trace Route Complete"
                        Exit For
                        
                    End If
                    
                Next ttl
                
                'clean up
                Call IcmpCloseHandle(hPort)
            
            Else
                'MsgBox "Unable to Open an Icmp File Handle", vbOKOnly, "RTM Server"
            End If  'If hPort
            
            'clean up
            Call SocketsCleanup
        Else
            'MsgBox "Unable to resolve hostname"
        End If
    Else
        'MsgBox "Unable to initialize the Windows Sockets", vbOKOnly, "RTM Server"
    End If  'if SocketsInitialize()
    
End Function


Private Sub ShowResults(timeToLive As Byte, _
                        tripTime As Long, _
                        sHostIP As String)
   
   Dim sTripTime As String
   Dim buff As String
   Dim tmp As String

  'format a string representing
  'the round trip time
   Select Case tripTime
      Case Is < 10:   sTripTime = "<10 ms"
      Case Is > 1200: sTripTime = "*"
      Case Else:      sTripTime = CStr(tripTime) & " ms"
   End Select
   
  'cache the textbox
   buff = Text4.Text
   
  'create a new entry
   tmp = "Hop #" & vbTab & _
          CStr(timeToLive) & vbTab & _
          sTripTime & vbTab & _
          sHostIP & vbCrLf

  'update textbox
   Text4.Text = buff & tmp
    
End Sub


Private Function TraceRTSendEcho(hPort As Long, _
                                 dwAddress As Long, _
                                 nChrsPerPacket As Long, _
                                 sHostIP As String, _
                                 echo As ICMP_ECHO_REPLY, _
                                 ipo As ICMP_OPTIONS) As Integer

   Dim sData As String
   Dim sError As String
   Dim sHostName As String
   Dim ttl As Integer
   
  'create a buffer to send
   sData = String$(nChrsPerPacket, "a")
                   
   If IcmpSendEcho(hPort, _
                   dwAddress, _
                   sData, _
                   Len(sData), _
                   ipo, _
                   echo, _
                   Len(echo) + 8, _
                   2400) = 1 Then
   
      'a reply was received, so update the display
       sHostIP = GetIPFromAddress(echo.Address)
              
       ShowResults ipo.ttl, echo.RoundTripTime, sHostIP
       
      'return 0 to continue with retrieval
       TraceRTSendEcho = 0
      
   Else
      
      'a timeout was received, so set the
      'return value to 1. In the TraceRT
      'calling routine, the TTL will be
      'de-incremented by 1, causing the
      'for / next to retry this hop.
       TraceRTSendEcho = 1
   
   End If
        
End Function

