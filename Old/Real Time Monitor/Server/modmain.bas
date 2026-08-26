Attribute VB_Name = "modMain"
Public strTrace As String

Public Function TraceRT(strIPAddress As String) As String
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
                strTrace = "Tracing Route to " + strIPAddress + ":" & vbCrLf & vbCrLf
                
                'The heart of the call
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
                        strTrace = strTrace & vbCrLf + "Trace Route Complete"
                        TraceRT = strTrace
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


Private Sub ShowResults(timeToLive As Byte, tripTime As Long, sHostIP As String)
   
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
   buff = strTrace
   
  'create a new entry
   tmp = "Hop #" & vbTab & _
          CStr(timeToLive) & vbTab & _
          sTripTime & vbTab & _
          sHostIP & vbCrLf

  'update textbox
   strTrace = buff & tmp
    
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


