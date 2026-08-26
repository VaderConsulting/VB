Attribute VB_Name = "modTracert"
Option Explicit
Public strTrace As String
Private Const WSADescription_Len As Long = 255  '256, 0-based
Private Const WSASYS_Status_Len As Long = 127   '128, 0-based
Private Const WS_VERSION_REQD As Long = &H101
Private Const SOCKET_ERROR As Long = -1
Private Const AF_INET As Long = 2
Private Const IP_SUCCESS As Long = 0
Private Const MIN_SOCKETS_REQD As Long = 1
Public Const EM_SETTABSTOPS As Long = &HCB

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADescription_Len) As Byte
   szSystemStatus(0 To WSASYS_Status_Len) As Byte
   imaxsockets As Integer
   imaxudp As Integer
   lpszvenderinfo As Long
End Type

Public Type ICMP_OPTIONS
   ttl             As Byte         'Time To Live
   Tos             As Byte         'Timeout
   Flags           As Byte         'option flags
   OptionsSize     As Long         '
   OptionsData     As Long         '
End Type

Public Type ICMP_ECHO_REPLY
   Address         As Long         'replying address
   Status          As Long         'reply status code
   RoundTripTime   As Long         'round-trip time, in milliseconds
   datasize        As Integer      'reply data size. Always an Int.
   Reserved        As Integer      'reserved for future use
   DataPointer     As Long         'pointer to the data in Data below
   Options         As ICMP_OPTIONS 'reply options, used in tracert
   ReturnedData    As String * 256 'the returned data follows the
                                    'reply message. The data string
                                    'must be sufficiently large enough
                                    'to hold the returned data.
End Type

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Long
Private Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal nbytes As Long)
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
    
Public Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    RequestOptions As ICMP_OPTIONS, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
    

Public Function GetIPFromHostName(ByVal sHostName As String) As String

  'converts a host name to an IP address.

   Dim ptrHosent As Long      'address of hostent structure
   Dim ptrName As Long        'address of name pointer
   Dim ptrAddress As Long     'address of address pointer
   Dim ptrIPAddress As Long   'address of string holding final IP address
   Dim dwAddress As Long      'the final IP address
   
   ptrHosent = gethostbyname(sHostName & vbNullChar)

   If ptrHosent <> 0 Then

     'assign pointer addresses and offset
     
     'ptrName is the official name of the host (PC).
     'If using the DNS or similar resolution system,
     'it is the Fully Qualified Domain Name (FQDN)
     'that caused the server to return a reply.
     'If using a local hosts file, it is the first
     'entry after the IP address.
      ptrName = ptrHosent
      
     'Null-terminated list of addresses for the host.
     'The Address is offset 12 bytes from the start of
     'the HOSENT structure. Addresses are returned
     'in network byte order.
      ptrAddress = ptrHosent + 12
      
     'get the actual IP address
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory dwAddress, ByVal ptrIPAddress, 4

      GetIPFromHostName = GetIPFromAddress(dwAddress)

   End If
   
End Function


Public Sub SocketsCleanup()
   
  'only show error if unable to clean up the sockets
   If WSACleanup() <> 0 Then
       Debug.Print "Windows Sockets error occurred during Cleanup.", vbExclamation
   End If
    
End Sub


Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   
  'when the socket version returned == version
  'required, return True
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function


Public Function GetIPFromAddress(Address As Long) As String
   
   Dim ptrString As Long
   
   ptrString = inet_ntoa(Address)
   GetIPFromAddress = GetStrFromPtrA(ptrString)
   
End Function


Public Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function


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
    
    strTrace = ""
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
                'strTrace = "Tracing Route to " + strIPAddress + ":" & vbCrLf
                
                'The heart of the call
                For ttl = 1 To 30 ' was 255 ' Number of Hops
                    
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
                    'Debug.Print "Current TTL: " & ttl
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
                    
                    'ttl = ttl - ttlAdjust ' <--- ?????????????????????????
                    'need some processing time
                    DoEvents
                    'Debug.Print "*****START*****"
                    'Debug.Print strTrace
                    'Debug.Print "******END****"
                    
                    If sHostIP = strIPAddress Or ttlAdjust = 1 Then
                        
                        'we're done
                        'strTrace = strTrace & vbCrLf + "Trace Route Complete"
                        TraceRT = strTrace
                        Exit For
                        
                    End If
                    
                Next ttl
                
                'clean up
                Call IcmpCloseHandle(hPort)
            
            Else
                Debug.Print "Unable to Open an Icmp File Handle", vbOKOnly, "RTM Server"
            End If  'If hPort
            
            'clean up
            Call SocketsCleanup
        Else
            Debug.Print "Unable to resolve hostname"
        End If
    Else
        Debug.Print "Unable to initialize the Windows Sockets", vbOKOnly, "RTM Server"
    End If  'if SocketsInitialize()
    
End Function


Private Sub ShowResults(timeToLive As Byte, tripTime As Long, sHostIP As String)
   
   Dim sTripTime As String
   Dim buff As String
   Dim tmp As String

  'format a string representing
  'the round trip time
    Select Case tripTime
        Case Is < 10
            sTripTime = "<10 ms"
        Case Is > 3000 ' was 1200
            sTripTime = "*"
        Case Else
            sTripTime = CStr(tripTime) & " ms"
    End Select
   
  'cache the textbox
   buff = strTrace
   
  'create a new entry
   tmp = "Hop #" & vbTab & _
          CStr(timeToLive) & vbTab & _
          sTripTime & vbTab & _
          sHostIP & vbCrLf

  'update textbox
   'strTrace = buff & tmp
   strTrace = buff & "," & sHostIP
    
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
                   6000) = 1 Then ' was 2400
   
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




