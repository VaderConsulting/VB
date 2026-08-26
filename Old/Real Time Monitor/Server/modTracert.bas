Attribute VB_Name = "modTracertAPI"
Option Explicit
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
       'MsgBox "Windows Sockets error occurred during Cleanup.", vbExclamation
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


