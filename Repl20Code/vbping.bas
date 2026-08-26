Attribute VB_Name = "vbPing"
Private Const IP_STATUS_BASE               As Long = 11000
Private Const IP_SUCCESS                   As Long = 0
Private Const IP_BUF_TOO_SMALL             As Long = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE      As Long = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE     As Long = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE     As Long = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE     As Long = (11000 + 5)
Private Const IP_NO_RESOURCES              As Long = (11000 + 6)
Private Const IP_BAD_OPTION                As Long = (11000 + 7)
Private Const IP_HW_ERROR                  As Long = (11000 + 8)
Private Const IP_PACKET_TOO_BIG            As Long = (11000 + 9)
Private Const IP_REQ_TIMED_OUT             As Long = (11000 + 10)
Private Const IP_BAD_REQ                   As Long = (11000 + 11)
Private Const IP_BAD_ROUTE                 As Long = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT       As Long = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM       As Long = (11000 + 14)
Private Const IP_PARAM_PROBLEM             As Long = (11000 + 15)
Private Const IP_SOURCE_QUENCH             As Long = (11000 + 16)
Private Const IP_OPTION_TOO_BIG            As Long = (11000 + 17)
Private Const IP_BAD_DESTINATION           As Long = (11000 + 18)
Private Const IP_ADDR_DELETED              As Long = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE           As Long = (11000 + 20)
Private Const IP_MTU_CHANGE                As Long = (11000 + 21)
Private Const IP_UNLOAD                    As Long = (11000 + 22)
Private Const IP_ADDR_ADDED                As Long = (11000 + 23)
Private Const IP_GENERAL_FAILURE           As Long = (11000 + 50)
Private Const MAX_IP_STATUS                As Long = (11000 + 50)
Private Const IP_PENDING                   As Long = (11000 + 255)
Private Const PING_TIMEOUT                 As Long = 500
Private Const WS_VERSION_REQD              As Long = &H101
Private Const MIN_SOCKETS_REQD             As Long = 1
Private Const SOCKET_ERROR                 As Long = -1
Private Const INADDR_NONE                  As Long = &HFFFFFFFF
Private Const MAX_WSADescription           As Long = 256
Private Const MAX_WSASYSStatus             As Long = 128

Private Type WSADATA
    wVersion                               As Integer
    wHighVersion                           As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)  As Byte
    wMaxSockets                            As Long
    wMaxUDPDG                              As Long
    dwVendorInfo                           As Long
End Type

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLen      As Integer
    hAddrList As Long
End Type

Private Declare Function GetHostByName Lib "WSOCK32.DLL" (ByVal hostname As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal s As String) As Long

' ------------------------------------------------------------
' Perform ICMP Ping to specified hostname
' Inputs
'      sAddress (String)
'      sDataToSend (String)
'      ECHO (ICMP_ECHO_REPLY)
' Outputs
'      Ping (Long)
Public Function Ping(sAddress As String, sDataToSend As String, ECHO As ICMP_ECHO_REPLY) As Long
    Dim hPort As Long
    Dim dwAddress As Long
       
    dwAddress = inet_addr(sAddress)
    If dwAddress <> INADDR_NONE Then
        hPort = IcmpCreateFile()
        If hPort Then
            Call IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO, Len(ECHO), PING_TIMEOUT)
            Ping = ECHO.Status
            Call IcmpCloseHandle(hPort)
        End If
    Else
        Ping = INADDR_NONE
    End If
End Function

' ------------------------------------------------------------
' Retrieve result code from the last Network operation
' Inputs
'      Status (Long)
' Outputs
'     GetStatusCode (String)
Public Function GetStatusCode(Status As Long) As String
    Dim msg As String
    
    Select Case Status
        Case IP_SUCCESS:               msg = "ip success"
        Case INADDR_NONE:              msg = "inet_addr: bad IP format"
        Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
        Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
        Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
        Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
        Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
        Case IP_NO_RESOURCES:          msg = "ip no resources"
        Case IP_BAD_OPTION:            msg = "ip bad option"
        Case IP_HW_ERROR:              msg = "ip hw_error"
        Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
        Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
        Case IP_BAD_REQ:               msg = "ip bad req"
        Case IP_BAD_ROUTE:             msg = "ip bad route"
        Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
        Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
        Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
        Case IP_SOURCE_QUENCH:         msg = "ip source quench"
        Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
        Case IP_BAD_DESTINATION:       msg = "ip bad destination"
        Case IP_ADDR_DELETED:          msg = "ip addr deleted"
        Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
        Case IP_MTU_CHANGE:            msg = "ip mtu_change"
        Case IP_UNLOAD:                msg = "ip unload"
        Case IP_ADDR_ADDED:            msg = "ip addr added"
        Case IP_GENERAL_FAILURE:       msg = "ip general failure"
        Case IP_PENDING:               msg = "ip pending"
        Case PING_TIMEOUT:             msg = "ping timeout"
        Case Else:                     msg = "unknown  msg returned"
    End Select
   
   GetStatusCode = CStr(Status) & "   [ " & msg & " ]"
   
End Function

' ------------------------------------------------------------
' Retrieve IP Address from a given Hostname
' Inputs
'      sHostname (String)
' Outputs
'      GetIPFromHostname (String)
Public Function GetIPFromHostName(ByVal sHostName As String) As String
    
    Dim nbytes As Long
    Dim ptrHosent As Long
    Dim ptrName As Long
    Dim ptrAddress As Long
    Dim ptrIPAddress As Long
    Dim sAddress As String
   
    sAddress = Space$(4)
    ptrHosent = GetHostByName(sHostName & vbNullChar)
    If ptrHosent <> 0 Then
        ptrName = ptrHosent
        ptrAddress = ptrHosent + 12
        CopyMemory ptrName, ByVal ptrName, 4
        CopyMemory ptrAddress, ByVal ptrAddress, 4
        CopyMemory ptrIPAddress, ByVal ptrAddress, 4
        CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4
        GetIPFromHostName = IPToText(sAddress)
    End If
End Function

' ------------------------------------------------------------
' Convert Hexadecimal IPAddress to dotted string
' Inputs
'      IPAddress (String)
' Outputs
'      IPToText (String)
Private Function IPToText(ByVal IPAddress As String) As String
    IPToText = CStr(Asc(IPAddress)) & "." & CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & CStr(Asc(Mid$(IPAddress, 4, 1)))
End Function

' ------------------------------------------------------------
' Close Windows Sockets
' Inputs
' Outputs
Public Sub SocketsCleanup()
    If WSACleanup() <> 0 Then
        MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
    End If
End Sub

' ------------------------------------------------------------
' Open Windows Sockets
' Inputs
' Outputs
'      SocketsInitialize (Boolean)
Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSADATA
       
    SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
End Function
