Attribute VB_Name = "modIP"
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128


Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal Hostname As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long


Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSADATA
    Dim lngRetVal As Integer
    Dim strLowByte As String
    Dim strHighByte As String
    Dim strMsg As String

    lngRetVal = WSAStartup(WS_VERSION_REQD, WSAD)

    If lngRetVal <> 0 Then
        'MsgBox "Winsock.dll is not responding."
        'UpdateStatus "Winsock.dll is not responding."
        SocketsInitialize = False
        'End
    End If

    If LoByte(WSAD.wversion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wversion) = _
        WS_VERSION_MAJOR And HiByte(WSAD.wversion) < WS_VERSION_MINOR) Then

        strHighByte = Trim(Str(HiByte(WSAD.wversion)))
        strLowByte = Trim(Str(LoByte(WSAD.wversion)))
        strMsg = "Windows Sockets version " & strLowByte & "." & strHighByte
        strMsg = strMsg & " is not supported by winsock.dll "
        'UpdateStatus strMsg
        SocketsInitialize = False
    End If

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        strMsg = "This application requires a minimum of "
        strMsg = strMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        'UpdateStatus strMsg
        SocketsInitialize = False
    End If
    
    SocketsInitialize = True
End Function

Sub SocketsCleanup()
    Dim lngRetVal As Long

    lngRetVal = WSACleanup()

    If lngRetVal <> 0 Then
        'MsgBox "Socket error " & Trim(Str(lngRetVal)) & " occurred in Cleanup "
        'UpdateStatus "Socket error " & Trim(Str(lngRetVal)) & " occurred in Cleanup "
        'End
    End If
End Sub

Public Function ResolveHostname(Hostname As String) As String
    Dim strHostName As String * 256
    Dim lngHosetnAddr As Long
    Dim udtHOST As HOSTENT
    Dim lngHostIP As Long
    Dim arrTempIPAddr() As Byte
    Dim i As Integer
    Dim strIPAddress As String
    
    strHostName = Hostname
   
    If Len(strHostName) = 0 Then Exit Function
    
    lngHosetnAddr = gethostbyname(Trim(strHostName))
    If lngHosetnAddr = 0 Then
        'UpdateStatus "Error getting IP. Could not resolve " & Hostname
        Exit Function
    End If

    'Exit Function
    
    RtlMoveMemory udtHOST, lngHosetnAddr, LenB(udtHOST)
    RtlMoveMemory lngHostIP, udtHOST.hAddrList, 4

    'get all of the IP addresses if machine is multi-homed
    Do
        ReDim arrTempIPAddr(1 To udtHOST.hLength)
        RtlMoveMemory arrTempIPAddr(1), lngHostIP, udtHOST.hLength

        For i = 1 To udtHOST.hLength
            strIPAddress = strIPAddress & arrTempIPAddr(i) & "."
        Next
        strIPAddress = Mid(strIPAddress, 1, Len(strIPAddress) - 1)
        
        'UpdateStatus "Resolved " & Hostname & " to " & strIPAddress
        
        ResolveHostname = strIPAddress

        strIPAddress = ""
        udtHOST.hAddrList = udtHOST.hAddrList + LenB(udtHOST.hAddrList)
        RtlMoveMemory lngHostIP, udtHOST.hAddrList, 4
    Loop While (lngHostIP <> 0)
End Function

Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function
