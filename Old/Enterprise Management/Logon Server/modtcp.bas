Attribute VB_Name = "modTCP"

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

' Declare API functions
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal Hostname As String, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal Hostname As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Function IPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim HOST As HOSTENT
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim i As Integer
    Dim sIPAddr As String
    
    If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        IPAddress = ""
        'MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)
    
    If lpHost = 0 Then
        IPAddress = ""
        'MsgBox "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    RtlMoveMemory HOST, lpHost, Len(HOST)
    RtlMoveMemory dwIPAddr, HOST.hAddrList, 4
    
    ReDim tmpIPAddr(1 To HOST.hLength)
    
    RtlMoveMemory tmpIPAddr(1), dwIPAddr, HOST.hLength
    
    For i = 1 To HOST.hLength
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    
    IPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
    
End Function

Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

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

Private Function WnetError(Errcode As Long) As String
    ' Network error code handling
    Select Case Errcode
        Case ERROR_NO_ERROR
            WnetError = "Success."
        Case WN_Not_Supported
           WnetError = "Function is not supported."
        Case WN_Out_Of_Memory
           WnetError = "Out of Memory."
        Case WN_Net_Error
           WnetError = "An error occurred on the network."
        Case WN_Bad_Pointer
           WnetError = "The Pointer was Invalid."
        Case WN_Bad_NetName
           WnetError = "Invalid Network Resource Name."
        Case WN_Bad_Password
           WnetError = "The Password was Invalid."
        Case WN_Bad_Localname
           WnetError = "The local device name was invalid."
        Case WN_Access_Denied
           WnetError = "A security violation occurred."
        Case ERROR_ACCESS_DENIED
           WnetError = "A security violation occurred."
        Case WN_Already_Connected
           WnetError = "The local device was connected to a remote resource."
        Case WN_Server_Down
           WnetError = "Cannot resolve server name or the server you wish to connect to is down."
        Case WN_Server_Down_2
           WnetError = "Cannot resolve server name or the server you wish to connect to is down.*"
        Case WN_NOT_FOUND
           WnetError = "The network name cannot be found."
        Case WN_IN_USE
            WnetError = "The Local Device name is already in use."
        Case WN_NOT_CONNECTED
            WnetError = "The Local Device name is not connected."
        Case ERROR_OPEN_FILES
            WnetError = "One or more files are in use on that connection."
        Case ERROR_ALREADY_ASSIGNED
            WnetError = "The Local Device name is already in use."
        Case ERROR_BAD_DEVICE
            WnetError = "Bad Device Error."
        Case ERROR_SESSION_CREDENTIAL_CONFLICT
            WnetError = "The credentials supplied conflict with an existing set of credentials"
        Case Else:
           WnetError = "Unrecognized Error " + Str(Errcode) + "."
      End Select
End Function

