Attribute VB_Name = "modMain"
Option Explicit

' For remainder of server types, look up SetServiceBits in MSDN
Public Const SV_TYPE_WORKSTATION = &H1
Public Const SV_TYPE_SERVER = &H2
Public Const SV_TYPE_SQLSERVER = &H4
Public Const SV_TYPE_DOMAIN_CTRL = &H8
Public Const SV_TYPE_DOMAIN_BAKCTRL = &H10
Public Const SV_TYPE_TIMESOURCE = &H20
Public Const SV_TYPE_AFP = &H40
Public Const SV_TYPE_NOVELL = &H80
Public Const SV_TYPE_DOMAIN_MEMBER = &H100
Public Const SV_TYPE_LOCAL_LIST_ONLY = &H40000000
Public Const SV_TYPE_PRINT = &H200
Public Const SV_TYPE_DIALIN = &H400
Public Const SV_TYPE_XENIX_SERVER = &H800
Public Const SV_TYPE_MFPN = &H4000
Public Const SV_TYPE_NT = &H1000
Public Const SV_TYPE_WFW = &H2000
Public Const SV_TYPE_SERVER_NT = &H8000
Public Const SV_TYPE_POTENTIAL_BROWSER = &H10000
Public Const SV_TYPE_BACKUP_BROWSER = &H20000
Public Const SV_TYPE_MASTER_BROWSER = &H40000
Public Const SV_TYPE_DOMAIN_MASTER = &H80000
Public Const SV_TYPE_DOMAIN_ENUM = &H80000000
Public Const SV_TYPE_WINDOWS = &H400000
Public Const SV_TYPE_ALL = &HFFFFFFFF


Const SIZE_SI_101 = 24

Const WN_Success = &H0
Const WN_Not_Supported = &H1
Const WN_Net_Error = &H2
Const WN_Bad_Pointer = &H4
Const WN_Bad_NetName = &H32
Const WN_Bad_Password = &H6
Const WN_Bad_Localname = &H33
Const WN_Access_Denied = &H7
Const WN_Out_Of_Memory = &HB
Const WN_Already_Connected = &H34
Const WN_Server_Down = &H35
Const WN_Server_Down_2 = &H52E
Const WN_NOT_FOUND = &H43
Const WN_IN_USE = 85
Const WN_NOT_CONNECTED = 2250

'Consts for return codes errors
Const ERROR_SUCCESS = 0
Const ERROR_NO_ERROR = 0&
Const ERROR_ALREADY_ASSIGNED = 85&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_BAD_DEVICE_TYPE = 66&
Const ERROR_BAD_NET_NAME = 67&
Const ERROR_BAD_PROFILE = 1206&
Const ERROR_BAD_DEVICE = 1200&
Const ERROR_BAD_PROVIDER = 1204&
Const ERROR_BUSY = 170&
Const ERROR_CANCEL_VIOLATION = 173&
Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Const ERROR_EXTENDED_ERROR = 1208&
Const ERROR_INVALID_PASSWORD = 86&
Const ERROR_NO_NET_OR_BAD_PATH = 1203&
Const ERROR_SESSION_CREDENTIAL_CONFLICT = 1219&
Const ERROR_NO_NETWORK = 1222&
Const ERROR_CANCELLED = 1223&
Const ERROR_NO_CONNECTION = 8
Const ERROR_NO_DISCONNECT = 9
Const ERROR_DEVICE_IN_USE = 2404&
Const ERROR_NOT_CONNECTED = 2250&
Const ERROR_OPEN_FILES = 2401&
Const ERROR_MORE_DATA = 234

Const CONNECT_UPDATE_PROFILE = &H1
Const RESOURCETYPE_DISK = &H1
Const RESOURCETYPE_PRINT = &H2
Const RESOURCETYPE_ANY = &H0
Const RESOURCE_CONNECTED = &H1
Const RESOURCE_REMEMBERED = &H3
Const RESOURCE_GLOBALNET = &H2
Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Const RESOURCEDISPLAYTYPE_SERVER = &H2
Const RESOURCEDISPLAYTYPE_SHARE = &H3
Const RESOURCEUSAGE_CONNECTABLE = &H1
Const RESOURCEUSAGE_CONTAINER = &H2

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

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
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

Private Type SERVER_INFO_101
    dwPlatformId As Long
    lpszServerName As Long
    dwVersionMajor As Long
    dwVersionMinor As Long
    dwType As Long
    lpszComment As Long
End Type

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal Hostname As String) As Long
'Private Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
'Private Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
'Private Declare Function NetServerEnum Lib "Netapi32.dll" (ByVal servername As String, ByVal level As Long, buffer As Long, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal servertype As Long, ByVal domain As String, resumehandle As Long) As Long
'Private Declare Function NetServerEnum Lib "Netapi32.dll" (servername As Any, ByVal level As Long, buffer As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal servertype As Any, ByVal domain As String, resumehandle As Any) As Long
'Public Declare Function NetServerEnum Lib "Netapi32.dll" (vServername As Any, ByVal lLevel As Long, vBufptr As Any, ByVal lPrefmaxlen As Long, lEntriesRead As Long, lTotalEntries As Long, vServerType As Any, ByVal sDomain As String, vResumeHandle As Any) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32.dll" (BufPtr As Any) As Long
'Private Declare Function lstrcpyW Lib "KERNEL32" (lpszDest As Any, ByVal lpszSrc As Long) As Long
Declare Function NetGetDCName Lib "netapi32" (lpServer As Any, lpDomain As Any, vBuffer As Any) As Long
Public Declare Function NetServerEnum _
    Lib "Netapi32.dll" ( _
    vServername As Any, _
    ByVal lLevel As Long, _
    vBufptr As Any, _
    lPrefmaxlen As Long, _
    lEntriesRead As Long, _
    lTotalEntries As Long, _
    vServerType As Any, _
    ByVal sDomain As String, _
    vResumeHandle As Any) _
    As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, vSrc As Any, ByVal lSize&)
Public Declare Sub lstrcpyW Lib "kernel32" (vDest As Any, ByVal sSrc As Any)
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)

Function ResolveHostname(Hostname As String) As String
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
        Exit Function
    End If

    RtlMoveMemory udtHOST, ByVal lngHosetnAddr, LenB(udtHOST)
    RtlMoveMemory lngHostIP, ByVal udtHOST.hAddrList, 4

    'get all of the IP addresses if machine is multi-homed
    Do
        ReDim arrTempIPAddr(1 To udtHOST.hLength)
        RtlMoveMemory arrTempIPAddr(1), ByVal lngHostIP, udtHOST.hLength

        For i = 1 To udtHOST.hLength
            strIPAddress = strIPAddress & arrTempIPAddr(i) & "."
        Next
        strIPAddress = Mid(strIPAddress, 1, Len(strIPAddress) - 1)
        
        ResolveHostname = strIPAddress

        strIPAddress = ""
        udtHOST.hAddrList = udtHOST.hAddrList + LenB(udtHOST.hAddrList)
        RtlMoveMemory lngHostIP, ByVal udtHOST.hAddrList, 4
    Loop While (lngHostIP <> 0)
End Function

Public Function Connect2(ByVal Localname As String, ByVal RemoteName As String, ByVal Username As String, ByVal Password As String) As Long
    Dim lpUserName As String
    Dim lpPassword As String
    Dim Errornum As Long, ErrorMsg As String, rc As Long
    Dim lpNetResource As NETRESOURCE
    
    On Error GoTo Err_Connect
    Errornum = 0
    ErrorMsg = "SUCCESS"
    lpNetResource.dwType = RESOURCETYPE_DISK
    lpNetResource.dwScope = RESOURCE_GLOBALNET
    lpNetResource.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    lpNetResource.dwUsage = RESOURCEUSAGE_CONNECTABLE
    lpNetResource.lpLocalName = Localname
    lpNetResource.lpRemoteName = RemoteName
    'lpPassword = Chr(0) ' THESE LINES SCREW THE CODE!!! LEAVE WELL ENOUGH ALONE
    'lpUserName = Chr(0) ' VB Requires these to be NULL, but the API doesn't accept
                         ' NULL, so don't assign values at all.
    rc = WNetAddConnection2(lpNetResource, lpPassword, lpUserName, 0)
    Connect2 = rc
    If rc <> 0 Then GoTo Err_Connect
    Exit Function
Err_Connect:
    Errornum = rc
    ErrorMsg = WnetError(rc)
    Connect2 = rc
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

'Private Function PointerToString(lpszString As Long) As String
'    Dim lpszStr1 As String, lpszStr2 As String, nRes As Long
'    lpszStr1 = String(1000, "*")
'    nRes = lstrcpyW(lpszStr1, lpszString)
'    lpszStr2 = (StrConv(lpszStr1, vbFromUnicode))
'    PointerToString = Left(lpszStr2, InStr(lpszStr2, Chr$(0)) - 1)
'End Function

'Public Function GetServerInfo(pszServer As String, pszDomain As String)
'    Dim nLevel As Long, i As Long, BufPtr As Long, TempBufPtr As Long
'    Dim nPrefMaxLen As Long, nEntriesRead As Long, nTotalEntries As Long
'    Dim nServerType As Long, nResumeHandle As Long, nRes As Long
'    Dim ServerInfo As SERVER_INFO_101
'
'    pszServer = StrConv(pszServer, vbUnicode)
'    pszDomain = StrConv(pszDomain, vbUnicode)
'
'    nLevel = 101
'    BufPtr = 0
'    nPrefMaxLen = &HFFFFFFFF
'    nEntriesRead = 0
'    nTotalEntries = 0
'    nServerType = SV_TYPE_ALL
'    nServerType = 0
'    nResumeHandle = 0
'
'    Do
'        nRes = NetServerEnum(ByVal pszServer, nLevel, BufPtr, nPrefMaxLen, nEntriesRead, nTotalEntries, ByVal nServerType, pszDomain, nResumeHandle)
'        If ((nRes = ERROR_SUCCESS) Or (nRes = ERROR_MORE_DATA)) And (nEntriesRead > 0) Then
'            TempBufPtr = BufPtr
'            For i = 1 To nEntriesRead
'                RtlMoveMemory ServerInfo, TempBufPtr, SIZE_SI_101
'                Debug.Print PointerToString(ServerInfo.lpszServerName)
'                TempBufPtr = TempBufPtr + SIZE_SI_101
'            Next i
'        Else
'            'MsgBox "NetServerEnum failed: " & nRes
'        End If
'        NetApiBufferFree (BufPtr)
'    Loop While nEntriesRead < nTotalEntries
'End Function

Public Sub ListServers(lType As Long, strDomain As String, strServer As String)
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
    
    sDomain = StrConv(strDomain, vbUnicode)
    lReturn = NetServerEnum(ByVal 0&, 101, Server_Info, lMax, lEntries, lTotal, ByVal lType, sDomain, vResume)
    
    If lReturn <> 0 And lReturn <> 234 Then
        Exit Sub
    End If
    
    X = 1
    lServerInfo101StructPtr = Server_Info

    Do While X <= lTotal
        RtlMoveMemory tServer_info_101, ByVal lServerInfo101StructPtr, Len(tServer_info_101)
        lstrcpyW bBuffer(0), tServer_info_101.lpszServerName
        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & Chr$(bBuffer(i))
            i = i + 2
        Loop
        
        If sServer = strServer Then
            Debug.Print sServer ' this is a server
        End If
        DoEvents
        X = X + 1
        sServer = ""
        lServerInfo101StructPtr = lServerInfo101StructPtr + Len(tServer_info_101)
    Loop
    lReturn = NetApiBufferFree(Server_Info)
End Sub

Public Sub GetDomainList(lType As Long)

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

    frmMain.cmbDomains.Clear

    lReturn = NetServerEnum( _
        ByVal 0&, _
        101, _
        Server_Info, _
        lMax, _
        lEntries, _
        lTotal, _
        ByVal lType, _
        sDomain, _
        vResume)

    If lReturn <> 0 Then
        'StatusBar1.Panels("msg").Text = "Error " + Str$(lReturn) + _
        " when trying to obtain server list " + Str$(lTotal)
        Exit Sub
    End If

    X = 1
    lServerInfo101StructPtr = Server_Info
    Do While X <= lTotal
        RtlMoveMemory _
            tServer_info_101, _
            ByVal lServerInfo101StructPtr, _
            Len(tServer_info_101)
        lstrcpyW bBuffer(0), _
            tServer_info_101.lpszServerName
        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & _
                Chr$(bBuffer(i))
            i = i + 2
        Loop
       frmMain.cmbDomains.AddItem sServer
        DoEvents
        X = X + 1
            sServer = ""
        lServerInfo101StructPtr = _
            lServerInfo101StructPtr + _
            Len(tServer_info_101)
    Loop
    lReturn = NetApiBufferFree(Server_Info)
End Sub

'Public Function GetPDCName(strDomain() As Byte) As String
'    Dim lpBuffer As Long, nRet As Long
'    Dim yServer() As Byte
'    Dim sLocal As String
'
'    yServer = MakeServerName(strDomain())
'    nRet = NetGetDCName(yServer(0), strDomain(), lpBuffer)
'    If nRet = 0 Then
'        sLocal = PointerToStringW(lpBuffer)
'    End If
'    If lpBuffer Then Call NetApiBufferFree(lpBuffer)
'    GetPDCName = sLocal
'End Function

'Public Function MakeServerName(ServerName() As Byte)
'    Dim yServer() As Byte
'    If ServerName() <> "" Then
'        If InStr(1, ServerName, "\\") = 0 Then
'            ServerName = "\\" & ServerName
'        End If
'    End If
'    yServer = ServerName & vbNullChar
'    MakeServerName = yServer
'End Function

'Public Function PointerToStringW(lpStringW As Long) As String
'   Dim buffer() As Byte
'   Dim nLen As Long
'
'   If lpStringW Then
'      nLen = lstrlenW(lpStringW) * 2
'      If nLen Then
'         ReDim buffer(0 To (nLen - 1)) As Byte
'         CopyMem buffer(0), ByVal lpStringW, nLen
'         PointerToStringW = buffer
'      End If
'   End If
'End Function

Public Function DisConnect2(ByVal Name As String, ByVal ForceOff As Boolean) As Long
    Dim rc As Long
    On Error GoTo Err_DisConnect
    'ErrorMsg = "SUCCESS"
    rc = WNetCancelConnection2(Name & Chr(0), CONNECT_UPDATE_PROFILE, ForceOff)
    DisConnect2 = rc
    If rc <> 0 Then GoTo Err_DisConnect
    Exit Function
Err_DisConnect:
    'Errornum = rc
    'ErrorMsg = WnetError(rc)
    'DisConnect2 = rc
End Function
