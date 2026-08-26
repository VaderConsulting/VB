Attribute VB_Name = "modMain"
Option Explicit

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
    WVersion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Public Type OSInfo
    Name As String
    Workstation As Boolean
    Server As Boolean
    Version As String
    ServicePack As String
    Primary As Boolean
End Type

Public Type Computer
    OS As OSInfo
    Name As String
End Type

Public Host As Computer

Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long

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
    rc = WNetAddConnection2(lpNetResource, lpPassword, lpUserName, 0)
    Connect2 = rc
    If rc <> 0 Then GoTo Err_Connect
    Exit Function
Err_Connect:
    'Errornum = rc
    'ErrorMsg = WnetError(rc)
    'Connect2 = rc
End Function

Public Function DisConnect2(ByVal Name As String, ByVal ForceOff As Boolean) As Long
    Dim rc As Long
    On Error GoTo Err_DisConnect
    rc = WNetCancelConnection2(Name & Chr(0), CONNECT_UPDATE_PROFILE, ForceOff)
    DisConnect2 = rc
    If rc <> 0 Then GoTo Err_DisConnect
    Exit Function
Err_DisConnect:
    'Errornum = rc
    'ErrorMsg = WnetError(rc)
    'DisConnect2 = rc
End Function

Sub main()
    frmMain.Show
End Sub
