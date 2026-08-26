Attribute VB_Name = "modAPI"
Option Explicit

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

Private Type NETCONNECTINFOSTRUCT
    cbStructure As Long
    dwFlags As Long
    dwSpeed As Long
    dwDelay As Long
    dwOptdataSize As Long
End Type

Dim lpNetResource As NETRESOURCE

Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : Connect2
' DateTime  : 01-04-2003 11:14
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date    Author      History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function Connect2(ByVal Localname As String, ByVal RemoteName As String, ByVal Username As String, ByVal Password As String) As Long
    Dim lpUserName As String
    Dim lpPassword As String
    Dim ErrorNum As Long
    Dim ErrorMsg As String
    Dim rc As Long

    On Error GoTo Err_Connect
    ErrorNum = 0
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
    'DebugMsg "Connecting drive returned error " & rc & " (" & WnetError(rc) & ")"
    If rc <> 0 Then GoTo Err_Connect
    Exit Function
Err_Connect:
    ErrorNum = rc
    ErrorMsg = WnetError(rc)
    Connect2 = rc
    'DebugMsg "Connecting drive returned error " & rc & " (" & WnetError(rc) & ")"
End Function

'---------------------------------------------------------------------------------------
' Procedure : DisConnect2
' DateTime  : 01-04-2003 11:14
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date    Author      History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function DisConnect2(ByVal Name As String, ByVal ForceOff As Boolean) As Long
    Dim ErrorNum As Long, ErrorMsg As String, rc As Long
    On Error GoTo Err_DisConnect
    ErrorNum = 0
    ErrorMsg = "SUCCESS"
    rc = WNetCancelConnection2(Name & Chr(0), CONNECT_UPDATE_PROFILE, ForceOff)
    DisConnect2 = rc
    If rc <> 0 Then GoTo Err_DisConnect
    Exit Function
Err_DisConnect:
    ErrorNum = rc
    ErrorMsg = WnetError(rc)
    DisConnect2 = rc
End Function

'---------------------------------------------------------------------------------------
' Procedure : WnetError
' DateTime  : 01-04-2003 11:14
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date    Author      History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
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

