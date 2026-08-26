Attribute VB_Name = "modMain"
Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long

Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Public Const NO_ERROR = 0
Public Const CONNECT_UPDATE_PROFILE = &H1
' The following includes all the constants defined for NETRESOURCE,
' not just the ones used in this application.
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCE_CONNECTED = &H1
Public Const RESOURCE_REMEMBERED = &H3
Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEUSAGE_CONNECTABLE = &H1
Public Const RESOURCEUSAGE_CONTAINER = &H2
' Error Constants:
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_BAD_DEV_TYPE = 66&
Public Const ERROR_BAD_DEVICE = 1200&
Public Const ERROR_BAD_NET_NAME = 67&
Public Const ERROR_BAD_PROFILE = 1206&
Public Const ERROR_BAD_PROVIDER = 1204&
Public Const ERROR_BUSY = 170&
Public Const ERROR_CANCELLED = 1223&
Public Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Public Const ERROR_EXTENDED_ERROR = 1208&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&

Public LogDir As String

Sub main()
    frmMain.Show
End Sub

