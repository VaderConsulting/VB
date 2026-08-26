VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Usage:
' Dim o As clsSpeed
' Dim errCode As Long, errMessage As String
    
' Set o = CreateObject("drNetwork.clsSpeed")
    
' Debug.Print o.GetConnectionSpeed("\\192.168.1.1\downloads", "admin", "pa$$word", errCode, errMessage) / 8000000 & "Mb / Sec"
' Debug.Print errCode, errMessage
    

Option Explicit
        
Private Declare Function MultinetGetConnectionPerformance Lib "mpr.dll" Alias "MultinetGetConnectionPerformanceA" (lpResource As NETRESOURCE, lpNetConnectInfoStruct As NETCONNECTINFOSTRUCT) As Long
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUsername As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long


'Consts for return code errors
Private Const ERROR_BAD_DEV_TYPE = 66&
Private Const ERROR_ALREADY_ASSIGNED = 85&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_BAD_NET_NAME = 67&
Private Const ERROR_BAD_PROFILE = 1206&
Private Const ERROR_BAD_PROVIDER = 1204&
Private Const ERROR_BUSY = 170&
Private Const ERROR_CANCEL_VIOLATION = 173&
Private Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Private Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Private Const ERROR_EXTENDED_ERROR = 1208&
Private Const ERROR_INVALID_PASSWORD = 86&
Private Const ERROR_NO_NET_OR_BAD_PATH = 1203&
Private Const ERROR_NO_NETWORK = 1222&
Private Const ERROR_NO_CONNECTION = 8
Private Const ERROR_NO_DISCONNECT = 9
Private Const ERROR_DEVICE_IN_USE = 2404&
Private Const ERROR_NOT_CONNECTED = 2250&
Private Const ERROR_OPEN_FILES = 2401&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_SESSION_CREDENTIAL_CONFLICT = 1219&

Private Const CONNECT_UPDATE_PROFILE = &H1
Private Const RESOURCETYPE_DISK = &H1

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

Private lpNetResource As NETRESOURCE
 
Public Function Connect(ByVal HostName As String, ByVal LocalName As String, ByVal Username As String, ByVal Password As String, ErrorNum As Long, ErrorMsg As String)
    Dim lpUsername As String
    Dim lpPassword As String
    Dim rc As Long

    On Error GoTo Err_Connect
    ErrorNum = 0
    ErrorMsg = ""
    lpNetResource.dwType = RESOURCETYPE_DISK
    lpNetResource.lpLocalName = LocalName & Chr(0)
    'Drive Letter to use
    lpNetResource.lpRemoteName = HostName & Chr(0)
    'Network Path to share
    lpNetResource.lpProvider = Chr(0)
    lpPassword = Password & Chr(0)
    'password on share pass "" if none
    lpUsername = Username & Chr(0)
    'username to connect as (if applicable)
    rc = WNetAddConnection2(lpNetResource, lpPassword, lpUsername, CONNECT_UPDATE_PROFILE)
    If rc <> 0 Then GoTo Err_Connect

    Exit Function

Err_Connect:
    ErrorNum = rc
    ErrorMsg = WnetError(rc)
End Function

Public Function Disconnect(ByVal Name As String, ByVal ForceOff As Boolean, ErrorNum As Long, ErrorMsg As String)
    Dim rc As Long
    On Error GoTo Err_DisConnect
    ErrorNum = 0
    ErrorMsg = ""
    rc = WNetCancelConnection2(Name & Chr(0), CONNECT_UPDATE_PROFILE, ForceOff)
    If rc <> 0 Then GoTo Err_DisConnect

    Exit Function

Err_DisConnect:
    ErrorNum = rc
    ErrorMsg = WnetError(rc)
End Function

Private Function WnetError(errCode As Long) As String
    Select Case errCode
        Case ERROR_BAD_DEV_TYPE
            WnetError = "Bad device."
        Case ERROR_ALREADY_ASSIGNED
            WnetError = "Already Assigned."
        Case ERROR_ACCESS_DENIED
            WnetError = "Access Denied."
        Case ERROR_BAD_NET_NAME
            WnetError = "Bad net name"
        Case ERROR_BAD_PROFILE
            WnetError = "Bad Profile"
        Case ERROR_BAD_PROVIDER
            WnetError = "Bad Provider"
        Case ERROR_BUSY
            WnetError = "Busy"
        Case ERROR_CANCEL_VIOLATION
            WnetError = "Cancel Violation"
        Case ERROR_CANNOT_OPEN_PROFILE
            WnetError = "Cannot Open Profile"
        Case ERROR_DEVICE_ALREADY_REMEMBERED
            WnetError = "Device already remembered"
        Case ERROR_EXTENDED_ERROR
            WnetError = "Device already remembered"
        Case ERROR_INVALID_PASSWORD
            WnetError = "Invalid Password"
        Case ERROR_NO_NET_OR_BAD_PATH
            WnetError = "Could not find the specified device"
        Case ERROR_NO_NETWORK
            WnetError = "No Network Present"
        Case ERROR_DEVICE_IN_USE
            WnetError = "Connection Currently in use "
        Case ERROR_NOT_CONNECTED
            WnetError = "No Connection Present"
        Case ERROR_OPEN_FILES
            WnetError = "Files open and the force parameter is false"
        Case ERROR_MORE_DATA
            WnetError = "Buffer to small to hold network name, make lpnLength bigger"
        Case ERROR_SESSION_CREDENTIAL_CONFLICT
            WnetError = "The Credentials Supplied Conflict With An Existing Set Of Credentials"
        Case Else:
            WnetError = "Unrecognized Error " + Str(errCode) + "."
    End Select
End Function

Public Function GetConnectionSpeed(ServerAndShare As String, Username As String, Password As String, ErrorCode As Long, ErrorMessage As String) As Long
    Dim NetzResource As NETRESOURCE
    Dim Connection_Info As NETCONNECTINFOSTRUCT
    Dim result As Long
    Dim cntr As Integer
    Dim i As Integer
    Dim lp As Long
    Dim s As String
    Dim LocalFilename As String
    
    On Error GoTo Speed_Error

    cntr = 0
    
    ' Disconnect z drive, nothing should use this anyway, return any error
    Disconnect "z:", True, ErrorCode, ErrorMessage
    If (ErrorCode <> 0) And (ErrorCode <> ERROR_NOT_CONNECTED) Then
        GetConnectionSpeed = 0
        Exit Function
    End If
 
    ' connect to specified share for testing
    Connect ServerAndShare, "z:", Username, Password, ErrorCode, ErrorMessage

    'check link speed
    If (ErrorCode <> 0) And (ErrorCode <> ERROR_ALREADY_ASSIGNED) Then
    Else ' Use a test file to create traffic across the link
        ErrorCode = 0
        ErrorMessage = ""
        
        ' Create a 64Kb text file locally, then copy it to the remote server
        
        'If Dir("c:\temp", vbDirectory) = "" Then
        '    MkDir "c:\temp"
        'End If
        
        'Randomize Timer
        
        's = ""
        'LocalFilename = "c:\temp\TEMP_FILE.TXT"
        
        'If Dir(LocalFilename) = "" Then
        '    ' Create 64Kb file
        '    For lp = 1 To 65536
        '        s = s & Chr(65 + Int(Rnd(1) * 26) + 1)
        '    Next lp
        '    Open LocalFilename For Output As #1
        '        Print #1, s
        '    Close 1
        'End If
        
        ' Copy file to remote location
        'CopyFile LocalFilename, ServerAndShare & "\deleteme.txt", True

        Connection_Info.cbStructure = 1024
        result = MultinetGetConnectionPerformance(lpNetResource, Connection_Info)

        GetConnectionSpeed = Connection_Info.dwSpeed * 100
        Debug.Print "Delay: " & Connection_Info.dwDelay & " mS"
        Disconnect "z:", True, ErrorCode, ErrorMessage
        If (ErrorCode <> 0) And (ErrorCode <> ERROR_ALREADY_ASSIGNED) Then
        Else
            ErrorCode = 0
            ErrorMessage = ""
        End If
    End If
    
    'If Dir(ServerAndShare & "\deleteme.txt") <> "" Then
    '    Kill ServerAndShare & "\deleteme.txt"
    'End If
    
    Exit Function
Speed_Error:
    ErrorMessage = "Unexpected Error: " & Err.Description
    ErrorCode = Err.Number
    Err.Clear
    'If Dir(ServerAndShare & "\deleteme.txt") <> "" Then
    '    Kill ServerAndShare & "\deleteme.txt"
    'End If
End Function

