Attribute VB_Name = "modGetIP"
Option Explicit

' *******************************************
' Constants
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128
' *******************************************
' Custom Types
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
' *******************************************
'API Declares
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal Hostname As String, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal Hostname As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
' *******************************************

Public Function GetIPAddress(Hostname As String) As String
    Dim sHostname As String * 256
    Dim lpHost As Long
    Dim HOST As HOSTENT
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim i As Integer
    Dim sIPAddr As String
    
    If Not SocketsInitialize() Then
      GetIPAddress = ""
      SocketsCleanup
      Exit Function
    End If
    
    sHostname = Hostname
    ' Get Local Computername
    If gethostname(sHostname, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        Debug.Print "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    sHostname = Hostname & Chr(0) & String(256 - Len(Hostname) = 1, 0)
    lpHost = gethostbyname(sHostname)
    
    If lpHost = 0 Then
        GetIPAddress = ""
        Debug.Print "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
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
    
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
    
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
        Debug.Print "Winsock.dll is not responding."
        SocketsInitialize = False
        'End
    End If

    If LoByte(WSAD.wversion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wversion) = _
        WS_VERSION_MAJOR And HiByte(WSAD.wversion) < WS_VERSION_MINOR) Then

        strHighByte = Trim(Str(HiByte(WSAD.wversion)))
        strLowByte = Trim(Str(LoByte(WSAD.wversion)))
        strMsg = "Windows Sockets version " & strLowByte & "." & strHighByte
        strMsg = strMsg & " is not supported by winsock.dll "
        Debug.Print strMsg
        SocketsInitialize = False
    End If

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        strMsg = "This application requires a minimum of "
        strMsg = strMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        Debug.Print strMsg
        SocketsInitialize = False
    End If
    
    SocketsInitialize = True
End Function

Sub SocketsCleanup()
    Dim lngRetVal As Long

    lngRetVal = WSACleanup()

    If lngRetVal <> 0 Then
        'MsgBox "Socket error " & Trim(Str(lngRetVal)) & " occurred in Cleanup "
        Debug.Print "Socket error " & Trim(Str(lngRetVal)) & " occurred in Cleanup "
        'End
    End If
End Sub

Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function





