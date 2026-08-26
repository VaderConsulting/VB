VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' GetIP.CLS
' Written by Frank Oquendo 7/13/99
' Adapted from code written by David Diederich
' Adapted for WAPS by D. Robinson 2000
' as listed at: http://www.netfokus.dk/vbadmincode/
' You may modify and/or redistribute this class module so long as
' you give credit to the original authors.
'
' GetIP returns the IP address of a host.
' GetIP has only two properties: HostName and IPAddress
' The IP address is automatically calculated when the host name is set.
' To calculate IP: objIP.HostName = "www.microsoft.com"
' To retrieve IP: tmp = objIP.IPAddress

Option Explicit
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hname As Long
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

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal HostName$, HostLen&) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal HostName$) As Long
Private Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Private m_hostname As String
Private m_ipaddress As String

Private Function hibyte(ByVal wParam As Integer)
 
    hibyte = wParam \ &H100 And &HFF&
 
End Function
 
Private Function lobyte(ByVal wParam As Integer)
 
    lobyte = wParam And &HFF&
 
End Function
 
Private Sub SocketsInitialize()
  Dim WSAD As WSADATA
  Dim iReturn As Integer
  Dim sLowByte As String, sHighByte As String, sMsg As String

      iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

      If iReturn <> 0 Then
          sMsg = "Winsock.dll is not responding."
          Err.Raise vbObjectError + 1, , sMsg
          End
      End If

      If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
          sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
          sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
          sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
          sMsg = sMsg & " is not supported by winsock.dll "
          Err.Raise vbObjectError + 2, , sMsg
          End
      End If

      If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
          sMsg = "This application requires a minimum of "
          sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
          Err.Raise vbObjectError + 3, , sMsg
          End
      End If

  End Sub

Private Sub SocketsCleanup()
    Dim lReturn As Long
    Dim sMsg As String
    
    lReturn = WSACleanup()
    
    If lReturn <> 0 Then
        sMsg = "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup """
        Err.Raise vbObjectError + 4, , sMsg
        End
    End If

End Sub

Private Sub GetAddress()

  Dim hostent_addr As Long
  Dim host As HOSTENT
  Dim hostip_addr As Long
  Dim temp_ip_address() As Byte
  Dim i As Integer
  Dim ip_address As String
  Dim sMsg As String

      hostent_addr = gethostbyname(m_hostname)

      If hostent_addr = 0 Then
          sMsg = "Winsock.dll is not responding."
          'Err.Raise vbObjectError + 1, , sMsg
          Exit Sub
      End If

      RtlMoveMemory host, hostent_addr, LenB(host)
      RtlMoveMemory hostip_addr, host.hAddrList, 4

      ReDim temp_ip_address(1 To host.hLength)
      RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

      For i = 1 To host.hLength
          ip_address = ip_address & temp_ip_address(i) & "."
      Next
      ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

      m_ipaddress = ip_address

End Sub

Private Sub Class_Initialize()
   SocketsInitialize
End Sub

Private Sub Class_Terminate()
   SocketsCleanup
End Sub

Public Property Let HostName(sHost As String)

    m_hostname = sHost
    GetAddress
    
End Property

Public Property Get HostName() As String

    HostName = m_hostname

End Property

Public Property Get IPAddress() As String

    IPAddress = m_ipaddress
    
End Property
