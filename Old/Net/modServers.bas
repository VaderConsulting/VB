Attribute VB_Name = "modServers"
Option Explicit
' Type used by NetServerSetInfo

Public Type SERVER_INFO_1005
    sv1005_comment As Long
End Type

' General definitions
Const ERROR_SUCCESS = 0
Const ERROR_MORE_DATA = 234
Const SV_TYPE_WORKSTATION = 1
Const SV_TYPE_SERVER = 2
Const SV_TYPE_DOMAIN_CTRL = 8
Const SV_TYPE_DOMAIN_BAKCTRL = 16
Const SV_TYPE_TIME_SOURCE = 32
Const SV_TYPE_AFP = 64
Const SV_TYPE_DOMAIN_MEMBER = 256
Const SV_TYPE_PRINTQ_SERVER = 512
Const SV_TYPE_DIALIN_SERVER = 1024
Const SV_TYPE_XENIX_SERVER = 2048
Const SV_TYPE_SERVER_UNIX = 2048
Const SV_TYPE_NT = 4096
Const SV_TYPE_WFW = 8192
Const SV_TYPE_POTENTIAL_BROWSER = 65536
Const SV_TYPE_BACKUP_BROWSER = 131072
Const SV_TYPE_MASTER_BROWSER = 262144
Const SV_TYPE_DOMAIN_MASTER = 524288
Const SV_TYPE_LOCAL_LIST_ONLY = 1073741824
Const SV_TYPE_DOMAIN_ENUM = 2147483648#
Const SV_TYPE_SQLSERVER = 4
Const SV_TYPE_NOVELL = 128
'Const SV_TYPE_DOMAIN_CTRL = 1048576
'Const SV_TYPE_DOMAIN_BAKCTRL = 2097152

Const SIZE_SI_101 = 24

Public Type SERVER_INFO_101
    dwPlatformId As Long
    lpszServerName As Long
    dwVersionMajor As Long
    dwVersionMinor As Long
    dwType As Long
    lpszComment As Long
End Type

Public Declare Function NetServerEnum Lib "netapi32.dll" ( _
    ByVal ServerName As String, _
    ByVal Level As Long, _
    Buffer As Long, _
    ByVal PrefMaxLen As Long, _
    EntriesRead As Long, _
    TotalEntries As Long, _
    ByVal servertype As Long, _
    ByVal domain As String, _
    resumehandle As Long) As Long

Public Declare Function NetGetDCName Lib "netapi32.dll" (ByVal sServerName As String, ByVal sDomainName As String, ByVal lPtr As Long) As Long
Public Declare Function NetServerSetInfo Lib "Netapi32" (sServerName As Byte, ByVal lLevel As Long, vBuffer As Long, ParmError As Long) As Long

Public Function SetServerInfo(ByVal xi_strComment As String, Optional ByVal xi_strServerName As String = "") As Boolean
    Dim p_bytServerName() As Byte
    Dim p_lngRtn As Long
    Dim p_lngSrvInfoRtn As Long
    Dim p_lngServEnumLevel As Long
    Dim p_lngParmError As Long
    Dim p_lngStrPtr As Long
    
    ' Initialize the variables
    If Trim$(xi_strServerName) = vbNullString Then
        p_bytServerName = vbNullChar
    Else
        p_bytServerName = Trim$(xi_strServerName) & vbNullChar
    End If
    
    p_lngServEnumLevel = 1005
    p_lngStrPtr = StrPtr(xi_strComment)
    p_lngRtn = NetServerSetInfo(sServerName:=p_bytServerName(0), _
    lLevel:=p_lngServEnumLevel, _
    vBuffer:=p_lngStrPtr, _
    ParmError:=p_lngParmError)
    
    If p_lngRtn = 0 Then
        SetServerInfo = True
    Else
        SetServerInfo = False
        Debug.Print Err.LastDllError
    End If

End Function


