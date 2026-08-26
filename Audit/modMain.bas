Attribute VB_Name = "modMain"
Option Explicit
' ------------- GetIPAddress Constants --------------
Private Const IP_SUCCESS As Long = 0
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1
Private Const ERROR_SUCCESS As Long = 0
' ---------------------------------------------------

' ------------- GetMACAddress Constants -------------
Private Const NO_ERROR = 0
' ---------------------------------------------------

' ------------- RemoteRegistry Constants ------------
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const REG_OPTION_NON_VOLATILE = 0

Private Const BUFFER_SIZE As Long = 255
' ---------------------------------------------------

' ------------- Destination Reachable Constants -----
Private Const NETWORK_ALIVE_LAN As Long = &H1 'active LAN cards
Private Const NETWORK_ALIVE_WAN As Long = &H2 'active RAS connections
' ---------------------------------------------------

' ------------- Service Constants -------------------
Public Const SIZEOF_SERVICE_STATUS As Long = 36

'windows constants
Public Const LB_SETTABSTOPS As Long = &H192
Public Const ERROR_MORE_DATA = 234
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4

'Service State for Enum Requests (Bit Mask)
Public Const SERVICE_ACTIVE = &H1
Public Const SERVICE_INACTIVE = &H2
Public Const SERVICE_STATE_ALL = SERVICE_ACTIVE Or SERVICE_INACTIVE
                                     
'Service Types (Bit Mask)
'corresponds to SERVICE_STATUS.dwServiceType
Public Const SERVICE_KERNEL_DRIVER As Long = &H1
Public Const SERVICE_FILE_SYSTEM_DRIVER As Long = &H2
Public Const SERVICE_ADAPTER As Long = &H4
Public Const SERVICE_RECOGNIZER_DRIVER As Long = &H8
Public Const SERVICE_WIN32_OWN_PROCESS As Long = &H10
Public Const SERVICE_WIN32_SHARE_PROCESS As Long = &H20
Public Const SERVICE_INTERACTIVE_PROCESS As Long = &H100
Public Const SERVICE_WIN32 As Long = SERVICE_WIN32_OWN_PROCESS Or SERVICE_WIN32_SHARE_PROCESS
Public Const SERVICE_DRIVER As Long = SERVICE_KERNEL_DRIVER Or SERVICE_FILE_SYSTEM_DRIVER Or SERVICE_RECOGNIZER_DRIVER
Public Const SERVICE_TYPE_ALL As Long = SERVICE_WIN32 Or SERVICE_ADAPTER Or SERVICE_DRIVER Or SERVICE_INTERACTIVE_PROCESS

' ADSI Status info
Public Const ADS_SERVICE_STOPPED As Long = 1
Public Const ADS_SERVICE_START_PENDING As Long = 2
Public Const ADS_SERVICE_STOP_PENDING As Long = 3
Public Const ADS_SERVICE_RUNNING As Long = 4
Public Const ADS_SERVICE_CONTINUE_PENDING As Long = 5
Public Const ADS_SERVICE_PAUSE_PENDING As Long = 6
Public Const ADS_SERVICE_PAUSED As Long = 7
Public Const ADS_SERVICE_ERROR As Long = 8
Public Const ADS_SERVICE_OWN_PROCESS As Long = 10
Public Const ADS_SERVICE_SHARE_PROCESS As Long = 20
Public Const ADS_SERVICE_KERNEL_DRIVER As Long = 1
Public Const ADS_SERVICE_FILE_SYSTEM_DRIVER As Long = 2

' ADSI Start type
Public Const ADS_SERVICE_BOOT_START = &H0
Public Const ADS_SERVICE_SYSTEM_START = &H1
Public Const ADS_SERVICE_AUTO_START = &H2
Public Const ADS_SERVICE_DEMAND_START = &H3
Public Const ADS_SERVICE_SERVICE_DISABLED = &H4

'Service State
'corresponds to SERVICE_STATUS.dwCurrentState
Public Const SERVICE_STOPPED As Long = &H1
Public Const SERVICE_START_PENDING As Long = &H2
Public Const SERVICE_STOP_PENDING As Long = &H3
Public Const SERVICE_RUNNING As Long = &H4
Public Const SERVICE_CONTINUE_PENDING As Long = &H5
Public Const SERVICE_PAUSE_PENDING As Long = &H6
Public Const SERVICE_PAUSED As Long = &H7

'Controls Accepted  (Bit Mask)
'corresponds to SERVICE_STATUS.dwControlsAccepted
Public Const SERVICE_ACCEPT_STOP As Long = &H1
Public Const SERVICE_ACCEPT_PAUSE_CONTINUE As Long = &H2
Public Const SERVICE_ACCEPT_SHUTDOWN   As Long = &H4

'Windows type used to call the Net API
Public Const MAX_PREFERRED_LENGTH As Long = -1
Public Const NERR_SUCCESS As Long = 0&

Public Const SV_TYPE_WORKSTATION         As Long = &H1
Public Const SV_TYPE_SERVER              As Long = &H2
Public Const SV_TYPE_SQLSERVER           As Long = &H4
Public Const SV_TYPE_DOMAIN_CTRL         As Long = &H8
Public Const SV_TYPE_DOMAIN_BAKCTRL      As Long = &H10
Public Const SV_TYPE_TIME_SOURCE         As Long = &H20
Public Const SV_TYPE_AFP                 As Long = &H40
Public Const SV_TYPE_NOVELL              As Long = &H80
Public Const SV_TYPE_DOMAIN_MEMBER       As Long = &H100
Public Const SV_TYPE_PRINTQ_SERVER       As Long = &H200
Public Const SV_TYPE_DIALIN_SERVER       As Long = &H400
Public Const SV_TYPE_XENIX_SERVER        As Long = &H800
Public Const SV_TYPE_SERVER_UNIX         As Long = SV_TYPE_XENIX_SERVER
Public Const SV_TYPE_NT                  As Long = &H1000
Public Const SV_TYPE_WFW                 As Long = &H2000
Public Const SV_TYPE_SERVER_MFPN         As Long = &H4000
Public Const SV_TYPE_SERVER_NT           As Long = &H8000
Public Const SV_TYPE_POTENTIAL_BROWSER   As Long = &H10000
Public Const SV_TYPE_BACKUP_BROWSER      As Long = &H20000
Public Const SV_TYPE_MASTER_BROWSER      As Long = &H40000
Public Const SV_TYPE_DOMAIN_MASTER       As Long = &H80000
Public Const SV_TYPE_SERVER_OSF          As Long = &H100000
Public Const SV_TYPE_SERVER_VMS          As Long = &H200000
Public Const SV_TYPE_WINDOWS             As Long = &H400000  'Win95 and above
Public Const SV_TYPE_DFS                 As Long = &H800000  'Root of DFS tree
Public Const SV_TYPE_CLUSTER_NT          As Long = &H1000000 'NT Cluster
Public Const SV_TYPE_TERMINALSERVER      As Long = &H2000000 'Terminal Server
Public Const SV_TYPE_DCE                 As Long = &H10000000 'IBM DSS
Public Const SV_TYPE_ALTERNATE_XPORT     As Long = &H20000000 'rtn alternate transport
Public Const SV_TYPE_LOCAL_LIST_ONLY     As Long = &H40000000 'rtn local only
Public Const SV_TYPE_DOMAIN_ENUM         As Long = &H80000000
Public Const SV_TYPE_ALL                 As Long = &HFFFFFFFF

Public Const SV_PLATFORM_ID_OS2       As Long = 400
Public Const SV_PLATFORM_ID_NT        As Long = 500

'Mask applied to svX_version_major in order to obtain the major version number.
Public Const MAJOR_VERSION_MASK        As Long = &HF
' ---------------------------------------------------

' ------------- File Search Constants ---------------
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
' ---------------------------------------------------

' ------------- Common Types ------------------------
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
' ---------------------------------------------------

' ------------- GetIPAddress Types ------------------
Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type
' ---------------------------------------------------

' ------------- GetOS Types -------------------------
Public Type OSInfo
    OS As String
    Version As String
    Processor As String
    Uni_Or_Multi As String
End Type
' ---------------------------------------------------

' ------------- RemoteRegistry Types ----------------
Public Type ValueType
    Value As String
    Data As String
End Type

' ---------------------------------------------------
' ------------- Hotfixes Types ----------------------
Public Type Hotfix
    Number As String
    Comments As String
    Description As String
End Type
' ---------------------------------------------------

' ------------- Software Types ----------------------
Public Type Application
    DisplayName As String
    HelpLink As String
    Publisher As String
    DisplayVersion As String
    Key As String
End Type
' ---------------------------------------------------

' ------------- LocalAccount Types ------------------
Public Type User
    Name As String
    Fullname As String
    Description As String
    Disabled As Boolean
    Locked As Boolean
    Profile As String
    LogonScript As String
    HomeDirectory As String
End Type
' ---------------------------------------------------

'-------------- Destination Reachable Types ---------
Private Type QOCINFO
    dwSize As Long
    dwFlags As Long
    dwInSpeed As Long
    dwOutSpeed As Long
End Type
' ---------------------------------------------------

' ------------- Service Types -----------------------
Public Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

Public Type ENUM_SERVICE_STATUS
    lpServiceName As Long
    lpDisplayName As Long
    ServiceStatus As SERVICE_STATUS
End Type

Public Type SERVER_INFO_100
    sv100_platform_id As Long
    sv100_name As Long
End Type

Public Type Service
    DisplayName As String
    Name As String
    State As String
    AccountName As String
    Startup As String
End Type
' ---------------------------------------------------

' ------------- FileSearch Types --------------------
Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long     'e.g. 0x00000042 = "0.42"
    dwFileVersionMS As Long    'e.g. 0x00030075 = "3.75"
    dwFileVersionLS As Long    'e.g. 0x00000031 = "0.31"
    dwProductVersionMS As Long 'e.g. 0x00030010 = "3.10"
    dwProductVersionLS As Long 'e.g. 0x00000031 = "0.31"
    dwFileFlagsMask As Long    'e.g. 0x3F for version "0.42"
    dwFileFlags As Long        'e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long           'e.g. VOS_DOS_WINDOWS16
    dwFileType As Long         'e.g. VFT_DRIVER
    dwFileSubtype As Long      'e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long       'e.g. 0
    dwFileDateLS As Long       'e.g. 0
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Type FILE_PARAMS
    bRecurse As Boolean
    bList As Boolean
    bFound As Boolean
    sFileRoot As String
    sFileNameExt As String
    sResult As String
    nFileCount As Long
    nFileSize As Double
End Type
' ---------------------------------------------------

' ------------- RemoteRegistry Enums ----------------
Public Enum HKEYTree
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
End Enum
' ---------------------------------------------------

' ------------- Application Public Variables --------
Public strIPAddress As String
Public strComputername As String
Public strMACAddress As String
Public strDomainName As String
Public strServicePack As String
' ---------------------------------------------------

' ------------- GetOS Public Variables --------------
Public PCInfo As OSInfo
' ---------------------------------------------------

' ------------- Hotfixes Public Variables -----------
Public Hotfixes() As Hotfix
' ---------------------------------------------------

' ------------- Hotfixes Public Variables -----------
Public AppList() As Application
' ---------------------------------------------------

' ------------- LocalAccount Public Variables -------
Public UserList() As User
' ---------------------------------------------------

' ------------- Service Public Variables ------------
Public ServiceList() As Service
' ---------------------------------------------------

' ------------- File Search Public Variables --------
Public FP As FILE_PARAMS
' ---------------------------------------------------

' ------------- GetIPAddress Declarations -----------
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal hostname As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
' ----------------------------------------------------

' ------------- GetMACAddress Declarations -----------
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Long
Private Declare Function SendARP Lib "iphlpapi.dll" (ByVal DestIP As Long, ByVal SrcIP As Long, pMacAddr As Long, PhyAddrLen As Long) As Long
' ----------------------------------------------------

' ------------- RemoteRegistry Declarations ----------
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
' ----------------------------------------------------
                                              
' ------------- Destination Reachable Declarations ---
Private Declare Function IsDestinationReachable Lib "sensapi.dll" Alias "IsDestinationReachableA" (ByVal lpszDestination As String, lpQOCInfo As QOCINFO) As Long
' ----------------------------------------------------

' ------------- Service Declarations -----------------
Private Declare Function OpenSCManager Lib "advapi32" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function EnumServicesStatus Lib "advapi32" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long
Private Declare Function NetServerEnum Lib "netapi32" (ByVal servername As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal servertype As Long, ByVal domain As Long, resume_handle As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' ----------------------------------------------------

' ------------- File Search Declarations -------------
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, nVerSize As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
' ----------------------------------------------------

' ------------- Shared Declarations ------------------
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
' ----------------------------------------------------
       
' ------------- GetIPAddress Code --------------------
Public Function GetIPAddress(strHostname As String)
    Dim sHostName As String
    
    If SocketsInitialize() Then
    'obtain and pass the host address to the function
      GetIPAddress = GetIPFromHostName(strHostname)
      SocketsCleanup
    Else
        frmMain.lstOutput.AddItem "Windows Sockets for 32 bit Windows environments is not successfully responding."
    End If
End Function

Private Function SocketsInitialize() As Boolean
   Dim WSAD As WSADATA
   Dim success As Long
  
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
End Function

Private Sub SocketsCleanup()
   If WSACleanup() <> 0 Then
       frmMain.lstOutput.AddItem "Windows Sockets error occurred in Cleanup."
   End If
End Sub

Private Function GetIPFromHostName(ByVal sHostName As String) As String
  'converts a host name to an IP address

   Dim nBytes As Long
   Dim ptrHosent As Long  'address of HOSENT structure
   Dim ptrName As Long    'address of name pointer
   Dim ptrAddress As Long 'address of address pointer
   Dim ptrIPAddress As Long
   Dim ptrIPAddress2 As Long

   ptrHosent = gethostbyname(sHostName & vbNullChar)

   If ptrHosent <> 0 Then

     'assign pointer addresses and offset

     'Null-terminated list of addresses for the host.  The Address is offset 12 bytes from the start of
     'the HOSENT structure. Note: Here we are retrieving only the first address returned. To return more than
     'one, define sAddress as a string array and loop through the 4-byte ptrIPAddress members returned. The last
     'item is a terminating null. All addresses are returned in network byte order.
      ptrAddress = ptrHosent + 12
     
     'get the IP address
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress2, ByVal ptrIPAddress, 4

      GetIPFromHostName = GetInetStrFromPtr(ptrIPAddress2)

   End If
End Function

Private Function GetInetStrFromPtr(Address As Long) As String
   GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))
End Function
' --------------------------------------------------------

' ------------- GetMACAddress Code -----------------------
Public Function RemoteMACAddress(strIPAddress As String)
   Dim sRemoteMacAddress As String
   
   If Len(strIPAddress) > 0 Then
      If GetRemoteMACAddress(strIPAddress, sRemoteMacAddress, "-") Then
         RemoteMACAddress = sRemoteMacAddress
      Else
         RemoteMACAddress = "(SendARP call failed)"
      End If
   End If
End Function

Private Function GetRemoteMACAddress(ByVal sRemoteIP As String, sRemoteMacAddress As String, sDelimiter As String) As Boolean
   Dim dwRemoteIP As Long
   Dim pMacAddr As Long
   Dim bpMacAddr() As Byte
   Dim PhyAddrLen As Long
    
  'convert the string IP into an unsigned long value containing a suitable binary representation of the Internet address given
   dwRemoteIP = ConvertIPtoLong(sRemoteIP)

   If dwRemoteIP <> 0 Then
   
     'must set this up first!
      PhyAddrLen = 6
   
    'assume failure
     GetRemoteMACAddress = False
     
     'retrieve the remote MAC address
      If SendARP(dwRemoteIP, 0&, pMacAddr, PhyAddrLen) = NO_ERROR Then
      
         If (pMacAddr <> 0) And (PhyAddrLen <> 0) Then
      
           'returned value is a long pointer to the MAC address, so copy data to a byte array
            ReDim bpMacAddr(0 To PhyAddrLen - 1)
            CopyMemory bpMacAddr(0), pMacAddr, ByVal PhyAddrLen
            
           'convert the byte array to a string and return success
            sRemoteMacAddress = MakeMacAddress(bpMacAddr(), sDelimiter)
            GetRemoteMACAddress = True
            
         End If 'pMacAddr
      End If  'SendARP
   End If  'dwRemoteIP
End Function

Private Function ConvertIPtoLong(sIpAddress) As Long
   ConvertIPtoLong = inet_addr(sIpAddress)
End Function

Private Function MakeMacAddress(b() As Byte, sDelim As String) As String
   Dim cnt As Long
   Dim buff As String
   
   On Local Error GoTo MakeMac_error
 
  'so far, MAC addresses are exactly 6 segments in size (0-5)
   If UBound(b) = 5 Then
   
     'concatenate the first five values together and separate with the delimiter char
      For cnt = 0 To 4
         buff = buff & Right$("00" & Hex(b(cnt)), 2) & sDelim
      Next
      
     'and append the last value
      buff = buff & Right$("00" & Hex(b(5)), 2)
         
   End If  'UBound(b)
   
   MakeMacAddress = buff
   
MakeMac_exit:
   Exit Function
   
MakeMac_error:
   MakeMacAddress = "(error building MAC address)"
   Resume MakeMac_exit
End Function
' ---------------------------------------------------

' ----------------- Remote Registry Functions -------
Public Function GetRegValue(sRemMachName As String, sKeyName As String, sValueName As String) As String
    Dim lRetVal As Long         ' used to hold return value for all API calls
    
    ' sRemMachName              ' used by RegConnectRegistry
    Dim lTopLevelKey As Long    ' used by RegConnectRegistry
    Dim lHKeyhandle As Long     ' used by RegConnectRegistry & RegOpenKeyEx
    
    ' sKeyName As String        ' used by RegOpenKeyEx
    Dim lhkey As Long           ' used by RegOpenKeyEx & RegQueryValueEx & RegCloseKey
    
    ' sValueName As String      ' used by RegQueryValueEx
    Dim vValue As String        ' used by RegQueryValueEx
        
    lTopLevelKey = HKEY_LOCAL_MACHINE

    'Get handle of a top level registry key on remote machine
    lRetVal = RegConnectRegistry(sRemMachName, lTopLevelKey, lHKeyhandle)
    
    'Get handle of the key which contains the value you need to check
    lRetVal = RegOpenKeyEx(lHKeyhandle, sKeyName, 0, KEY_ALL_ACCESS, lhkey)
    
    'Get the value
    lRetVal = QueryValueEx(lhkey, sValueName, vValue)
    
    'Always a good idea to clean up
    RegCloseKey (lhkey)
    
    If lRetVal = 0 Or lRetVal = -1 Then    'QueryValueEx succeeded
        'Display the value obtained
        GetRegValue = vValue
    Else
        'Display error message
        Dim msg As String
        Select Case lRetVal
            Case 0, 2
                ' No error, or error 2 (value not found)
            Case Else
                msg = "An Error occured, Return value = " & lRetVal
        End Select
    End If
End Function

Private Function QueryValueEx(ByVal lhkey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhkey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE And lrc <> 6 Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhkey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhkey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit
End Function

Public Function KeyValues(hKey As HKEYTree, Path As String) As ValueType()
    Dim Values() As ValueType
    Dim hHKEY As Long
    Dim sName As String
    Dim Ret As Long
    Dim sData As String
    Dim RetData As Long
    Dim Count As Long
    
    ' Open the key
    If RegOpenKey(hKey, Path, hHKEY) = 0 Then
        
        ' Pad out a buffer
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
        
        ' Get the name of each value
        Do While RegEnumValue(hHKEY, Count, sName, Ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            ' Add the key
            If RetData > 0 Then
                ReDim Preserve Values(0 To Count) As ValueType
                Values(Count).Value = Left$(sName, Ret)
                Values(Count).Data = Left$(sData, RetData - 1)
            End If
            
            ' Set buffer for the next key
            Count = Count + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Loop
        
        ' Close key
        RegCloseKey hHKEY
    Else
        Err.Raise vbObjectError + 2, "KeyValues", "KeyValues" & vbCr & "Error opening registry key"
    End If
    KeyValues = Values
End Function

Public Function SubKeys(sRemMachName As String, sKeyName As String) As String()
    Dim Keys() As String
    Dim hHKEY As Long
    Dim strSubKeyName As String
    Dim Ret As Long
    Dim Count As Long
    '  -------------------
    Dim lTopLevelKey As Long    ' used by RegConnectRegistry
    Dim lhkey As Long           ' used by RegOpenKeyEx & RegQueryValueEx & RegCloseKey
    Dim lRetVal As Long         ' used to hold return value for all API calls
    Dim lastwrite As FILETIME   ' Receives last-write-to time
    Dim keyname As String       ' Receives keyname for each subkey
    Dim keylength As Long       ' Length of keyname
    Dim classname As String     ' Receives class name for each subkey
    Dim classlen As Long        ' Length of classname
    Dim handle As Long          ' handle to the opened key
    Dim lHKeyhandle As Long
    
    lTopLevelKey = HKEY_LOCAL_MACHINE
    
    'Get handle of a top level registry key on remote machine
    lRetVal = RegConnectRegistry(sRemMachName, lTopLevelKey, lHKeyhandle)
    
    'Get handle of the key which contains the value you need to check
    lRetVal = RegOpenKeyEx(lHKeyhandle, sKeyName, 0, KEY_ENUMERATE_SUB_KEYS, handle)
    
    If lRetVal <> 0 Then
        MsgBox "Registry key could not be opened."
    Else
        ' Enumerate keys starting from 0
        Count = 0
            
        ' Get the name of each key
        While lRetVal = 0
            ' Set buffer for the next key
            keyname = Space(255): keylength = 255
            classname = Space(255): classlen = 255
            
            lRetVal = RegEnumKeyEx(handle, Count, keyname, keylength, ByVal 0, classname, classlen, lastwrite)
            
            ' <> ERROR_NO_MORE_ITEMS
            If lRetVal = 0 Then
                ' Add the key
                ReDim Preserve Keys(0 To Count) As String
                keyname = Left(keyname, keylength)
                Keys(Count) = keyname
            
                ' Increment Count for the API call to retrieve the next key
                Count = Count + 1
            End If
        Wend
    End If

    ' Close key
    RegCloseKey (lhkey)
    
    On Error GoTo ERRORHANDLE:
    SubKeys = Keys
    Exit Function
ERRORHANDLE:
    'Do nothing
End Function
' ---------------------------------------------------

' ----------------- GetOSInfo functions -------------
Public Sub GetOSInfo(strComputername As String, PCInfo As OSInfo)
    Dim oComputer As IADsComputer
    On Error GoTo NoAccess
    Set oComputer = GetObject("WinNT://" & strComputername & ",computer")
    
    PCInfo.OS = oComputer.OperatingSystem
    PCInfo.Version = oComputer.OperatingSystemVersion
    PCInfo.Processor = oComputer.Processor
    PCInfo.Uni_Or_Multi = oComputer.ProcessorCount
    
    frmMain.lstOutput.AddItem "OS: " & PCInfo.OS
    frmMain.lstOutput.AddItem "Version: " & PCInfo.Version
    frmMain.lstOutput.AddItem "Processor: " & PCInfo.Processor
    frmMain.lstOutput.AddItem "ProcessorCount: " & PCInfo.Uni_Or_Multi
    Set oComputer = Nothing
    Exit Sub
NoAccess:
    frmMain.lstOutput.AddItem "Error " & Err.Number & " (" & Err.Description & ")"
    Set oComputer = Nothing
End Sub
' ---------------------------------------------------

' ----------------- Hotfixes Functions --------------
Public Function GetHotfixes(strComputername As String) As Hotfix()
    Dim Fixes() As String
    Dim HFixes() As Hotfix
    Dim i As Integer
    
    On Error GoTo NoAccess
    
    ' This is performed in two parts.
    
    ' Firstly, get the list of hotfixes
    Fixes = SubKeys(strComputername, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\HotFix")
    ReDim HFixes(UBound(Fixes()))
    
    For i = 0 To UBound(Fixes())
        HFixes(i).Number = Fixes(i)
        frmMain.lstOutput.AddItem "Hotfix: " & Fixes(i)
        ' For each hotfix found, get extra information
        HFixes(i).Comments = GetRegValue(strComputername, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\HotFix\" & HFixes(i).Number, "Comments")
        HFixes(i).Description = GetRegValue(strComputername, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\HotFix\" & HFixes(i).Number, "Fix Description")
        frmMain.lstOutput.AddItem "Comments: " & HFixes(i).Comments
        frmMain.lstOutput.AddItem "Fix Description: " & HFixes(i).Description
        frmMain.lstOutput.AddItem ""
    Next i
    
    GetHotfixes = HFixes
    Exit Function
NoAccess:
    frmMain.lstOutput.AddItem "Error " & Err.Number & " (" & Err.Description & ")"
End Function
' ---------------------------------------------------

' ----------------- Software Functions --------------
Public Function GetApplicationList(strComputername As String) As Application()
    Dim Apps() As String
    Dim Applications() As Application
    Dim i As Integer
    
    On Error GoTo NoAccess
    
    ' This is performed in two parts.
    
    ' Firstly, get the list of software
    Apps = SubKeys(strComputername, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    ReDim Applications(UBound(Apps()))
    
    For i = 0 To UBound(Apps())
        Applications(i).Key = Apps(i)
        'frmMain.lstOutput.AddItem "Application: " & Apps(i)
        
        ' For each application found, get extra information
        Applications(i).DisplayName = GetRegValue(strComputername, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Applications(i).Key, "DisplayName")
        Applications(i).DisplayVersion = GetRegValue(strComputername, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Applications(i).Key, "DisplayVersion")
        Applications(i).Publisher = GetRegValue(strComputername, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Applications(i).Key, "Publisher")
        Applications(i).HelpLink = GetRegValue(strComputername, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Applications(i).Key, "HelpLink")
        If Applications(i).DisplayName <> "" Then
            frmMain.lstOutput.AddItem "Display Name: " & Applications(i).DisplayName
            frmMain.lstOutput.AddItem "Display Version: " & Applications(i).DisplayVersion
            frmMain.lstOutput.AddItem "Publisher: " & Applications(i).Publisher
            'frmMain.lstOutput.AddItem "Help URL: " & Applications(i).HelpLink
        frmMain.lstOutput.AddItem ""
        End If
    Next i
    
    GetApplicationList = Applications
    Exit Function
NoAccess:
    frmMain.lstOutput.AddItem "Error " & Err.Number & " (" & Err.Description & ")"
End Function
' ---------------------------------------------------

' ----------------- LocalAccount functions ----------
Public Function ListUsers(strComputer)
    Dim objComputer As IADsComputer
    Dim objUser As IADsUser
    Dim Count As Integer
    
    On Error GoTo NoAccess
    
    Set objComputer = GetObject("WinNT://" & strComputer & ",Computer")
    
    objComputer.Filter = Array("User")
    For Each objUser In objComputer
        ReDim Preserve UserList(Count)
        UserList(Count).Name = objUser.Name
        UserList(Count).Fullname = objUser.Fullname
        UserList(Count).Description = objUser.Description
        UserList(Count).Disabled = CBool(objUser.AccountDisabled)
        UserList(Count).Locked = CBool(objUser.IsAccountLocked)
        UserList(Count).Profile = objUser.Profile
        UserList(Count).LogonScript = objUser.LoginScript
        UserList(Count).HomeDirectory = objUser.HomeDirectory
        
        frmMain.lstOutput.AddItem "Name: " & UserList(Count).Name
        frmMain.lstOutput.AddItem "Fullname: " & UserList(Count).Fullname
        frmMain.lstOutput.AddItem "Description: " & UserList(Count).Description
        frmMain.lstOutput.AddItem "AccountDisabled: " & UserList(Count).Disabled
        frmMain.lstOutput.AddItem "IsAccountLocked: " & UserList(Count).Locked
        frmMain.lstOutput.AddItem "Profile: " & UserList(Count).Profile
        frmMain.lstOutput.AddItem "LoginScript: " & UserList(Count).LogonScript
        frmMain.lstOutput.AddItem "HomeDirectory: " & UserList(Count).HomeDirectory
        frmMain.lstOutput.AddItem ""
        Count = Count + 1
    Next
    
    Exit Function
NoAccess:
    frmMain.lstOutput.AddItem "Error " & Err.Number & " (" & Err.Description & ")"
End Function
' ---------------------------------------------------

Public Function DestinationReachable(strComputername As String) As Boolean
    Dim c As Integer
    Dim TotalRequired As Integer
    Dim result As Long
    Dim tmp As String
    Dim qoc As QOCINFO

    With qoc
        .dwSize = Len(qoc)
    End With
    
    result = IsDestinationReachable(strComputername, qoc)
    
    With qoc
        Select Case .dwFlags
            Case NETWORK_ALIVE_LAN
                tmp = "computer has one or more active LAN cards"
            Case NETWORK_ALIVE_WAN
                tmp = "computer has one or more active RAS connections"
            Case Else
        End Select
        
        Select Case result
            Case 0
                frmMain.lstOutput.AddItem "Destination reachable: No"
                DestinationReachable = False
                Disable
            Case 1
                frmMain.lstOutput.AddItem "Destination reachable: Yes (" & tmp & ")"
                DestinationReachable = True
                Enable
        End Select
    End With
End Function
' ---------------------------------------------------

' ----------------- GUI functions -------------------
Public Sub Enable()
    frmMain.cmdGetMAC.Enabled = True
    frmMain.cmdDomain.Enabled = True
    frmMain.cmdOS.Enabled = True
    frmMain.cmdServicePack.Enabled = True
    frmMain.cmdHotfixes.Enabled = True
    frmMain.cmdApplications.Enabled = True
    'frmMain.cmdExecutables.Enabled = True
    frmMain.cmdLocalAccounts.Enabled = True
    frmMain.cmdServices.Enabled = True
    frmMain.cmdSave.Enabled = True
End Sub

Public Sub Disable()
    frmMain.cmdGetMAC.Enabled = False
    frmMain.cmdDomain.Enabled = False
    frmMain.cmdOS.Enabled = False
    frmMain.cmdServicePack.Enabled = False
    frmMain.cmdHotfixes.Enabled = False
    frmMain.cmdApplications.Enabled = False
    'frmMain.cmdExecutables.Enabled = False
    frmMain.cmdLocalAccounts.Enabled = False
    frmMain.cmdServices.Enabled = False
    frmMain.cmdSave.Enabled = False
End Sub
' ---------------------------------------------

' ----------------- Service functions----------
Public Function EnumSystemServices(SERVICE_TYPE As Long, sMachine As String) As Long
    Dim hSCManager As Long
    Dim pntr() As ENUM_SERVICE_STATUS
    Dim cbBuffSize As Long
    Dim cbRequired As Long
    Dim dwReturned As Long
    Dim hEnumResume As Long
    Dim cbBuffer As Long
    Dim success As Long
    Dim i As Long
    
    ' These five just help keep the code lines from becoming too long for html display
    Dim sSvcName As String
    Dim sDispName As String
    Dim dwState As Long
    Dim dwType As Long
    Dim dwCtrls As Long
    
    Dim Services() As Service
    
    ' Establish a connection to the service control manager on the computer and open the local service control manager database.
    hSCManager = OpenSCManager("\\" & sMachine, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    
    If hSCManager <> 0 Then
        'Get buffer size by calling EnumServicesStatus.
        
        'To determine the required buffer size, call EnumServicesStatus with cbBuffer and hEnumResume set to zero. EnumServicesStatus
        'fails (returns 0), and Err.LastDLLError returns ERROR_MORE_DATA, filling cbRequired with the size, in bytes, of the buffer
        'required to hold the array of structures and their data.
        success = EnumServicesStatus(hSCManager, SERVICE_WIN32, SERVICE_TYPE, ByVal &H0, &H0, cbRequired, dwReturned, hEnumResume)
        
        'If success is 0 and the LastDllError is ERROR_MORE_DATA, use returned info to create
        'the required data buffer
        If success = 0 And Err.LastDllError = ERROR_MORE_DATA Then
            'Calculate number of structures needed and redimension the array
            cbBuffer = (cbRequired \ SIZEOF_SERVICE_STATUS) + 1
            ReDim pntr(0 To cbBuffer)
            
            'Set cbBuffSize equal to the size of the buffer
            cbBuffSize = cbBuffer * SIZEOF_SERVICE_STATUS
            
            'Enumerate the services. If the function succeeds, the return value is nonzero. If the function fails,
            'the return value is zero. In addition, hEnumResume must be set to 0.
            hEnumResume = 0
            If EnumServicesStatus(hSCManager, SERVICE_WIN32, SERVICE_TYPE, pntr(0), cbBuffSize, cbRequired, dwReturned, hEnumResume) Then
                'pntr() array is now filled with service data, so it is a simple matter of extracting the required information.
                For i = 0 To dwReturned - 1
                    ReDim Preserve Services(i)
                    sDispName = GetStrFromPtrA(ByVal pntr(i).lpDisplayName)
                    sSvcName = GetStrFromPtrA(ByVal pntr(i).lpServiceName)
                    dwState = pntr(i).ServiceStatus.dwCurrentState
                    dwType = pntr(i).ServiceStatus.dwServiceType
                    dwCtrls = pntr(i).ServiceStatus.dwControlsAccepted
                    
                    GetServiceInfo sMachine, sSvcName, Services(i).AccountName, Services(i).Startup
                    
                    Services(i).DisplayName = sDispName
                    Services(i).Name = sSvcName
                    Services(i).State = GetServiceState(dwState)
                    
                    frmMain.lstOutput.AddItem "Display name: " & Services(i).DisplayName
                    frmMain.lstOutput.AddItem "Service name: " & Services(i).Name
                    frmMain.lstOutput.AddItem "State: " & Services(i).State
                    frmMain.lstOutput.AddItem "Account: " & Services(i).AccountName
                    frmMain.lstOutput.AddItem "Startup: " & Services(i).Startup
                    
                    frmMain.lstOutput.AddItem ""
                Next
            Else
                frmMain.lstOutput.AddItem "EnumServicesStatus error " & CStr(Err.LastDllError)
            End If  'If EnumServicesStatus
        Else
            frmMain.lstOutput.AddItem "ERROR_MORE_DATA not returned.  Error " & CStr(Err.LastDllError)
        End If  'If success = 0 And Err.LastDllError
    Else
        frmMain.lstOutput.AddItem "OpenSCManager failed.  Error = " & CStr(Err.LastDllError)
    End If  'If hSCManager <> 0
    
    'Clean up
    Call CloseServiceHandle(hSCManager)
    
    Services = ServiceList
    
    'return the number of services returned as a sign of success
    EnumSystemServices = dwReturned
End Function

Private Function GetServers(sDomain As String) As Long
  ' Lists all servers of the specified type that are visible in a domain.
   Dim bufptr          As Long
   Dim dwEntriesread   As Long
   Dim dwTotalentries  As Long
   Dim dwResumehandle  As Long
   Dim se100           As SERVER_INFO_100
   Dim success         As Long
   Dim nStructSize     As Long
   Dim cnt             As Long

   nStructSize = LenB(se100)
   
  'Call passing MAX_PREFERRED_LENGTH to have the API allocate required memory for the return values.
  '
  'The call is enumerating all machines on the network (SV_TYPE_ALL); however, by Or'ing
  'specific bit masks for defined types you can customize the returned data. For example, a
  'value of 0x00000003 combines the bit masks for SV_TYPE_WORKSTATION (0x00000001) and
  'SV_TYPE_SERVER (0x00000002).
  '
  'dwServerName must be Null. The level parameter (100 here) specifies the data structure being
  'used (in this case a SERVER_INFO_100 structure).
  '
  'The domain member is passed as Null, indicating machines on the primary domain are to be retrieved.
  'If you decide to use this member, pass StrPtr("domain name"), not the string itself.
   success = NetServerEnum(0&, 100, bufptr, MAX_PREFERRED_LENGTH, dwEntriesread, dwTotalentries, SV_TYPE_WORKSTATION Or SV_TYPE_SERVER, 0&, dwResumehandle)

  'if all goes well
   If success = NERR_SUCCESS And _
      success <> ERROR_MORE_DATA Then
      
    'loop through the returned data, adding each machine to the list
      For cnt = 0 To dwEntriesread - 1
        'get one chunk of data and cast into an LOCALGROUP_INFO_1 type in order to add the name to a list
         CopyMemory se100, ByVal bufptr + (nStructSize * cnt), nStructSize
      Next
   End If
   
  'clean up, regardless of success
   Call NetApiBufferFree(bufptr)
End Function

Public Function GetServiceControl(dwControl As Long) As String
   Dim tmp As String

   If dwControl Then
      If (dwControl And SERVICE_ACCEPT_STOP) Then tmp = tmp & "stop, "
      If (dwControl And SERVICE_ACCEPT_PAUSE_CONTINUE) Then tmp = tmp & "pause, "
      If (dwControl And SERVICE_ACCEPT_SHUTDOWN) Then tmp = tmp & "shutdown"
   End If

   GetServiceControl = tmp
End Function

Public Function GetServiceType(dwType As Long) As String
   Dim sType As String
   
   ' The following is from the API
   'If (dwType And SERVICE_WIN32_OWN_PROCESS) Then sType = sType & "own proc, "
   'If (dwType And SERVICE_WIN32_SHARE_PROCESS) Then sType = sType & "share, "
   'If (dwType And SERVICE_KERNEL_DRIVER) Then sType = sType & "kernel, "
   'If (dwType And SERVICE_FILE_SYSTEM_DRIVER) Then sType = sType & "filesys, "
   'If (dwType And SERVICE_INTERACTIVE_PROCESS) Then sType = sType & "interactive"
   
   ' The following is from the ADSI documentation
   If (dwType And ADS_SERVICE_OWN_PROCESS) Then sType = sType & "own process, "
   If (dwType And ADS_SERVICE_SHARE_PROCESS) Then sType = sType & "shared process, "
   If (dwType And ADS_SERVICE_KERNEL_DRIVER) Then sType = sType & "kernel, "
   If (dwType And ADS_SERVICE_FILE_SYSTEM_DRIVER) Then sType = sType & "driver"
   
   GetServiceType = sType
End Function

Public Function GetServiceState(dwState As Long) As String
   Select Case dwState
      Case SERVICE_STOPPED: GetServiceState = "Stopped"
      Case SERVICE_START_PENDING: GetServiceState = "Start pending"
      Case SERVICE_STOP_PENDING: GetServiceState = "Stop pending"
      Case SERVICE_RUNNING: GetServiceState = "Running"
      Case SERVICE_CONTINUE_PENDING: GetServiceState = "Continue pending"
      Case SERVICE_PAUSE_PENDING: GetServiceState = "Pause pending"
      Case SERVICE_PAUSED: GetServiceState = "Paused"
   End Select
End Function

Public Function GetServiceStartType(dwType As Long) As String
    Select Case dwType
      Case ADS_SERVICE_BOOT_START: GetServiceStartType = "Boot"
      Case ADS_SERVICE_SYSTEM_START: GetServiceStartType = "System"
      Case ADS_SERVICE_AUTO_START: GetServiceStartType = "Auto"
      Case ADS_SERVICE_DEMAND_START: GetServiceStartType = "Manual"
      Case ADS_SERVICE_SERVICE_DISABLED: GetServiceStartType = "Disabled"
   End Select
End Function

Public Function GetPointerToByteStringW(ByVal dwData As Long) As String
    Dim tmp() As Byte
    Dim tmplen As Long
    
    If dwData <> 0 Then
        tmplen = lstrlenW(dwData) * 2
        If tmplen <> 0 Then
            ReDim tmp(0 To (tmplen - 1)) As Byte
            CopyMemory tmp(0), ByVal dwData, tmplen
            GetPointerToByteStringW = tmp
        End If
    End If
End Function

Public Function GetServiceInfo(strComputer As String, strService As String, ByRef strAccountName As String, ByRef strStartupType As String) As String
    Dim cp As IADsComputer
    Dim sr As IADsService
    On Error GoTo Cleanup
    
    Set cp = GetObject("WinNT://" & strComputer & ",computer")
    Set sr = cp.GetObject("Service", strService)
    
    strAccountName = sr.ServiceAccountName
    strStartupType = GetServiceStartType(sr.StartType)
    
Cleanup:
    If (Err.Number <> 0) Then
        frmMain.lstOutput.AddItem ("An error has occurred. " & Err.Number)
    End If
    Set cp = Nothing
    Set sr = Nothing
End Function
' ---------------------------------------------------

' ----------------- File Search functions -----------
Public Function SearchForFiles(FP As FILE_PARAMS) As Double
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long
    Dim nSize As Long
    Dim sPath As String
    Dim sRoot As String
    Dim sTmp As String
    
    sRoot = QualifyPath(FP.sFileRoot)
    sPath = sRoot & "*.*"
    
    'obtain handle to the first match
    hFile = FindFirstFile(sPath, WFD)
    
    'if valid ...
    If hFile <> INVALID_HANDLE_VALUE Then
        
        'This is where the method obtains the file list and data for the folder passed.
        '
        'GetFileInformation function returns the size, in bytes, of the files found matching the
        'filespec in the passed folder, so its assigned to nSize. It is not directly assigned
        'to FP.nFileSize because nSize is incremented below if a recursive search was specified.
        nSize = GetFileInformation(FP)
        FP.nFileSize = nSize
        
        Do
            'if the returned item is a folder...
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            
                '..and the Recurse flag was specified
                If FP.bRecurse Then
                
                    'remove trailing nulls
                    sTmp = TrimNull(WFD.cFileName)
                    
                    'and if the folder is not the default self and parent folders...
                    If sTmp <> "." And sTmp <> ".." Then
                        '..then the item is a real folder, which may contain other sub folders, so assign
                        'the new folder name to FP.sFileRoot and recursively call this function again with
                        'the amended information.
                        '
                        'Since nSize is a local variable whose value is both set above as well as returned as the
                        'function call value, nSize needs to be added to previous calls in order to maintain accuracy.
                        '
                        'However, because the nFileSize member of FILE_PARAMS is passed back and forth through
                        'the calls, nSize is simply assigned to it after the recursive call finishes.
                        FP.sFileRoot = sRoot & sTmp
                        nSize = nSize + SearchForFiles(FP)
                        FP.nFileSize = nSize
                    End If
                End If
            End If
            
            'continue looping until FindNextFile returns 0 (no more matches)
        Loop While FindNextFile(hFile, WFD)
        
        'close the find handle
        hFile = FindClose(hFile)
    End If
    
    'because this routine is recursive, return the size of matching files
    SearchForFiles = nSize
End Function

Private Function GetFileInformation(FP As FILE_PARAMS) As Long
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long
    Dim nSize As Long
    Dim sPath As String
    Dim sRoot As String
    Dim sTmp As String
    'Dim itmx As ListItem
    
    'FP.sFileRoot (assigned to sRoot) contains the path to search.
    '
    'FP.sFileNameExt (assigned to sPath) contains the full path and filespec.
    sRoot = QualifyPath(FP.sFileRoot)
    sPath = sRoot & FP.sFileNameExt
    
    'obtain handle to the first filespec match
    hFile = FindFirstFile(sPath, WFD)
    
    'if valid ...
    If hFile <> INVALID_HANDLE_VALUE Then
    
        Do
            'remove trailing nulls
            sTmp = TrimNull(WFD.cFileName)
            
            'Even though this routine uses filespecs, *.* is still valid and will cause the search
            'to return folders as well as files, so a check against folders is still required.
            If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            
                'file found, so increase the file count
                FP.nFileCount = FP.nFileCount + 1
                
                'retrieve the size and assign to nSize to be returned at the end of this function call
                nSize = nSize + (WFD.nFileSizeHigh * (MAXDWORD + 1)) + WFD.nFileSizeLow
                
                'add to the list if the flag indicates
                If FP.bList Then
                    'got the data, so add it to the listview
                    'Set itmx = ListView1.ListItems.Add(, , LCase$(sTmp))
                    
                    'itmx.SubItems(1) = GetFileVersion(sRoot & sTmp)
                    'itmx.SubItems(3) = GetFileSizeStr(WFD.nFileSizeHigh + WFD.nFileSizeLow)
                    'itmx.SubItems(2) = GetFileDescription(sRoot & sTmp)
                    'itmx.SubItems(4) = LCase$(sRoot)
                End If
            End If
        Loop While FindNextFile(hFile, WFD)
        'close the handle
        hFile = FindClose(hFile)
    End If
    
    'return the size of files found
    GetFileInformation = nSize
End Function

Private Function GetFileSizeStr(fsize As Long) As String
    GetFileSizeStr = Format$((fsize), "###,###,###")   '& " kb"
End Function

Private Function QualifyPath(sPath As String) As String
  'assures that a passed path ends in a slash
  
    If Right$(sPath, 1) <> "\" Then
        QualifyPath = sPath & "\"
    Else
        QualifyPath = sPath
    End If
End Function

Public Function TrimNull(startstr As String) As String
    'returns the string up to the first null, if present, or the passed string
    Dim pos As Integer
    
    pos = InStr(startstr, Chr$(0))
    
    If pos Then
        TrimNull = Left$(startstr, pos - 1)
        Exit Function
    End If
    
    TrimNull = startstr
End Function

Private Function HiWord(dw As Long) As Long
    If dw And &H80000000 Then
        HiWord = (dw \ 65535) - 1
    Else
        HiWord = dw \ 65535
    End If
End Function
  
Private Function LoWord(dw As Long) As Long
    If dw And &H8000& Then
        LoWord = &H8000& Or (dw And &H7FFF&)
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

Private Function GetFileDescription(sSourceFile As String) As String
    Dim FI As VS_FIXEDFILEINFO
    Dim sBuffer() As Byte
    Dim nBufferSize As Long
    Dim lpBuffer As Long
    Dim nVerSize As Long
    Dim nUnused As Long
    Dim tmpVer As String
    Dim sBlock As String
    
    If Len(sSourceFile) > 0 Then
        'set file that has the encryption level info and call to get required size
        nBufferSize = GetFileVersionInfoSize(sSourceFile, nUnused)
        
        ReDim sBuffer(nBufferSize)
        
        If nBufferSize > 0 Then
            'get the version info
            Call GetFileVersionInfo(sSourceFile, 0&, nBufferSize, sBuffer(0))
            Call VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize)
            Call CopyMemory(FI, ByVal lpBuffer, Len(FI))
            
            If VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lpBuffer, nVerSize) Then
                If nVerSize Then
                    tmpVer = GetPointerToString(lpBuffer, nVerSize)
                    tmpVer = Right("0" & Hex(Asc(Mid(tmpVer, 2, 1))), 2) & _
                    Right("0" & Hex(Asc(Mid(tmpVer, 1, 1))), 2) & _
                    Right("0" & Hex(Asc(Mid(tmpVer, 4, 1))), 2) & _
                    Right("0" & Hex(Asc(Mid(tmpVer, 3, 1))), 2)
                    sBlock = "\StringFileInfo\" & tmpVer & "\FileDescription"
                    
                    'Get predefined version resources
                    If VerQueryValue(sBuffer(0), sBlock, lpBuffer, nVerSize) Then
                        If nVerSize Then
                            'get the file description
                            GetFileDescription = GetStrFromPtrA(lpBuffer)
                        End If  'If nVerSize
                    End If  'If VerQueryValue
                End If  'If nVerSize
            End If  'If VerQueryValue
        End If  'If nBufferSize
    End If  'If sSourcefile
End Function

Private Function GetPointerToString(lpString As Long, nBytes As Long) As String
    Dim Buffer As String
    
    If nBytes Then
        Buffer = Space$(nBytes)
        CopyMemory ByVal Buffer, ByVal lpString, nBytes
        GetPointerToString = Buffer
    End If
End Function

Private Function GetFileVersion(sDriverFile As String) As String
    Dim FI As VS_FIXEDFILEINFO
    Dim sBuffer() As Byte
    Dim nBufferSize As Long
    Dim lpBuffer As Long
    Dim nVerSize As Long
    Dim nUnused As Long
    Dim tmpVer As String
    
    'GetFileVersionInfoSize determines whether the operating system can obtain version information about a specified
    'file. If version information is available, it returns the size in bytes of that information. As with other
    'file installation functions, GetFileVersionInfoSize works only with Win32 file images.
    '
    'A empty variable must be passed as the second parameter, which the call returns 0 in.
    nBufferSize = GetFileVersionInfoSize(sDriverFile, nUnused)
    
    If nBufferSize > 0 Then
        'create a buffer to receive file-version
        '(FI) information.
        ReDim sBuffer(nBufferSize)
        Call GetFileVersionInfo(sDriverFile, 0&, nBufferSize, sBuffer(0))
        
        'VerQueryValue function returns selected version info from the specified version-information resource. Grab
        'the file info and copy it into the  VS_FIXEDFILEINFO structure.
        Call VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize)
        Call CopyMemory(FI, ByVal lpBuffer, Len(FI))
        
        'extract the file version from the FI structure
        tmpVer = Format$(HiWord(FI.dwFileVersionMS)) & "." & Format$(LoWord(FI.dwFileVersionMS), "00") & "."
        
        If FI.dwFileVersionLS > 0 Then
            tmpVer = tmpVer & Format$(HiWord(FI.dwFileVersionLS), "00") & "." & Format$(LoWord(FI.dwFileVersionLS), "00")
        Else
            tmpVer = tmpVer & Format$(FI.dwFileVersionLS, "0000")
        End If
    End If
    
    GetFileVersion = tmpVer
End Function
' ---------------------------------------------------

' ----------------- Common functions ----------------
Private Function GetMachineName() As String
   Dim sHostName As String * 256
  
   If gethostname(sHostName, 256) = ERROR_SUCCESS Then
      GetMachineName = Trim$(sHostName)
   End If
End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String
   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

Public Function Cleanup()
    PCInfo.OS = ""
    PCInfo.Processor = ""
    PCInfo.Uni_Or_Multi = ""
    PCInfo.Version = ""
    
    Erase Hotfixes()
    Erase AppList()
    Erase UserList()
    Erase ServiceList()
    
    strIPAddress = ""
    strComputername = ""
    strMACAddress = ""
    strDomainName = ""
    strServicePack = ""
    
    FP.bFound = False
    FP.bList = False
    FP.bRecurse = False
    FP.nFileCount = 0
    FP.nFileSize = 0
    FP.sFileNameExt = "*.exe"
    FP.sFileRoot = ""
    FP.sResult = ""
End Function
' ---------------------------------------------------


