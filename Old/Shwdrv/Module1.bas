Attribute VB_Name = "Module1"
    Option Explicit
    
    Public Declare Function NetServerEnum _
        Lib "netapi32.dll" ( _
        vServerName As Any, _
        ByVal lLevel As Long, _
        vBufptr As Any, _
        lPrefmaxlen As Long, _
        lEntriesRead As Long, _
        lTotalEntries As Long, _
        vServerType As Any, _
        ByVal sDomain As String, _
        vResumeHandle As Any) _
        As Long
   
    Public Declare Sub RtlMoveMemory _
        Lib "kernel32" ( _
        dest As Any, _
        vSrc As Any, _
        ByVal lSize&)
    
    Public Declare Sub lstrcpyW _
        Lib "kernel32" ( _
        vDest As Any, _
        ByVal sSrc As Any)
        
    Declare Sub lstrcpy _
        Lib "kernel32" ( _
        vDest As Any, _
        ByVal vSrc As Any)
    
    Declare Sub lstrcpynW _
        Lib "kernel32" ( _
        ByVal vDest As Any, _
        ByVal vSrc As Any, _
    lLength As Long)
    
    Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
        Alias "GetDiskFreeSpaceExA" _
        (ByVal lpDirName As String, _
        lpBytesToCall As Currency, _
        lpTotalBytes As Currency, _
        lpFreeBytes As Currency) _
        As Long
    
    Public Declare Sub CopyMemory Lib "kernel32" _
        Alias "RtlMoveMemory" _
        (hpvDest As Any, _
        hpvSource As Any, _
        ByVal cbCopy As Long)
    
    
    Public Declare Function NetApiBufferFree _
        Lib "netapi32.dll" ( _
        ByVal lpBuffer As Long) _
        As Long
    
    Declare Function NetWkstaGetInfo _
        Lib "netapi32.dll" ( _
        ByVal sServerName$, _
        ByVal lLevel&, _
        vBuffer As Any) _
        As Long
   
    Type SERVER_INFO_100
        sv100_platform_id As Long
        sv100_servername As Long
    End Type
    
    Public Type SERVER_INFO_101
        dw_platform_id As Long
        ptr_name As Long
        dw_ver_major As Long
        dw_ver_minor As Long
        dw_type As Long
        ptr_comment As Long
    End Type
    
    Type WKSTA_INFO_100
        wki100_platform_id As Long
        wki100_computername As Long
        wki100_langroup As Long
        wki100_ver_major As Long
        wki100_ver_minor As Long
    End Type
        
        Public Const ERROR_EXTENDED_ERROR = 1208&
        Public Const ERROR_NO_MORE_ITEMS = 259&
        Public Const ERROR_MORE_DATA = 234

        Public Const ERROR_NONE = 0&
        Public Const NOERROR = 0
        Public Const NO_ERROR = 0
        Public Const ERROR_SUCCESS = 0&
        Public Const NERR_Success As Long = 0&
        
        ' More of the SV_TYPE constants are here so that
        ' the ListServers call can be used to enumate
        ' different types of servers in other functions
        
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
    
    Public Function GetLocalSystemName()
        Dim lReturnCode As Long
        Dim bBuffer(512) As Byte
        Dim i As Integer
        Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
        Dim lwkstaInfo100StructPtr As Long
        Dim sLocalName As String
        
        ' Retrieves the local system name
        lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
     
        lwkstaInfo100StructPtr = lwkstaInfo100
                     
        If lReturnCode = 0 Then
                     
            RtlMoveMemory twkstaInfo100, ByVal _
            lwkstaInfo100StructPtr, Len(twkstaInfo100)
             
            lstrcpyW bBuffer(0), twkstaInfo100.wki100_computername
    
            i = 0
            Do While bBuffer(i) <> 0
                sLocalName = sLocalName & Chr(bBuffer(i))
                i = i + 2
            Loop
                
            GetLocalSystemName = sLocalName
             
        End If
    
    End Function
    
    Public Function GetDomainName() As String
        
        Dim lReturnCode As Long
        Dim bBuffer(512) As Byte
        Dim i As Integer
        Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
        Dim lwkstaInfo100StructPtr As Long
        Dim sDomainName As String
        
        ' Retrieves the Domain Name
        lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
     
        lwkstaInfo100StructPtr = lwkstaInfo100
                     
        If lReturnCode = 0 Then
                     
            RtlMoveMemory twkstaInfo100, ByVal lwkstaInfo100StructPtr, Len(twkstaInfo100)
             
            lstrcpyW bBuffer(0), twkstaInfo100.wki100_langroup

            i = 0
            Do While bBuffer(i) <> 0
                sDomainName = sDomainName & Chr(bBuffer(i))
                i = i + 2
            Loop
                
            GetDomainName = sDomainName
             
        End If
            
    End Function
    
    
    
