VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CServer"
Attribute VB_Creatable = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object of type RESOURCEDISPLAYTYPE_SERVER. A server."
Option Explicit

Dim Info            As NETRESOURCE
Dim typInfo         As SERVER_INFO_101
Dim lServerType     As Long
Dim bHaveServerInfo As Boolean

Dim colEventLogs    As Collection
Dim colGroups       As Collection
Dim colLocalGroups  As Collection
Dim colOpenFiles    As Collection
Dim colSchedule     As Collection
Dim colServices     As Collection
Dim colShares       As Collection
Dim colUsers        As Collection

Private Type FILE_INFO_3
    ID              As Long
    Permissions     As Long
    NumLocks        As Long
    PathName        As Long
    UserName        As Long
End Type

Private Type FILE_INFO_3_STRING
    ID              As Long
    Permissions     As Long
    NumLocks        As Long
    PathName        As String
    UserName        As String
End Type

'    LPTSTR servername,
'    LPTSTR basepath,
'    LPTSTR username,
'    DWORD level,
'    LPBYTE *bufptr,
'    DWORD prefmaxlen,
'    LPDWORD entriesread,
'    LPDWORD totalentries,
'    LPDWORD resume_handle
Private Declare Function NetFileEnum Lib "netapi32.dll" _
   (ByVal ServerName As Byte, _
    ByVal BasePath As Any, _
    ByVal UserName As Any, _
    ByVal Level As Long, _
    bufptr As Any, _
    ByVal PreferredMaxLen As Long, _
    EntriesRead As Long, _
    TotalEntries As Long, _
    ResumeHandle As Long) As Long

Private Type SERVER_INFO_101
    Platform_ID     As Long 'DWORD     sv101_platform_id
    Name            As Long 'LPTSTR    sv101_name
    VersionMajor    As Long 'DWORD     sv101_version_major;
    VersionMinor    As Long 'DWORD     sv101_version_minor;
    Type            As Long 'DWORD     sv101_type
    Comment         As Long 'LPTSTR    sv101_comment;
End Type

'    LPTSTR servername,
'    DWORD level,
'    LPBYTE *bufptr
Private Declare Function NetServerGetInfo Lib "netapi32" _
   (ServerName As Byte, _
    ByVal Level As Long, _
    bufptr As Byte) As Long
    
Private Type Local_Group_Info_1
    Name    As Long
    Comment As Long
End Type

'    LPWSTR servername,
'    DWORD level,
'    LPBYTE *bufptr,
'    DWORD prefmaxlen,
'    LPDWORD entriesread,
'    LPDWORD totalentries,
'    LPDWORD ResumeHandle
Private Declare Function NetLocalGroupEnum Lib "netapi32" _
   (ServerName As Byte, _
    ByVal Level As Long, _
    bufptr As Any, _
    ByVal MaxLen As Long, _
    EntriesRead As Long, _
    TotalEntries As Long, _
    ResumeHandle As Long) As Long
    
Const SV_PLATFORM_ID_OS2 = &H400
Const SV_PLATFORM_ID_NT = &H500
    
Const SV_TYPE_WORKSTATION = &H1
Const SV_TYPE_SERVER = &H2
Const SV_TYPE_SQLSERVER = &H4
Const SV_TYPE_DOMAIN_CTRL = &H8
Const SV_TYPE_DOMAIN_BAKCTRL = &H10
Const SV_TYPE_TIME_SOURCE = &H20
Const SV_TYPE_AFP = &H40
Const SV_TYPE_NOVELL = &H80
Const SV_TYPE_DOMAIN_MEMBER = &H100
Const SV_TYPE_PRINTQ_SERVER = &H200
Const SV_TYPE_DIALIN_SERVER = &H400
Const SV_TYPE_XENIX_SERVER = &H800
Const SV_TYPE_SERVER_UNIX = SV_TYPE_XENIX_SERVER
Const SV_TYPE_NT = &H1000
Const SV_TYPE_WFW = &H2000
Const SV_TYPE_SERVER_MFPN = &H4000
Const SV_TYPE_SERVER_NT = &H8000
Const SV_TYPE_POTENTIAL_BROWSER = &H10000
Const SV_TYPE_BACKUP_BROWSER = &H20000
Const SV_TYPE_MASTER_BROWSER = &H40000
Const SV_TYPE_DOMAIN_MASTER = &H80000
Const SV_TYPE_SERVER_OSF = &H100000
Const SV_TYPE_SERVER_VMS = &H200000
Const SV_TYPE_WINDOWS = &H400000                '/* Windows95 and above */
Const SV_TYPE_DFS = &H800000                    '/* Root of a DFS tree */
Const SV_TYPE_ALTERNATE_XPORT = &H20000000      '/* return list for alternate transport */
Const SV_TYPE_LOCAL_LIST_ONLY = &H40000000      '/* Return local list only */
Const SV_TYPE_DOMAIN_ENUM = &H80000000
Const SV_TYPE_ALL = &HFFFF                      '/* handy for NetServerEnum2 */

Private Declare Function NetScheduleJobEnum Lib "netapi32.dll" _
   (ByVal ServerName As Byte, _
    PointerToBuffer As Any, _
    ByVal PreferredMaximumLength As Long, _
    EntriesRead As Long, _
    TotalEntries As Long, _
    ResumeHandle As Long) As Long

Private Type AT_ENUM
    JobID       As Long
    JobTime     As Long
    DaysOfMonth As Long
    DaysOfWeek  As Byte
    Flags       As Byte
    Command     As Long
End Type

Private Type ENUM_SERVICE_STATUS
    lpServiceName As Long
    lpDisplayName As Long
    ServiceStatus As SERVICE_STATUS
End Type

Private Declare Function EnumServicesStatus& Lib "advapi32.dll" Alias "EnumServicesStatusA" _
   (ByVal hSCManager As Long, _
    ByVal dwServiceType As Long, _
    ByVal dwServiceState As Long, _
    lpServices As ENUM_SERVICE_STATUS, _
    ByVal cbBufSize As Long, _
    pcbBytesNeeded As Long, _
    lpServicesReturned As Long, _
    lpResumeHandle As Long)


Sub Enumerate()
Attribute Enumerate.VB_Description = "This method builds a collection of CShare objects representing the share names on the server, which can then be retrieved using either the Share or the Shares methods."
    Set colShares = EnumerateNR(Info)
End Sub
Property Get IsBackupBrowser() As Boolean
Attribute IsBackupBrowser.VB_Description = "True if the Server is a Backup Browser in its domain."
    IsBackupBrowser = (typInfo.Type And SV_TYPE_BACKUP_BROWSER) > 0
End Property

Property Get IsMasterBrowser() As Long
Attribute IsMasterBrowser.VB_Description = "True if the Server is the Master Browser in its domain."
    IsMasterBrowser = (typInfo.Type And SV_TYPE_MASTER_BROWSER) > 0
End Property
Property Get Name() As String
Attribute Name.VB_Description = "The Server name."
    Name = Info.lpRemoteName
End Property

Property Get OpenFiles() As Collection
    If colOpenFiles.Count = 0 Then _
        pGetOpenFiles
    Set OpenFiles = colOpenFiles
End Property

Property Get IsBDC() As Boolean
Attribute IsBDC.VB_Description = "True if this server is an NT Backup Domain controller (BDC).  Implemented by the NetServerGetInfo NetApi function. "
    If Not bHaveServerInfo Then
        IsBDC = False
    Else
        IsBDC = (typInfo.Type And SV_TYPE_DOMAIN_BAKCTRL) > 0
    End If
End Property

Property Get IsBrowser() As Boolean
Attribute IsBrowser.VB_Description = "True if the server is either a master or a backup browser."
    IsBrowser = ((typInfo.Type And SV_TYPE_BACKUP_BROWSER) > 0) _
             Or ((typInfo.Type And SV_TYPE_MASTER_BROWSER) > 0)
End Property

Property Get IsNovell() As Boolean
Attribute IsNovell.VB_Description = "True if the server is a Novell server."
    IsNovell = (typInfo.Type And SV_TYPE_NOVELL) > 0
End Property
Property Get IsNT() As Boolean
Attribute IsNT.VB_Description = "True if this server is either an NT Server or an NT Workstation."
    IsNT = (typInfo.Type And SV_TYPE_NT) > 0
End Property

Property Get IsNTServer() As Boolean
Attribute IsNTServer.VB_Description = "True if this server is an NT Server."
    IsNTServer = (typInfo.Type And SV_TYPE_SERVER_NT) > 0
End Property


Property Get IsPDC() As Boolean
Attribute IsPDC.VB_Description = "True if this server is the NT Primary Domain Controller (PDC) in the servers domain."
    If Not bHaveServerInfo Then
        IsPDC = False
    Else
        IsPDC = (typInfo.Type And SV_TYPE_DOMAIN_CTRL) > 0
    End If
End Property
Property Get IsSQLServer() As Boolean
Attribute IsSQLServer.VB_Description = "True if this server is a SQL Server."
    If Not bHaveServerInfo Then
        IsSQLServer = False
    Else
        IsSQLServer = (typInfo.Type And SV_TYPE_SQLSERVER) > 0
    End If
End Property
Property Get IsWFW() As Boolean
Attribute IsWFW.VB_Description = "True if this a Windows for Workgroups computer functioning as a server."
    IsWFW = (typInfo.Type And SV_TYPE_WFW) > 0
End Property
Property Get IsNTWorkstation() As Boolean
Attribute IsNTWorkstation.VB_Description = "True if this server is an NT Workstation."
    If Not bHaveServerInfo Then
        IsNTWorkstation = False
    Else
        IsNTWorkstation = (typInfo.Type And SV_TYPE_WORKSTATION) > 0
    End If
End Property

Private Function LocalGroups() As Collection
Attribute LocalGroups.VB_Description = "NOT YET IMPLEMENTED"
    Dim Result          As Long
    Dim bufptr          As Long
    Dim EntriesRead     As Long
    Dim TotalEntries    As Long
    Dim ResumeHandle    As Long
    Dim BufLen          As Long
    Dim SNArray()       As Byte
    Dim GNArray(99)     As Byte
    Dim GName           As String
    Dim i               As Integer
    Dim UNPtr           As Long
    Dim TempPtr         As MungeLong
    Dim TempStr         As MungeInt
    
    Dim colGroups       As Collection
    Dim oGroup          As CGroup
 
    SNArray = Info.lpRemoteName & vbNullChar       ' Move to byte array
    BufLen = 1024                      ' Buffer size
    ResumeHandle = 0                   ' Start with the first entry
    
    Do
        Result = NetLocalGroupEnum(SNArray(0), 1, bufptr, BufLen, _
                    EntriesRead, TotalEntries, ResumeHandle)
        
        If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads
                                             ' required
            Debug.Print "Error " & Result & " enumerating local group " & _
                        EntriesRead & " of " & TotalEntries
            Exit Function
        End If
        
        Set colGroups = New Collection
        
        For i = 1 To EntriesRead
            ' Get pointer to string from beginning of buffer
            ' Copy 4 byte block of memory in 2 steps
            Result = PtrToInt(TempStr.XLo, bufptr + (i - 1) * 4, 2)
            Result = PtrToInt(TempStr.XHi, bufptr + (i - 1) * 4 + 2, 2)
            LSet TempPtr = TempStr ' munge 2 Integers to a Long
            
            ' Copy string to array and convert to a string
            Result = PtrToStr(GNArray(0), TempPtr.x)
            GName = Left(GNArray, StrLen(TempPtr.x))
            
            Result = PtrToInt(TempStr.XLo, bufptr + (i - 1) * 4, 2)
            Result = PtrToInt(TempStr.XHi, bufptr + (i - 1) * 4 + 2, 2)
            LSet TempPtr = TempStr ' munge 2 Integers to a Long
            
            Result = PtrToStr(GNArray(0), TempPtr.x)
            GName = Left(GNArray, StrLen(TempPtr.x))
            
            Set oGroup = New NOblets.CGroup
            If oGroup.Init(sNetGroup:=GName, sServer:=Info.lpRemoteName, _
                    bGlobal:=False) Then
                colGroups.Add oGroup, Key:=oGroup.GroupName
            End If
        Next i
    Loop Until EntriesRead = TotalEntries
    ' The above condition only valid for reading accounts on NT
    ' but not OK for OS/2 or LanMan
    
    Result = NetApiBufferFree(bufptr)         ' Don't leak memory
    
    Set LocalGroups = colGroups
 
End Function

Private Sub pGetOpenFiles()
    Dim oFile           As COpenFile
    Dim bytServer()     As Byte
    Dim bytBuf()        As Byte
    Dim bufptr          As Long
    Dim lRC             As Long
    Dim lPtr            As Long
    Dim lEntriesRead    As Long
    Dim lTotalEntries   As Long
    Dim lResumeHandle   As Long
    
    Exit Sub
    
    ReDim bytBuf(16384) As Byte
    bufptr = VarPtr(bytBuf(0))
    
    bytServer() = Info.lpRemoteName
    
    ' *************************************
    ' DOES NOT WORK - HANGS, GPFs ON BREAK
    ' *************************************
    lRC = NetFileEnum(bytServer(0), 0&, 0&, 3&, bufptr, 16384, _
            lEntriesRead, lTotalEntries, lResumeHandle)
    If lRC <> 0 Then Exit Sub
    
    
End Sub


Private Sub pGetSchedule()
    Dim AtEnum      As AT_ENUM
    Dim lngBufSize  As Long
    Dim lRC         As Long
    Dim lngNumRead  As Long
    Dim lngNumTot   As Long
    Dim hResume     As Long
    Dim bytServer() As Byte
    Dim bytBuf()    As Byte
    Dim ptrBuff     As Long

    bytServer = Info.lpRemoteName
    
    lngBufSize = LenB(AtEnum) * 256

    lRC = NetScheduleJobEnum(ByVal bytServer(0), ptrBuff, lngBufSize, lngNumRead, lngNumTot, hResume)
    If lRC = 0 Then Exit Sub

    
End Sub

Private Function pGetServerInfo() As Boolean
    Dim lRC     As Long
    Dim bytServer()     As Byte
    Dim bytBuffer()     As Byte
    Dim lBufLen         As Long
    Dim lpBuf           As Long
    
    lBufLen = 1024
    ReDim bytBuffer(lBufLen)
    
    bytServer = Info.lpRemoteName & vbNullChar
    lRC = NetServerGetInfo(bytServer(0), 101&, bytBuffer(0))
    If lRC = 0 Then
        pGetServerInfo = True
        CopyMemory lpBuf, VarPtr(bytBuffer(0)), Len(lpBuf)
        CopyMemory typInfo, lpBuf, Len(typInfo)
    Else
        pGetServerInfo = False
    End If
End Function

Property Get GlobalGroups() As Collection
Attribute GlobalGroups.VB_Description = "This method returns a collection of CGroup objects representing the global groups defined at this server.  Valid on domain controllers only."
    If colGroups.Count = 0 Then
        Set colGroups = EnumerateGroups(Info.lpRemoteName, "")
    End If
    Set GlobalGroups = colGroups
End Property

Private Function InitOn(sNetworkName As String) As CShare
    Set InitOn = pInitOn(Info, sNetworkName)
End Function

Property Get Comment() As String
Attribute Comment.VB_Description = "The comment, or description, for this server."
    Comment = Info.lpComment
End Property

Property Get DiskShares() As Collection
Attribute DiskShares.VB_Description = "Returns the subcollection of Shares of ShareType RESOURCETYPE_FILE."
    Dim colDisks    As New Collection
    Dim oShare      As CShare
    For Each oShare In colShares
        If oShare.ShareType = RESOURCETYPE_DISK Then
            colDisks.Add oShare
        End If
    Next
    Set DiskShares = colDisks
End Property

Property Get LocalName() As String
Attribute LocalName.VB_Description = "The LocalName returned for this server by the WNetEnum functions.  For a server, this property is NULL at present."
    LocalName = Info.lpLocalName
End Property

Property Get Platform() As Long
Attribute Platform.VB_Description = "The platform to use for subsequent NETAPI calls.  400 for OS/2 based servers,  500 for NT based servers."
    Platform = typInfo.Platform_ID
End Property

Property Get PrintShares() As Collection
Attribute PrintShares.VB_Description = "Returns the subcollection of Shares of ShareType RESOURCETYPE_PRINTER."
    Dim colPrinters As New Collection
    Dim oShare      As CShare
    For Each oShare In colShares
        If oShare.ShareType = RESOURCETYPE_PRINT Then
            colPrinters.Add oShare
        End If
    Next
    Set PrintShares = colPrinters
End Property

Property Get Provider() As String
Attribute Provider.VB_Description = "Returns the network provider through which this server is known to the network."
    Provider = Info.lpProvider
End Property

Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the Shares collection, setting it to Nothing and then rebuilding it."
    Set colShares = Nothing
    Set colShares = EnumerateNR(Info)
End Sub

Property Get RemoteName() As String
Attribute RemoteName.VB_Description = "The remote name of this server."
    RemoteName = Info.lpRemoteName
End Property
Property Get Schedule() As Collection
    If colSchedule.Count = 0 Then _
        pGetSchedule
    Set Schedule = colSchedule
End Property

Property Get ServerType() As Long
Attribute ServerType.VB_Description = "The Type property returned from NetServerGetInfo.  Use the SV_TYPE_ constants to AND and get various server types."
    ServerType = typInfo.Type
End Property

Property Get Service(vService As Variant) As CService
Attribute Service.VB_Description = "Finds and returns CService object indicated by vService.  vService can be either the index or the key of the object in the Services collection."
    Set Service = colServices(vService)
End Property

Property Get Share(vShareName As Variant) As CShare
Attribute Share.VB_Description = "Finds and returns CShare object indicated by vShare.  vShare can be either the index or the key of the object in the CServers Shares collection."
    Set Share = colShares(vShareName)
End Property

Property Get Shares() As Collection
Attribute Shares.VB_Description = "Returns a collection of CShare objects, which are resources local to the server which are made available for use by other computers."
    Set Shares = colShares
End Property

Property Get Users() As Collection
Attribute Users.VB_Description = "Returns a collection of CUser objects.  If the server is a domain controller, the users are domain users.  If the server is not a DC, the users are members of the servers local user database."
    If colUsers.Count = 0 Then
        Set colUsers = EnumerateUsers(Info.lpRemoteName, "")
    End If
    Set Users = colUsers
End Property

Property Get VerMajor() As Long
    VerMajor = typInfo.VersionMajor
End Property
Property Get VerMinor() As Long
    VerMinor = typInfo.VersionMinor
End Property


Property Get Version() As String
Attribute Version.VB_Description = "The concatenation of the Version_Major and Version_Minor properties returned by NetServerGetInfo NetApi function."
    Version = CStr(typInfo.VersionMajor) & "." & CStr(typInfo.VersionMinor)
End Property
    
Private Sub Class_Initialize()
    Set colShares = New Collection
    Set colUsers = New Collection
    Set colLocalGroups = New Collection
    Set colGroups = New Collection
    Set colServices = New Collection
    Set colOpenFiles = New Collection
    Set colSchedule = New Collection
End Sub

Private Sub Class_Terminate()
    Set colShares = Nothing
    Set colUsers = Nothing
    Set colGroups = Nothing
    Set colLocalGroups = Nothing
    Set colServices = Nothing
    Set colOpenFiles = Nothing
    Set colSchedule = Nothing
End Sub

Sub Init(Scope As Long, dwType As Long, DisplayType As Long, Usage As Long, _
    LocalName As String, RemoteName As String, Comment As String, Provider As String)
Attribute Init.VB_Description = "Initializes the CServer object with parameters from a NETRESOURCE type."

    With Info
        .dwScope = Scope
        .dwType = dwType
        .dwDisplayType = DisplayType
        .dwUsage = Usage
        .lpLocalName = LocalName
        .lpRemoteName = RemoteName
        .lpComment = Comment
        .lpProvider = Provider
    End With
    
    bHaveServerInfo = pGetServerInfo()
End Sub

Property Get Services() As Collection
Attribute Services.VB_Description = "Returns a collection of CService objects, representing the services on the Server, both active and inactive."
    If colServices.Count = 0 Then _
        Call pGetServices
    Set Services = colServices
End Property

Private Sub pGetServices()
    Dim lRC                 As Long
    Dim hSC                 As Long
    Dim cBuf                As Long
    Dim bufEnum()           As ENUM_SERVICE_STATUS
    Dim pcBytesNeeded       As Long
    Dim lServicesReturned   As Long
    Dim lResumeHandle       As Long
    
    Dim sServiceName        As String
    Dim sDisplayName        As String
    Dim iCnt                As Integer
    Dim oService            As CService
    Dim oErrMsg             As New CErrMsgNT

    ReDim bufEnum(1 To 100)
    cBuf = LenB(bufEnum(1)) * 100
    
    hSC = OpenSCManager(Info.lpRemoteName, "ServicesActive", SC_MANAGER_ENUMERATE_SERVICE)
    If hSC = 0 Then
        Debug.Print oErrMsg.ErrMsg(Err.LastDllError)
        Exit Sub
    End If
    
    lRC = EnumServicesStatus(ByVal hSC, ByVal SERVICE_WIN32, ByVal SERVICE_STATE_ALL, _
        bufEnum(1), ByVal cBuf, pcBytesNeeded, lServicesReturned, lResumeHandle)
    
    If lRC = 0 Then
        Exit Sub
    End If
    
    For iCnt = 1 To lServicesReturned
        Set oService = New CService
        sServiceName = StrFromPtr(bufEnum(iCnt).lpServiceName, 32&)
        sDisplayName = StrFromPtr(bufEnum(iCnt).lpDisplayName, 32&)
        If oService.Init(Info.lpRemoteName, sServiceName, sDisplayName, _
                bufEnum(iCnt).ServiceStatus.dwCurrentState) Then
            colServices.Add oService, sServiceName
        End If
'        Debug.Print sServiceName, sDisplayName, _
            StatusStr(bufEnum(iCnt).ServiceStatus.dwCurrentState - 1)
    Next iCnt
    
    lRC = CloseServiceHandle(hSC)
End Sub
