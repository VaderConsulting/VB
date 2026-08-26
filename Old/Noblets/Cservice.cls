VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CService"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Attribute VB_Description = "An encapsulation of a NT Server service.  The CService objects are members of a CServer's Services collection."
Option Explicit

Dim pbGotInfo       As Boolean

Dim fhSC            As Long     ' OpenSCManager handle

Dim psName          As String
Dim psDisplayName   As String
Dim psServer        As String
Dim plStatus        As Long

' QueryServiceConfig info
Dim plServiceType   As Long
Dim plStartType     As Long
Dim plErrorControl  As Long
Dim psBinaryPath    As String
Dim psStartUser     As String

Private Type QUERY_SERVICE_CONFIG
    dwServiceType       As Long
    dwStartType         As Long
    dwErrorControl      As Long
    lpBinaryPathName    As String
    lpLoadOrderGroup    As String
    dwTagId             As Long
    lpDependencies      As String
    lpServiceStartName  As String
    lpDisplayName       As String
End Type

Private Type QUERY_SERVICE_LOCK_STATUS
    fIsLocked       As Long
    lpLockOwner     As String
    dwLockDuration  As Long
End Type

Private Declare Function QueryServiceStatus Lib "advapi32.dll" _
   (ByVal hService As Long, _
    lpServiceStatus As SERVICE_STATUS) As Long

Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" _
   (ByVal hSCManager As Long, _
    ByVal lpServiceName As String, _
    ByVal dwDesiredAccess As Long) As Long

'Private Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" _
   (ByVal hService As Long, _
    lpServiceConfig As QUERY_SERVICE_CONFIG, _
    ByVal cbBufSize As Long, _
    pcbBytesNeeded As Long) As Long

Private Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" _
   (ByVal hService As Long, _
    lpServiceConfig As Any, _
    ByVal cbBufSize As Long, _
    pcbBytesNeeded As Long) As Long

Private Const STR_SERVICES_ACTIVE = "ServicesActive"

Property Get DisplayName() As String
Attribute DisplayName.VB_Description = "The longer display name of the Service."
    DisplayName = psDisplayName
End Property

Property Get EXEName() As String
Attribute EXEName.VB_Description = "The name of the binary executable file that implements the Service."
    EXEName = psBinaryPath
End Property

Function Init(sServer As String, sName As String, sDisplayName, lStatus As Long) As Boolean
Attribute Init.VB_Description = "Initializes a CService object representing the state of the sName service on server sServer.  A dependent object, created by a CServer object."
    
    Init = False
    
    If Len(Trim$(sName)) = 0 Then Exit Function
        
    psServer = sServer
    psName = sName
    psDisplayName = sDisplayName
    plStatus = lStatus
    
    fhSC = OpenSCManager(psServer, STR_SERVICES_ACTIVE, SC_MANAGER_CONNECT)
    If fhSC = 0 Then Exit Function
        
    Call pGetInfo
    
    Init = True
End Function

Property Get Name() As String
Attribute Name.VB_Description = "The keyword name of the Service."
    Name = psName
End Property


Private Sub pGetInfo()
    Dim lRC             As Long
    Dim hSCM            As Long
    Dim hService        As Long
    
    Dim bytBuf()        As Byte
    Dim lBufLen         As Long
    
    Dim typQueryConfig  As QUERY_SERVICE_CONFIG
    Dim lBytesNeeded    As Long

    Dim sDepends        As String
    Dim nDepends        As String
    Dim i               As Integer
    
    hService = OpenService(fhSC, psName, SERVICE_QUERY_CONFIG)
    If hService = 0 Then Exit Sub
    
    lBufLen = 256
    ReDim bytBuf(1 To lBufLen) As Byte
    
    lRC = QueryServiceConfig(hService, bytBuf(1), _
                lBufLen, lBytesNeeded)
    
    CopyMemory plServiceType, VarPtr(bytBuf(1)), Len(plServiceType)
    
    CopyMemory plStartType, VarPtr(bytBuf(5)), Len(plStartType)
    
    CopyMemory plErrorControl, VarPtr(bytBuf(9)), Len(plErrorControl)
    
    CopyMemory lRC, VarPtr(bytBuf(13)), Len(lRC)    ' BinaryPath
    psBinaryPath = StrFromPtr(lRC, lBufLen)
    
    CopyMemory lRC, VarPtr(bytBuf(25)), Len(lRC)    ' Dependencies
    
    CopyMemory lRC, VarPtr(bytBuf(29)), Len(lRC)    ' StartUser
    psStartUser = StrFromPtr(lRC, 32&)
    
    CloseServiceHandle hService
End Sub

Private Function pGetStatus() As Boolean
    Dim lRC         As Long
    Dim hService    As Long
    Dim typSS       As SERVICE_STATUS
    
    pGetStatus = False

    hService = OpenService(fhSC, psName, SERVICE_QUERY_STATUS)
    If hService = 0 Then Exit Function

    lRC = QueryServiceStatus(hService, typSS)
    If lRC = 0 Then Exit Function
    
    plStatus = typSS.dwCurrentState
    CloseServiceHandle hService
    pGetStatus = True
End Function

Property Get Server() As String
Attribute Server.VB_Description = "The name of the Server on which this Service instance is running."
    Server = psServer
End Property

Property Get StartTypeString() As String
    StartTypeString = vStartTypeStr(plStartType)
End Property

Property Get StatusString() As String
Attribute StatusString.VB_Description = "The status of the service, returned as a string (""Running"", ""Stopped"", etc.)"
    pGetStatus
    StatusString = vStatusStr(plStatus - 1)
End Property

Property Get StartType() As Long
Attribute StartType.VB_Description = "The startup type of the Service (Automatic, Manual, Boot, etc.)"
    StartType = plStartType
End Property
Property Get StartUser() As String
Attribute StartUser.VB_Description = "The user name with which the Service starts.  If NULL, the Service starts as the LocalSystem account."
    StartUser = psStartUser
End Property


Property Get Status() As Long
Attribute Status.VB_Description = "The present status of the Service (RUNNING, STOPPED, PENDING, etc.)  When invoked, this function calls QueryServiceStatus to get the real-time status."
    pGetStatus
    Status = plStatus
End Property

Private Sub Class_Initialize()
    pbGotInfo = False
    If Not bServicesInit Then Call SetServiceStrings
End Sub

Private Sub Class_Terminate()
    CloseServiceHandle fhSC
End Sub


