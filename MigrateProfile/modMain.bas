Attribute VB_Name = "modMain"
Private Const CONNECT_UPDATE_PROFILE = &H1
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_REMEMBERED = &H3
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Private Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Private Const RESOURCEDISPLAYTYPE_SERVER = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE = &H3
Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2

Private Const WN_Success = &H0
Private Const WN_Not_Supported = &H1
Private Const WN_Net_Error = &H2
Private Const WN_Bad_Pointer = &H4
Private Const WN_Bad_NetName = &H32
Private Const WN_Bad_Password = &H6
Private Const WN_Bad_Localname = &H33
Private Const WN_Access_Denied = &H7
Private Const WN_Out_Of_Memory = &HB
Private Const WN_Already_Connected = &H34
Private Const WN_Server_Down = &H35
Private Const WN_Server_Down_2 = &H52E
Private Const WN_NOT_FOUND = &H43
Private Const WN_IN_USE = 85
Private Const WN_NOT_CONNECTED = 2250

'Private Consts for return codes errors
Private Const ERROR_NO_ERROR As Long = 0
Private Const ERROR_ALREADY_ASSIGNED As Long = 85
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_BAD_DEVICE_TYPE As Long = 66
Private Const ERROR_BAD_NET_NAME As Long = 67
Private Const ERROR_BAD_PROFILE As Long = 1206
Private Const ERROR_BAD_DEVICE As Long = 1200
Private Const ERROR_BAD_PROVIDER As Long = 1204
Private Const ERROR_BUSY As Long = 170
Private Const ERROR_CANCEL_VIOLATION As Long = 173
Private Const ERROR_CANNOT_OPEN_PROFILE As Long = 1205
Private Const ERROR_DEVICE_ALREADY_REMEMBERED As Long = 1202
Private Const ERROR_EXTENDED_ERROR As Long = 1208
Private Const ERROR_INVALID_PASSWORD As Long = 86
Private Const ERROR_NO_NET_OR_BAD_PATH As Long = 1203
Private Const ERROR_SESSION_CREDENTIAL_CONFLICT As Long = 1219
Private Const ERROR_NO_NETWORK As Long = 1222
Private Const ERROR_CANCELLED As Long = 1223
Private Const ERROR_NO_CONNECTION As Long = 8
Private Const ERROR_NO_DISCONNECT As Long = 9
Private Const ERROR_DEVICE_IN_USE As Long = 2404
Private Const ERROR_NOT_CONNECTED As Long = 2250
Private Const ERROR_OPEN_FILES As Long = 2401
Private Const ERROR_MORE_DATA As Long = 234

Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

Public strHTML As String
Public oFSO As Scripting.FileSystemObject
Public oDrive As Scripting.Drive
Public oFolder As Scripting.Folder
Public oFile As Scripting.File

Public oEvents As New Collection
Public bReportErrors As Boolean
Public bReportWarnings As Boolean
Public bReportInformation As Boolean
Public bReportToEventLog As Boolean
Public strApp_Path As String
Public strAppStatus As String

Public bWriteDebugMessagesToEventLog As Boolean

Public strAppPath As String
Public strDomainName As String
Public strTempDir As String

Public Enum EventLogError
    Error = vbLogEventTypeError                ' 1
    Warning = vbLogEventTypeWarning            ' 2
    Information = vbLogEventTypeInformation    ' 4
    DebugInfo = 8
End Enum

Public Enum ApplicationStatus
    statusUnknown = 0
    StatusError = 1
    statusWarning = 2
    statusOK = 4
End Enum

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

Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public lpNetResource As NETRESOURCE

Public Declare Function AddPrinterConnection Lib "winspool.drv" Alias "AddPrinterConnectionA" (ByVal pName As String) As Long
Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'---------------------------------------------------------------------------------------
' Procedure : AddInfo
' DateTime  : 18-04-2003 18:29
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   18-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function AddInfo(strInfo As String, Optional HTMLString As String = "")
    If strInfo <> "" Then
        frmMain.lblInfo.Caption = strInfo
        frmMain.lblInfo.Refresh
    End If
    
    strHTML = strHTML & HTMLString
    
End Function

Public Function Connect2(ByVal Localname As String, ByVal RemoteName As String) As Long
    Dim ErrorNum As Long
    Dim ErrorMsg As String
    Dim Username As String
    Dim Password As String
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

    rc = WNetAddConnection2(lpNetResource, Username, Password, 0)
    Connect2 = rc
    If rc <> 0 Then GoTo Err_Connect

    Exit Function

Err_Connect:
    ErrorNum = rc
    ErrorMsg = WnetError(rc)
    Connect2 = rc

End Function

Public Sub GetDrives()
    AddInfo "Drives...", "<H2>Drives</H2>"
    
    frmMain.lstDrives.Clear
    
    For Each oDrive In oFSO.Drives
        Select Case oDrive.DriveType
            Case 0
                AddInfo "UNKNOWN: " & oDrive.DriveLetter & " (" & oDrive.FileSystem & ")", "UNKNOWN: " & oDrive.DriveLetter & " (" & oDrive.FileSystem & ")<br>"
            Case 1
                AddInfo "Removable Drive: " & oDrive.DriveLetter, "Removable Drive: " & oDrive.DriveLetter & "<br>"
            Case 2
                AddInfo "Fixed Drive: " & oDrive.DriveLetter & " (" & oDrive.FileSystem & ")", "Fixed Drive: " & oDrive.DriveLetter & " (" & oDrive.FileSystem & ")<br>"
            Case 3
                frmMain.lstDrives.AddItem oDrive.DriveLetter & " = " & oDrive.ShareName
                AddInfo "Network Drive: " & oDrive.DriveLetter & " = " & oDrive.ShareName, "Network Drive: " & oDrive.DriveLetter & " = " & oDrive.ShareName & "&nbsp;[WILL BE RECONNECTED]<br>"
            Case 4
                AddInfo "CD\DVD Drive: " & oDrive.DriveLetter, "CD\DVD Drive: " & oDrive.DriveLetter & "<br>"
            Case 5
                AddInfo "RAM Drive: " & oDrive.DriveLetter, "RAM Drive: " & oDrive.DriveLetter & "<br>"
        End Select
    Next
    AddInfo "", "<br>"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetDriveType
' DateTime  : 18-04-2003 18:30
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   18-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetDriveType(strDriveSpec As String) As Integer
    Dim oFSO As Scripting.FileSystemObject
    Dim oDrive As Drive
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Set oDrive = oFSO.GetDrive(strDriveSpec)
    
    GetDriveType = oDrive.DriveType
    
    Set oDrive = Nothing
    Set oFSO = Nothing
End Function

Public Sub GetHomePage()
    frmMain.txtHomePage.Text = GetRegString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page", "")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetPathfromUNC
' DateTime  : 14-03-2003 10:29
' Author    : Dave Robinson
' Purpose   : Retrieve the path or share or printer (etc) portion of the given UNC
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Function GetPathfromUNC(strUNCPath As String) As String
    Dim intTemp1 As Integer
    Dim strPath As String
    On Error GoTo GetPathfromUNC_Error

    If InStr(1, strUNCPath, "\\") = 0 Then Exit Function
    intTemp1 = InStr(3, Trim(strUNCPath), "\")
    strPath = Right(Trim(strUNCPath), Len(Trim(strUNCPath)) - intTemp1)
    GetPathfromUNC = UCase(strPath)

    On Error GoTo 0
    Exit Function

GetPathfromUNC_Error:

    LogEvent "**** Error " & Err.Number & " (" & Err.Description & ") in procedure GetPathfromUNC", Error
End Function

Public Sub GetPrinters()
    On Error GoTo Er
    
    AddInfo "Printers...", "<H2>Printers</H2>"
    
    frmMain.lstPrinters.Clear
    
    For Each oPrinter In Printers
        If Left(oPrinter.DeviceName, 2) = "\\" Then
            frmMain.lstPrinters.AddItem oPrinter.DeviceName
            'frmMain.lstPrinters.Selected(frmMain.lstPrinters.ListCount - 1) = True
            AddInfo "This printer will be reconnected if possible", "***The following printer will be reconnected if Drivers are present on the Server.<br>"
        Else
            'frmMain.lstPrinters.Selected(frmMain.lstPrinters.ListCount - 1) = False
        End If
        AddInfo oPrinter.DeviceName, oPrinter.DeviceName & "<br>"
    Next
    
    If Printer.DeviceName <> "" Then
        AddInfo "Default printer: " & Printer.DeviceName, "Default printer: " & Printer.DeviceName
        AddInfo "", "<br>"
    Else
        AddInfo "There are no printers defined on this system.", "There are no printers defined on this system.<br>"
        AddInfo "", "<br>"
    End If
    
    AddInfo "", "<hr>"
    Exit Sub
Er:
    Resume Next
End Sub

Public Sub GetProfileDir()
    frmMain.txtProfileDir.Text = Environ$("USERPROFILE")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetServerFromPath
' DateTime  : 14-03-2003 10:29
' Author    : Dave Robinson
' Purpose   : Retrieve the server portion of the given UNC
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Function GetServerFromPath(strUNCPath As String) As String
    Dim intTemp1 As Integer
    Dim strPath As String
    On Error GoTo GetServerFromPath_Error

    If InStr(1, strUNCPath, "\\") = 0 Then Exit Function
    intTemp1 = InStr(3, Trim(strUNCPath), "\")
    strPath = Left(Trim(strUNCPath), intTemp1 - 1)
    GetServerFromPath = UCase(Right(strPath, Len(strPath) - 2))

    On Error GoTo 0
    Exit Function

GetServerFromPath_Error:

    LogEvent "**** Error " & Err.Number & " (" & Err.Description & ") in procedure GetServerFromPath", Error
End Function

'---------------------------------------------------------------------------------------
' Procedure : LogEvent
' DateTime  : 18-04-2003 18:29
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author          History
' 1.0   18-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Sub LogEvent(strMessage As String, Optional HTMLString As String = "", Optional EventType As EventLogError = Information)
    Dim bDoReport As Boolean
    
    If bReportErrors And EventType = Error Then
        bDoReport = True
    End If
    
    If bReportWarnings And EventType = Warning Then
        bDoReport = True
    End If
    
    If bReportInformation And EventType = Information Then
        bDoReport = True
    End If
    
    If bReportInformation And EventType = Information Then
        bDoReport = True
    End If
    
    If bDebugMode And EventType = DebugInfo Then
        bDoReport = True
    End If
    
    '* Only add this event if we have determined it should be (according to the config options)
    If bDoReport Then
        oEvents.Add Format(Now, "DD/MM/YYYY|HH:NN:SS AM/PM") & "|" & EventType & "|" & strMessage
        strHTML = strHTML & HTMLString
        If (EventType <> Information) And EventType <> DebugInfo Then
            App.LogEvent strMessage, EventType
            SetStatus EventType
        ElseIf EventType = DebugInfo Then
            If bWriteDebugMessagesToEventLog Then
                App.LogEvent strMessage, vbLogEventTypeInformation
            End If
            SetStatus statusUnknown
        End If
    End If
End Sub

Public Sub SaveData()
    Dim i As Integer
    
    Open strAppPath & Environ$("USERNAME") & "-" & Environ$("COMPUTERNAME") & ".txt" For Output As #1
        Print #1, "Username      : " & Environ$("USERNAME")
        Print #1, "Computername  : " & Environ$("COMPUTERNAME")
        Print #1, "Domainname    : " & frmMain.txtDomainName
        Print #1, "Network Drives: " & frmMain.lstDrives.ListCount
        For i = 0 To frmMain.lstDrives.ListCount
            If Len(Trim(frmMain.lstDrives.List(i))) > 0 Then
                Print #1, "Drive         : " & frmMain.lstDrives.List(i)
            End If
        Next i
        Print #1, "Printers      : " & frmMain.lstPrinters.ListCount
        For i = 0 To frmMain.lstPrinters.ListCount
            If Len(Trim(frmMain.lstPrinters.List(i))) > 0 Then
                Print #1, "Printer       : " & frmMain.lstPrinters.List(i)
            End If
        Next i
        Print #1, "Home page     : " & frmMain.txtHomePage
        Print #1, "Profile Dir   : " & frmMain.txtProfileDir
    Close 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetStatus
' DateTime  : 14-03-2003 10:38
' Author    : Dave Robinson
' Purpose   : Set the application status
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------

Public Sub SetStatus(Status As ApplicationStatus)
    On Error GoTo SetStatus_Error

    Select Case Status
        Case ApplicationStatus.statusUnknown
            strAppStatus = "Unknown or Debug Mode"
        Case ApplicationStatus.statusOK
            strAppStatus = "OK"
        Case ApplicationStatus.StatusError
            strAppStatus = "Error"
        Case ApplicationStatus.statusWarning
            strAppStatus = "Warning"
    End Select

    On Error GoTo 0
    Exit Sub

SetStatus_Error:

    LogEvent "**** Error " & Err.Number & " (" & Err.Description & ") in procedure SetStatus of Module modGlobal", Error
End Sub

Public Function WnetError(Errcode As Long) As String

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
      Case 1203
        WnetError = "No network provider accepted the given network path"
      Case Else

        WnetError = "Unrecognized Error " + Str$(Errcode) + "."
    End Select

End Function

Public Function GetSP() As String
    ' Get operating system service pack.
    Dim VerInfo As OSVERSIONINFO
    Dim Ret%
    
    VerInfo.dwOSVersionInfoSize = Len(VerInfo)
    Ret% = GetVersionEx(VerInfo)
    If Ret% = 0 Then
        frmMain.lblInfo.Caption = "Error Getting Version Information"
    Else
        GetSP = Trim(Replace(VerInfo.szCSDVersion, Chr(0), ""))
    End If
    
End Function

Public Function GetOS() As String
    ' Get operating system.
    Dim VerInfo As OSVERSIONINFO
    Dim Ret%
    Dim strVersion As String
    
    VerInfo.dwOSVersionInfoSize = Len(VerInfo)
    Ret% = GetVersionEx(VerInfo)
    If Ret% = 0 Then
        frmMain.lblInfo.Caption = "Error Getting Version Information"
    Else
        strVersion = Trim(Replace(VerInfo.dwMajorVersion & VerInfo.dwMinorVersion, Chr(0), ""))
        Select Case strVersion
            Case "4.0"
                GetOS = "Windows NT 4" ' WHOA!  What the hell's happening here????.  Should be 2000+
            Case "5.0", "5.", "5"
                GetOS = "Windows 2000"
            Case "5.1"
                GetOS = "Windows XP"
            Case "5.2"
                GetOS = "Windows 2003"
        End Select
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetIniValue
' DateTime  : 03-04-2003 08:21
' Author    : Dave Robinson
' Purpose   : Retrieves a value from an ini file corresponding to the section and key name passed.
'
'  V    Date        Author          History
' 1.0   03-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetIniValue(lpSectionName As String, lpKeyName As String, defaultValue As String, inifile As String) As String
    Dim success As Long
    Dim nSize As Long
    Dim Ret As String
    
    On Error GoTo GetIniValue_Error
    
    'call the API with the parameters passed.
    'The return value is the length of the string in ret, including the terminating null. If a default value was passed, and the section or
    'key name are not in the file, that value is returned. If no default value was passed (""), then success will = 0 if not found.
    
    'Pad a string large enough to hold the data.
    Ret = Space$(2048)
    nSize = Len(Ret)
    success = GetPrivateProfileString(lpSectionName, lpKeyName, defaultValue, Ret, nSize, inifile)
    
    If success Then
        GetIniValue = Left$(Ret, success)
    End If
    
    On Error GoTo 0
    Exit Function
    
GetIniValue_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetIniValue of Module modMain", vbCritical

End Function

Public Sub Test()
    Dim oFileVersion As New cFileVersion
    Dim strFileVersion As String
    Dim l As Long
    
    l = oFileVersion.OpenFile(Environ$("SYSTEMROOT") & "\System32\Drivers\ipsec.sys")
    strFileVersion = oFileVersion.FileVersion
    strExistingIPSecversion = Mid(strFileVersion, InStrRev(strFileVersion, ".") + 1)
    
End Sub
