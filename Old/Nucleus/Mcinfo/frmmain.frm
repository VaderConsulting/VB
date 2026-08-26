VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envy"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SERVICE_STATUS
  dwServiceType As Long
  dwCurrentState As Long
  dwControlsAccepted As Long
  dwWin32ExitCode As Long
  dwServiceSpecificExitCode As Long
  dwCheckPoint As Long
  dwWaitHint As Long
End Type

Private Type QUERY_SERVICE_CONFIG
  dwServiceType As Long
  dwStartType As Long
  dwErrorControl As Long
  lpBinaryPathName As Long 'String
  lpLoadOrderGroup As Long ' String
  dwTagId As Long
  lpDependencies As Long 'String
  lpServiceStartName As Long 'String
  lpDisplayName As Long  'String
End Type

Private Const SERVICE_STOPPED = &H1
Private Const SERVICE_START_PENDING = &H2
Private Const SERVICE_STOP_PENDING = &H3
Private Const SERVICE_RUNNING = &H4
Private Const SERVICE_CONTINUE_PENDING = &H5
Private Const SERVICE_PAUSE_PENDING = &H6
Private Const SERVICE_PAUSED = &H7
Private Const SERVICE_ACCEPT_STOP = &H1
Private Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Private Const SERVICE_ACCEPT_SHUTDOWN = &H4
Private Const SC_MANAGER_CONNECT = &H1
Private Const SERVICE_INTERROGATE = &H80
Private Const GENERIC_READ = &H80000000
Private Const ERROR_INSUFFICIENT_BUFFER = 122

Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceConfig Lib "advapi32.dll" Alias "QueryServiceConfigA" (ByVal hService As Long, lpServiceConfig As Byte, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Sub Form_Load()
  'On Error Resume Next
  ' Registry stuff
  Dim Hostname As String, EngineVer As String, ProductVer As String, Serial As String
  Dim VirusDefDate As String, VirusDefVer As String, SQLString As String
  
  ' Service stuff
  Dim hSCM  As Long
  Dim hSVC As Long
  Dim pSTATUS As SERVICE_STATUS
  Dim udtConfig As QUERY_SERVICE_CONFIG
  Dim lRet As Long
  Dim lBytesNeeded As Long
  Dim sTemp As String
  Dim pFileName As Long, msg As String
  Dim Service As String, ServiceState As String
  
  ' DrWatson stuff
  Dim DrWatsonDate As Date
  
  ' AV Stuff
  Dim VirusesFound As Long, ScanDate As Date
  Dim DatDateTime As Date, CurrentDate As Date, DatFilename As String, OldCollection As Integer
  
  ' General variables
  Dim intTemp As Integer
  Dim strTemp As String
  Dim lngTemp As Long
  Dim boolTemp As Boolean
  
  ' Start here
  ' Get computer name
  Hostname = UCase(Environ$("computername"))
  ' Create registry key
  boolTemp = CreateKey(Hostname, "Software", "CSC")
  ' Create registry key
  boolTemp = CreateKey(Hostname, "Software\CSC", "Nucleus")
  ' Get current value of Collecting key
  strTemp = GetValue(Hostname, "Software\CSC\Nucleus", "Collecting", False)
  If Trim(strTemp) <> "" Then
    OldCollection = CInt(strTemp)
  Else
    OldCollection = 0
  End If
  ' Add one to it, and save it back so we know this app is running
  SetValueString Hostname, "Software\CSC\Nucleus", "Collecting", CStr(OldCollection + 1), REG_SZ
  
  ' Get version info from Registry
  EngineVer = Trim(GetValue(Hostname, "Software\McAfee\VirusScan", "szEngineVer", False))
  ProductVer = Trim(GetValue(Hostname, "Software\McAfee\VirusScan", "szProductVer", False))
  Serial = Trim(GetValue(Hostname, "Software\McAfee\VirusScan", "szSerialNum", False))
  VirusDefDate = Trim(GetValue(Hostname, "Software\McAfee\VirusScan", "szVirDefDate", False))
  VirusDefVer = Trim(GetValue(Hostname, "Software\McAfee\VirusScan", "szVirDefVer", False))
  
  ' Create directory to store information
  If Dir("C:\Nucleus", vbDirectory) = "" Then
    MkDir "C:\Nucleus"
  End If
  
  Open "C:\Nucleus\McInfo.log" For Output As #99
    Print #99, "Registry input complete"
  Close 99
  ' Open The Service Control Manager
  Service = "McShield"
  hSCM = OpenSCManager(Hostname, vbNullString, SC_MANAGER_CONNECT)
  If hSCM = 0 Then
    ' Host not found - NOT POSSIBLE for Localhost
  End If
  
  ' Open the specific Service to obtain a handle
  '
  hSVC = OpenService(hSCM, Trim(Service), GENERIC_READ)
  If hSVC = 0 Then
    'Debug.Print Hostname & "-" & Service & " does not exist"
    GoTo CloseHandles
  End If
  
  ' Fill the Service Status Structure
  '
  lRet = QueryServiceStatus(hSVC, pSTATUS)
  If lRet = 0 Then
    GoTo CloseHandles
  End If
  
  ' Report the Current State
  '
  Select Case pSTATUS.dwCurrentState
  Case SERVICE_STOPPED
    sTemp = "Stopped"
  Case SERVICE_START_PENDING
    sTemp = "Being Started"
  Case SERVICE_STOP_PENDING
    sTemp = "Being Stopped"
  Case SERVICE_RUNNING
    sTemp = "Running"
  Case SERVICE_CONTINUE_PENDING
    sTemp = "Being Continued"
  Case SERVICE_PAUSE_PENDING
    sTemp = "Being Paused"
  Case SERVICE_PAUSED
    sTemp = "Paused"
  Case SERVICE_ACCEPT_STOP
    sTemp = "Stopped"
  Case SERVICE_ACCEPT_PAUSE_CONTINUE
    sTemp = "SERVICE_ACCEPT_PAUSE_CONTINUE"
  Case SERVICE_ACCEPT_SHUTDOWN
    sTemp = "Being Shutdown"
  Case Else
    sTemp = "UNKNOWN"
  End Select
  
  
CloseHandles:
   'Close the Handle to the Service
  CloseServiceHandle (hSVC)
  
  ' Close the Handle to the Service Control Manager
  CloseServiceHandle (hSCM)

  ServiceState = sTemp
  
  Open "C:\Nucleus\McInfo.log" For Append As #99
    Print #99, "Service Status complete"
  Close 99
  
  ' DrWatson Log stuff
  d = Dir("c:\winnt\drwtsn32.log")
  If d <> "" Then
    DrWatsonDate = FileDateTime("c:\winnt\drwtsn32.log")
  Else
    DrWatsonDate = CDate("01 Jan 1970") ' Default to something NOT likely
  End If
  
  Open "C:\Nucleus\McInfo.log" For Append As #99
    Print #99, "Dr Watson info complete"
  Close 99
  
  ' Last AV Scan stuff
  d = Dir("c:\Program Files\Network Associates\NetShield NT\Scanlog.txt")
  If d <> "" Then
    FileCopy "c:\Program Files\Network Associates\NetShield NT\Scanlog.txt", "C:\Nucleus\scanlog.txt"
    Open "C:\Nucleus\Scanlog.txt" For Input As #1
      Do Until EOF(1)
        Line Input #1, McAfeeInfo
        P1 = InStr(1, McAfeeInfo, "Files infected")
        If P1 > 0 Then
          VirusesFound = CInt(Right(McAfeeInfo, Len(McAfeeInfo) - P1 - 24))
          P2 = InStr(1, McAfeeInfo, "Scan Summary")
          If P2 > 0 Then
            ScanDate = CDate(Trim(Left(McAfeeInfo, P2 - 1)))
          End If
        End If
      Loop
    Close 1
    Kill "C:\Nucleus\Scanlog.txt" ' Cleanup
  Else
    VirusesFound = 0
  End If
  
  Open "C:\Nucleus\McInfo.log" For Append As #99
    Print #99, "Last AV Scan info complete"
  Close 99
  
  ' Get FileDateTime of Distribution .dat file
  'On Error Resume Next
  d = Dir("e:\Applications\Apps\Virusscan\update\dat-*.zip")
  If d <> "" Then
    Do Until d = ""
      Open "C:\Nucleus\McInfo.log" For Append As #99
        Print #99, "Looking at " & d
      Close 99
      CurrentDate = FileDateTime("e:\Applications\Apps\Virusscan\update\" & d)
      If CurrentDate > DatDateTime Then
        DatDateTime = CDate(CurrentDate)
        DatFilename = d
      End If
      d = Dir()
    Loop
  Else
     DatDateTime = CDate("01 Jan 1970") ' Default to something NOT likely
     DatFilename = "NO DISTRIBUTION FILES FOUND"
  End If
  
  Open "C:\Nucleus\McInfo.log" For Append As #99
    Print #99, "Distribution .dat info complete"
  Close 99
  
  ' Create SQL String ready for execution by SQL Server
  SQLString = ""
  SQLString = "UPDATE tblServer SET"
  SQLString = SQLString & " AVEngineVersion ='" & EngineVer & "'"
  SQLString = SQLString & ", AVProductVersion ='" & ProductVer & "'"
  SQLString = SQLString & ", AVSerialNumber = '" & Serial & "' "
  SQLString = SQLString & ", AVDefsDate = '" & Format(VirusDefDate, "YYYYMMDD") & "'"
  SQLString = SQLString & ", AVDefsVersion = '" & VirusDefVer & "' "
  SQLString = SQLString & ", AVServiceState = '" & ServiceState & "' "
  SQLString = SQLString & ", DrWatsonDate = '" & Format(DrWatsonDate, "YYYYMMDD") & "'"
  SQLString = SQLString & ", AVLastScan = '" & Format(ScanDate, "YYYYMMDD") & "'"
  SQLString = SQLString & ", AVViruses = " & CLng(VirusesFound) & ""
  SQLString = SQLString & ", AVDistVersion = '" & DatFilename & "'"
  SQLString = SQLString & ", AVDistDate = '" & Format(DatDateTime, "YYYYMMDD") & "'"
  SQLString = SQLString & " WHERE Hostname like '" & Hostname & "'"
  
  ' Output SQL String to report
  Open "C:\Nucleus\AVReport.txt" For Output As #1
    Print #1, SQLString
  Close 1
  
  Open "C:\Nucleus\McInfo.log" For Append As #99
    Print #99, "Report complete"
  Close 99
  
  
  ' Get current value of Collecting key
  OldCollection = CInt(GetValue(Hostname, "Software\CSC\Nucleus", "Collecting", False))
  If OldCollection = 0 Then OldCollection = 1
  ' Reset flag so we know this app is finished
  SetValueString Hostname, "Software\CSC\Nucleus", "Collecting", CStr(OldCollection - 1), REG_SZ
  
  ' Clean up
  If Dir("C:\Nucleus\McInfo.log") <> "" Then
    Kill "C:\Nucleus\McInfo.log"
  End If
  
  Unload Me
  End
End Sub

Private Function GetValue(ByVal Hostname As String, ByVal Key As String, ByVal Value As String, ByRef ECode As Boolean) As String
    Dim OpenKeyVal As Long
    Dim OpenHiveVal As Long
    Dim RResult As Long
    Dim InfoTextStr As String
    
    'Init
    ECode = False
    Hostname = Trim(Hostname)
    Key = Trim(Key)
    Value = Trim(Value)
    GetValue = ""
        
    RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then Exit Function
    
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
    InfoTextStr = RegistryQueryValue(OpenKeyVal, Value, REG_SZ)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    GetValue = InfoTextStr
    ECode = True
End Function

Private Function CreateKey(ByVal Hostname As String, ByVal Key As String, NewKey As String) As Boolean
    Dim OpenKeyVal As Long
    Dim OpenHiveVal As Long
    Dim RResult As Long
    Dim lCreateResult As Long
    
    'Init
    Hostname = Trim(Hostname)
    Key = Trim(Key)
        
    RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then
      CreateKey = False
      Exit Function
    End If
    
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
    lCreateResult = RegistryCreateKey(OpenKeyVal, NewKey)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    CreateKey = True
    
End Function

Private Function SetValueString(Hostname As String, Key As String, Value As String, Data As String, regType As Long) As Long
    Dim OpenKeyVal As Long
    Dim RResult As Long
    Dim OpenHiveVal As Long
    Dim strValue, CMPTRName, keytogo As String, InfoTextStr As String
    Dim x As Integer
    
    'GetIPCConnection (Trim(Hostname))
    
    CMPTRName = Trim(Hostname)
    
    RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
    keytogo = Trim(Key)
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
    RegistryWriteValue Data, OpenKeyVal, Value, regType
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    'DisIPCConnection (Trim(Hostname))
        
End Function
