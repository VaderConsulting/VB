VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{B186D399-9B91-41CB-9242-E12B96CA99E1}#1.0#0"; "Ping.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "My Server Monitor"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3240
   StartUpPosition =   1  'CenterOwner
   Begin Ping.drPing pngServer 
      Left            =   120
      Top             =   3480
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   10000
      Left            =   3000
      Top             =   2520
   End
   Begin MSChart20Lib.MSChart chtDisk 
      CausesValidation=   0   'False
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "frmMain.frx":0442
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      Caption         =   "Diskspace"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
  
  ' ------------------------------ DISK SPACE ------------------------
  Private Declare Function NetServerEnum Lib "netapi32.dll" (vServerName As Any, ByVal lLevel As Long, vBufptr As Any, lPrefmaxlen As Long, lEntriesRead As Long, lTotalEntries As Long, vServerType As Any, ByVal sDomain As String, vResumeHandle As Any) As Long
  Private Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, vSrc As Any, ByVal lSize&)
  Private Declare Sub lstrcpyW Lib "kernel32" (vDest As Any, ByVal sSrc As Any)
  Private Declare Sub lstrcpy Lib "kernel32" (vDest As Any, ByVal vSrc As Any)
  Private Declare Sub lstrcpynW Lib "kernel32" (ByVal vDest As Any, ByVal vSrc As Any, lLength As Long)
  Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirName As String, lpBytesToCall As Currency, lpTotalBytes As Currency, lpFreeBytes As Currency) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
  Private Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal lpBuffer As Long) As Long
  Private Declare Function NetWkstaGetInfo Lib "netapi32.dll" (ByVal sServerName$, ByVal lLevel&, vBuffer As Any) As Long
  
  Private Type SERVER_INFO_100
    sv100_platform_id As Long
    sv100_servername As Long
  End Type
  
  Private Type SERVER_INFO_101
    dw_platform_id As Long
    ptr_name As Long
    dw_ver_major As Long
    dw_ver_minor As Long
    dw_type As Long
    ptr_comment As Long
  End Type
  
  Private Type WKSTA_INFO_100
    wki100_platform_id As Long
    wki100_computername As Long
    wki100_langroup As Long
    wki100_ver_major As Long
    wki100_ver_minor As Long
  End Type
      
  Private Const ERROR_EXTENDED_ERROR = 1208&
  Private Const ERROR_NO_MORE_ITEMS = 259&
  Private Const ERROR_MORE_DATA = 234
  
  Private Const ERROR_NONE = 0&
  Private Const NOERROR = 0
  Private Const NO_ERROR = 0
  Private Const ERROR_SUCCESS = 0&
  Private Const NERR_Success As Long = 0&
  
  ' More of the SV_TYPE constants are here so that
  ' the ListServers call can be used to enumate
  ' different types of servers in other functions
  
  Private Const SV_TYPE_WORKSTATION = &H1
  Private Const SV_TYPE_SERVER = &H2
  Private Const SV_TYPE_SQLSERVER = &H4
  Private Const SV_TYPE_DOMAIN_CTRL = &H8
  Private Const SV_TYPE_DOMAIN_BAKCTRL = &H10
  Private Const SV_TYPE_TIMESOURCE = &H20
  Private Const SV_TYPE_AFP = &H40
  Private Const SV_TYPE_NOVELL = &H80
  Private Const SV_TYPE_DOMAIN_MEMBER = &H100
  Private Const SV_TYPE_LOCAL_LIST_ONLY = &H40000000
  Private Const SV_TYPE_PRINT = &H200
  Private Const SV_TYPE_DIALIN = &H400
  Private Const SV_TYPE_XENIX_SERVER = &H800
  Private Const SV_TYPE_MFPN = &H4000
  Private Const SV_TYPE_NT = &H1000
  Private Const SV_TYPE_WFW = &H2000
  Private Const SV_TYPE_SERVER_NT = &H8000
  Private Const SV_TYPE_POTENTIAL_BROWSER = &H10000
  Private Const SV_TYPE_BACKUP_BROWSER = &H20000
  Private Const SV_TYPE_MASTER_BROWSER = &H40000
  Private Const SV_TYPE_DOMAIN_MASTER = &H80000
  Private Const SV_TYPE_DOMAIN_ENUM = &H80000000
  Private Const SV_TYPE_WINDOWS = &H400000
  Private Const SV_TYPE_ALL = &HFFFFFFFF
  
  ' Define GB and MB values
  Const lGByte As Long = 1073741824
  Const lMByte As Long = 1048576
  Dim freePercent As String
  Dim usedPercent As String
  Dim freeSpace As String
  Dim usedSpace As String
  Dim totalSpace As String
  ' -----------------------PERFORMANCE----------------------
  
'Private Enum PERF_DETAIL
'  PERF_DETAIL_NOVICE = 100      ' The uninformed can understand it
'  PERF_DETAIL_ADVANCED = 200    ' For the advanced user
'  PERF_DETAIL_EXPERT = 300      ' For the expert user
'  PERF_DETAIL_WIZARD = 400      ' For the system designer
'End Enum
'
'Private Enum PDH_STATUS
'  PDH_CSTATUS_VALID_DATA = &H0
'  PDH_CSTATUS_NEW_DATA = &H1
'  PDH_CSTATUS_NO_MACHINE = &H800007D0
'  PDH_CSTATUS_NO_INSTANCE = &H800007D1
'  PDH_MORE_DATA = &H800007D2
'  PDH_CSTATUS_ITEM_NOT_VALIDATED = &H800007D3
'  PDH_RETRY = &H800007D4
'  PDH_NO_DATA = &H800007D5
'  PDH_CALC_NEGATIVE_DENOMINATOR = &H800007D6
'  PDH_CALC_NEGATIVE_TIMEBASE = &H800007D7
'  PDH_CALC_NEGATIVE_VALUE = &H800007D8
'  PDH_DIALOG_CANCELLED = &H800007D9
'  PDH_CSTATUS_NO_OBJECT = &HC0000BB8
'  PDH_CSTATUS_NO_COUNTER = &HC0000BB9
'  PDH_CSTATUS_INVALID_DATA = &HC0000BBA
'  PDH_MEMORY_ALLOCATION_FAILURE = &HC0000BBB
'  PDH_INVALID_HANDLE = &HC0000BBC
'  PDH_INVALID_ARGUMENT = &HC0000BBD
'  PDH_FUNCTION_NOT_FOUND = &HC0000BBE
'  PDH_CSTATUS_NO_COUNTERNAME = &HC0000BBF
'  PDH_CSTATUS_BAD_COUNTERNAME = &HC0000BC0
'  PDH_INVALID_BUFFER = &HC0000BC1
'  PDH_INSUFFICIENT_BUFFER = &HC0000BC2
'  PDH_CANNOT_CONNECT_MACHINE = &HC0000BC3
'  PDH_INVALID_PATH = &HC0000BC4
'  PDH_INVALID_INSTANCE = &HC0000BC5
'  PDH_INVALID_DATA = &HC0000BC6
'  PDH_NO_DIALOG_DATA = &HC0000BC7
'  PDH_CANNOT_READ_NAME_STRINGS = &HC0000BC8
'End Enum
  
'Private Declare Function PdhVbGetOneCounterPath Lib "PDH.DLL" (ByVal PathString As String, ByVal PathLength As Long, ByVal DetailLevel As Long, ByVal CaptionString As String) As Long
'Private Declare Function PdhVbCreateCounterPathList Lib "PDH.DLL" (ByVal PERF_DETAIL As Long, ByVal CaptionString As String) As Long
'Private Declare Function PdhVbGetCounterPathFromList Lib "PDH.DLL" (ByVal Index As Long, ByVal Buffer As String, ByVal BufferLength As Long) As Long
'Private Declare Function PdhOpenQuery Lib "PDH.DLL" (ByVal Reserved As Long, ByVal dwUserData As Long, ByRef hQuery As Long) As PDH_STATUS
'Private Declare Function PdhCloseQuery Lib "PDH.DLL" (ByVal hQuery As Long) As PDH_STATUS
'Private Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As PDH_STATUS
'Private Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As PDH_STATUS
'Private Declare Function PdhVbIsGoodStatus Lib "PDH.DLL" (ByVal StatusValue As Long) As Long
'Private Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double

'Dim hQuery As Long
  ' --------------------------------------------------------

Private Sub tmrUpdate_Timer()
  Dim ServerName As String
  Dim ServerSpace As String, ServerCPU As String
  Dim PingTime As String
  On Error Resume Next
  tmrUpdate.Enabled = False
  ' Get % space used on a specific driveletter:
  lblStatus = "Retrieving disk space..."
  lblStatus.Refresh
  ServerName = "AZISPER1S01"
  'ServerName = "Server"
  PingTime = pngServer.Ping(ServerName)  ' is the server up?
  If PingTime <> "1048596" Then          ' YES!
    ServerSpace = GetSpace(ServerName, "\c$")
    'ServerSpace = "91"
    chtDisk.Row = 1
    chtDisk.Column = 1
    'chtDisk.RowLabel = ServerName & "\c$"
    If ServerSpace = "" Then
      lblStatus = "Error..."
      lblStatus.Refresh
    Else
      lblMain = "Diskspace " & CInt(ServerSpace) & "% used"
      chtDisk.Data = CInt(ServerSpace)
      If CInt(ServerSpace) > 90 Then       ' >90% used!!!!!
        MsgBox "Diskspace low on " & ServerName & " (" & 100 - CInt(ServerSpace) & "% free)", vbExclamation, "My Server Monitor Error"
      End If
      chtDisk.Visible = True
      lblStatus = "Idle"
    End If
  Else                                  ' NO!
    lblStatus = "Error contacting " & ServerName
    chtDisk.Visible = False
  End If
  ' Get CPU Counter:
  'lblStatus = "Retrieving performance data..."
  'lblStatus.Refresh
  'ServerCPU = PerformanceData(ServerName, "Processor(0)\% Processor Time")
  'PdhCloseQuery (hQuery)  ' Free the query
  'If ServerCPU = "<ERROR>" Then
  '  lblStatus = "Error collecting Performance Data"
  '  lblStatus.Refresh
  'Else
  '  chtCPU.Row = 1
  '  chtCPU.Column = 1
  '  chtCPU.Data = CInt(ServerCPU)
  'End If
  tmrUpdate.Enabled = True
  lblStatus.Refresh
End Sub


Public Function GetSpace(Server As String, DriveLetter As String) As String
  Dim DiskUsed As String, CurrentDisk As String
  Dim up As String, us As String, fp As String, fs As String, ts As String
  Dim msg As String
  DoEvents
  DiskUsed = FindDriveInfo(Server, CStr(DriveLetter))
  DoEvents
  
  CurrentDisk = Left(DriveLetter, 1)
  up = usedPercent
  us = usedSpace
  fp = freePercent
  fs = freeSpace
  ts = totalSpace
  GetSpace = up
End Function

Public Function FindDriveInfo(Server As String, sDrive As String) As String
    Dim lReturn As Long
    Dim cBytesToCall As Currency
    Dim cBytesOnDrive As Currency
    Dim cFreeBytes As Currency
    Dim cUsedBytes As Currency
    Dim sServer As String
    Dim Output1 As String
    Dim Output2 As String
    Dim Output3 As String
    Dim TotalBytes As String
    Dim UsedBytes As String

    sServer = "\\" & Server & "\" & sDrive
     
    ' Use in place of GetDiskFreeSpaceEx to bypass the
    ' 2 GB limit.  Use the Currency Data Type multiplied by
    ' 10000 so that 9,223,372,036,854,775,807 is the limit
    ' of the return value (see "Hardcore Visual Basic", Chap. 2
    ' for more information
            
    lReturn = GetDiskFreeSpaceEx(sServer, cBytesToCall, cBytesOnDrive, cFreeBytes)
    
    DoEvents
    
    If lReturn = 0 Then
        Exit Function
    End If

    If lReturn <> 0 Then
    ' Format the output based on the returned value so that the string output is in MBs or GBs.

        ' Total Free Bytes on Drive
        Output1 = Format((cBytesOnDrive * 10000) / IIf(cBytesOnDrive * 10000 >= lGByte, lGByte, lMByte), IIf(cBytesOnDrive * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
                  
        ' Total Available Bytes on Drive
        Output2 = Format((cFreeBytes * 10000) / IIf(cFreeBytes * 10000 >= lGByte, lGByte, lMByte), IIf(cFreeBytes * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
                  
        ' Calculate Used Bytes
        cUsedBytes = cBytesOnDrive - cFreeBytes
        
        ' Total Used Bytes on Drive
        Output3 = Format((cUsedBytes * 10000) / IIf(cUsedBytes * 10000 >= lGByte, lGByte, lMByte), IIf(cUsedBytes * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
        
        TotalBytes = Format((cBytesOnDrive * 10000) / IIf(cBytesOnDrive * 10000 >= lGByte, lGByte, lMByte), IIf(cBytesOnDrive * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
        
        UsedBytes = Format((cUsedBytes * 10000) / IIf(cUsedBytes * 10000 >= lGByte, lGByte, lMByte), IIf(cUsedBytes * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
        
    End If
    
    ' If Output1 (Total Free Bytes on Drive) is 0, then
    ' you know the drive does not exist.  Otherwise, send the output to FindDriveInfo
    
    If Output1 <> "0.MB" Then
        freePercent = Int((cFreeBytes / cBytesOnDrive) * 100)
        usedPercent = Int((cUsedBytes / cBytesOnDrive) * 100)
        usedSpace = UsedBytes
        freeSpace = Output2
        totalSpace = TotalBytes
        FindDriveInfo = usedPercent
    Else
        ' Do nothing in particular
    End If
End Function
    
'Public Function PerformanceData(Hostname As String, Counter As String) As String
'  Dim pdhStatus As PDH_STATUS, Status As PDH_STATUS
'  Dim hCounter As Long
'  Dim dblCounterValue As Double
'  Dim Res As Long
'  Dim strInfo As String, strData As String, PDHPath As String
'  Dim i As Long, P As Integer
'  ' Ensure the counter name starts with "\"
'  If Left(Counter, 1) <> "\" Then Counter = "\" & Counter
'  PDHPath = "\\" & Hostname & Counter
'  ' Create the query, and verify that the request succeeded.
'  pdhStatus = PdhOpenQuery(0, 1, hQuery)
'  If pdhStatus <> ERROR_SUCCESS Then
'    'Set pdhStatus = Nothing
'    PerformanceData = "<ERROR>"
'    Exit Function
'  End If
'  ' Associate the counter path to the query
'  Status = PdhVbAddCounter(hQuery, PDHPath, hCounter)
'  PdhCollectQueryData (hQuery)
'  dblCounterValue = PdhVbGetDoubleCounterValue(hCounter, Res)
'  ' Verify that when we queried, the returned value was valid
'  If (Res = PDH_CSTATUS_VALID_DATA) Or (Res = PDH_CSTATUS_NEW_DATA) Then
'    strInfo = PDHPath + " = " + Format$(dblCounterValue, "0") + Chr$(13) + Chr$(10)
'    strData = Format$(dblCounterValue, "0")
'    P = InStr(3, PDHPath, "\")
'  End If
'  PerformanceData = strData
'End Function
'
