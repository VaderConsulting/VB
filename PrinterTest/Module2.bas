Attribute VB_Name = "Module2"
Option Explicit

Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFF    'Wait forever
Const WAIT_OBJECT_0 = 0    'The state of the specified object is signaled.
Const WAIT_TIMEOUT = &H102 'The time-out interval elapsed, and the object’s state is nonsignaled.

Public Const WM_WININICHANGE = &H1A
Public Const HWND_BROADCAST As Long = &HFFFF&

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type


Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As ACL
    Dacl As ACL
End Type


Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long        ' // Windows 95 only
    dmICMIntent As Long        ' // Windows 95 only
    dmMediaType As Long        ' // Windows 95 only
    dmDitherType As Long       ' // Windows 95 only
    dmReserved1 As Long        ' // Windows 95 only
    dmReserved2 As Long        ' // Windows 95 only
End Type

Type PRINTER_INFO_2
    pServerName As String
    pPrinterName As String
    pShareName As String
    pPortName As String
    pDriverName As String
    pComment As String
    pLocation As String
    pDevMode As DEVMODE
    pSepFile As String
    pPrintProcessor As String
    pDatatype As String
    pParameters As String
    pSecurityDescriptor As SECURITY_DESCRIPTOR
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type


Public Type PRINTER_INFO_5
    pPrinterName As String
    pPortName As String
    Attributes As Long
    DeviceNotSelectedTimeout As Long
    TransmissionRetryTimeout As Long
End Type

Public Type PRINTER_DEFAULTS
    pDatatype As Long
    pDevMode As DEVMODE
    DesiredAccess As Long
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As String) As Long

Public Declare Function AddPrinterConnection Lib "winspool.drv" Alias "AddPrinterConnectionA" (ByVal pName As String) As Long
Public Declare Function DeletePrinterConnection Lib "winspool.drv" Alias "DeletePrinterConnectionA" (ByVal pName As String) As Long

'---------------------------------------------------------------------------------------
' Procedure : CreateGUID
' DateTime  : 12/03/2003 22:52
' Author    : Dave Robinson
' Purpose   : Used to create a GUID for filenames
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Function CreateGUID() As String
    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim RetVal As Long
    
    On Error GoTo CreateGUIDError
    
    lLen = 40
    bGUID = String(lLen, 0)
    
    CoCreateGuid uGUID
    
    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    
    sGUID = bGUID
    If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then
        RetVal = RetVal - 1
    End If
    
    sGUID = Replace(sGUID, "{", "")
    sGUID = Replace(sGUID, "}", "")
    
    CreateGUID = Left$(sGUID, RetVal)
    CreateGUID = Replace(CreateGUID, Chr(0), "")
    Exit Function
CreateGUIDError:
    Err.Clear
End Function

'---------------------------------------------------------------------------------------
' Procedure : ShellWait
' DateTime  : 12/03/2003 22:52
' Author    : Dave Robinson
' Purpose   : Used to execute a process, but forcing VB to wait on it before continuing.
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Sub ShellWait(strPath As String, Optional lTimetoWait As Long = INFINITE)
    Dim lPid As Long
    Dim lHnd As Long
    Dim lRet As Long
    
    If Trim$(strPath) = "" Then Exit Sub
    
    lPid = Shell(strPath, vbNormalFocus)
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)       'Get a handle to the shelled process.
        If lHnd <> 0 Then                              'If successful, wait for the
            lRet = WaitForSingleObject(lHnd, lTimetoWait) ' application to end.
            CloseHandle (lHnd)                         'Close the handle.
        End If
    End If
End Sub

