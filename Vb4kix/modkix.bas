Attribute VB_Name = "modKix"
Option Explicit

Public Errorcode As Long
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Const REG_NONE                     As String = "REG_NONE"
Public Const REG_SZ                       As String = "REG_NONE"
Public Const REG_EXPAND_SZ                As String = "REG_EXPAND_SZ"
Public Const REG_BINARY                   As String = "REG_BINARY"
Public Const REG_DWORD                    As String = "REG_DWORD"
Public Const REG_DWORD_LITTLE_ENDIAN      As String = "REG_DWORD_LITTLE_ENDIAN"
Public Const REG_DWORD_BIG_ENDIAN         As String = "REG_DWORD_BIG_ENDIAN"
Public Const REG_LINK                     As String = "REG_LINK"
Public Const REG_MULTI_SZ                 As String = "REG_MULTI_SZ"
Public Const REG_RESOURCE_LIST            As String = "REG_RESOURCE_LIST"
Public Const REG_FULL_RESOURCE_DESCRIPTOR As String = "REG_FULL_RESOURCE_DESCRIPTOR"

Public Function CD(strDirname As String) As Long
    On Error GoTo CatchError
    ChDir strDirname
    Exit Function
CatchError:
    CD = Err.Number
    Err.Clear
End Function

Public Function Copy(strSource As String, strDestination As String) As Long
    On Error GoTo CatchError
    FileCopy strSource, strDestination
    Exit Function
CatchError:
    Copy = Err.Number
    Err.Clear
End Function

Public Function Del(strFilename As String) As Long
    On Error GoTo CatchError
    Kill strFilename
    Exit Function
CatchError:
    Del = Err.Number
    Err.Clear
End Function

Public Function Go(strDrive As String) As Long
    Go = CD(strDrive)
End Function

Public Function MD(strDirname As String) As Long
    On Error GoTo CatchError
    MkDir strDirname
    Exit Function
CatchError:
    MD = Err.Number
    Err.Clear
End Function

Public Function Quit() As Long
    On Error GoTo CatchError
    End
    Exit Function
CatchError:
    Quit = Err.Number
    Err.Clear
End Function

Public Function RD(strDirname As String) As Long
    On Error GoTo CatchError
    RmDir strDirname
    Exit Function
CatchError:
    RD = Err.Number
    Err.Clear
End Function

Public Function Run(strAppname As String) As Long
    On Error GoTo CatchError
    Shell strAppname, vbNormalNoFocus
    Exit Function
CatchError:
    Run = Err.Number
    Err.Clear
End Function

Public Function SetL(strVariable As String) As Long
    On Error GoTo CatchError
    ' Set environment variable in the local environment of this app
    Exit Function
CatchError:
    SetL = Err.Number
    Err.Clear
End Function

Public Function SetM(strFilename As String) As Long
    On Error GoTo CatchError
    ' Set environment variable in the local environment of this computer
    Exit Function
CatchError:
    SetM = Err.Number
    Err.Clear
End Function

Public Function SetTime(strSource As String) As Long
    Dim wmi As Object, process As Object, startupInfo As Object, cmd As String
    On Error GoTo CatchError
    ' Set the local time according to the source.
    ' Source may be:
    '  1. A server name
        Set wmi = GetObject("winmgmts:{(SystemTime)}!//" & strSource)
        Set process = wmi.get("win32_process")
        Set startupInfo = wmi.get("win32_processstartup")
        startupInfo.showWindow = 0 '0=hidden 1=normal 7=minnoactivate
    
        cmd = "net time \\" & timeserver & " /set /y"
        process.Create cmd, , startupInfo
    '  2. A Domainname
    '  3. * - Browse the local domain for any time source
    Exit Function
CatchError:
    SetTime = Err.Number
    Err.Clear
End Function

Public Function kixShell(strAppname As String) As Long
    On Error GoTo CatchError
    ' TODO: Modify this to wait for the shelled application
    Shell strAppname, vbNormalFocus
    Exit Function
CatchError:
    kixShell = Err.LastDllError
    Err.Clear
End Function

Public Function Sleep(dblSeconds As Double) As Long
    Dim d As Date
    On Error GoTo CatchError
    d = Date
    Do Until Now >= DateAdd("s", dblSeconds, d)
        DoEvents
    Loop
    Exit Function
CatchError:
    Sleep = Err.Number
    Err.Clear
End Function

Public Function Use(Optional strOptions As String) As Long
    On Error GoTo CatchError
    Select Case strOptions
        Case "LIST"
        'TODO: Display drive connections
        
    End Select
    Exit Function
CatchError:
    Del = Err.Number
    Err.Clear
End Function

Public Function AddKey(strSubkey As String) As Long
    Dim strComputer As String, oReg As Object
    On Error GoTo CatchError
    strComputer = "."
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    oReg.CreateKey HKEY_LOCAL_MACHINE, strSubkey
    Exit Function
CatchError:
    AddKey = Err.Number
    Err.Clear
End Function

Public Function AddPrinterConnection(strPrintername As String, Optional bSetDefault As Boolean = False) As Long
    Dim wshNetwork As Object
    On Error GoTo CatchError
    Set wshNetwork = CreateObject("WScript.Network")
    wshNetwork.AddWindowsPrinterConnection strPrintername
    If bSetDefault Then
        wshNetwork.SetDefaultPrinter strPrintername
    End If
    Exit Function
CatchError:
    AddPrinterConnection = Err.Number
    Err.Clear
End Function

Public Function AddProgramGroup(strGroupname As String, Optional bCommon As Boolean = False) As Long
    On Error GoTo CatchError
    'TODO: Add a program group
    If bCommon Then
        
    Else
        
    End If
    Exit Function
CatchError:
    AddProgramGroup = Err.Number
    Err.Clear
End Function

Public Function AddProgramItem(strCommandLine As String, strName As String, strIconPath As String, lngIconIndex As Long, strDefaultDirectory As String, bMinimise As Boolean, bReplace As Boolean, bRunInOwnSpace As Boolean) As Long
    On Error GoTo CatchError
    'TODO: Add a program item
    If bCommon Then
        
    Else
        
    End If
    Exit Function
CatchError:
    AddProgramGroup = Err.Number
    Err.Clear
End Function

Public Function BackupEventLog(strEventLog As String, strBackupFile As String) As Long
    Dim objWMIService As Object, colRetrievedEvents As New Collection
    Dim strOutput As String
    On Error GoTo CatchError
    
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colRetrievedEvents = objWMIService.ExecQuery("Select * from Win32_NTLogEvent")
    
    Open strBackupFile For Output As #1
    
        For Each objEvent In colRetrievedEvents
            strOutput = ""
            strOutput = strOutput & objEvent.Category
            strOutput = strOutput & "," & objEvent.ComputerName
            strOutput = strOutput & "," & objEvent.EventCode
            strOutput = strOutput & "," & objEvent.Message
            strOutput = strOutput & "," & objEvent.RecordNumber
            strOutput = strOutput & "," & objEvent.SourceName
            strOutput = strOutput & "," & objEvent.TimeWritten
            strOutput = strOutput & "," & objEvent.Type
            strOutput = strOutput & "," & objEvent.User
            Print #1, strOutput
        Next
    
    Close 1

    Exit Function
CatchError:
    AddProgramGroup = Err.Number
    Err.Clear
End Function

Public Function WriteValue(strSubkey As String, strEntry As String, strExpression, strDatatype As String) As Long
    Dim strComputer As String, oReg As Object
    On Error GoTo CatchError
    strComputer = "."
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    
    Select Case strDatatype
        Case REG_SZ
            oReg.SetStringValue HKEY_LOCAL_MACHINE, strSubkey, strEntry, strExpression
        Case REG_DWORD
            oReg.SetDWORDValue HKEY_LOCAL_MACHINE, strSubkey, strEntry, strExpression
    End Select
    Exit Function
CatchError:
    AddKey = Err.Number
    Err.Clear
End Function
