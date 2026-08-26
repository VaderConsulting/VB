VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon in progress"
   ClientHeight    =   1470
   ClientLeft      =   1110
   ClientTop       =   1470
   ClientWidth     =   4800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1470
   ScaleWidth      =   4800
   StartUpPosition =   1  'CenterOwner
   WindowState     =   1  'Minimized
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Closing this application WILL CANCEL your logon."
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox lstStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.ListBox lstGroups 
      Height          =   255
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   "The Logon Script is processing...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Enforce variable declaration
Option Explicit

'Dim csvGroupname(200) As String
'Dim csvPDrive(200) As String
'Dim csvTDrive(200) As String
'Dim csvNDrive(200) As String
'Dim csvXDrive(200) As String
'Dim csvODrive(200) As String
'Dim csvQDrive(200) As String
'Dim csvPrinter1(200) As String
'Dim csvPrinter2(200) As String
'Dim csvPrinter3(200) As String
'Dim DriveToMap(6) As String
'Dim PrinterToMap(25) As String
'Dim travellerPrinter(25) As String
'Dim DriveLetter(6) As String
'Dim Ticks(6) As Long

Dim Username As String
Dim LogonServer As String, siteServer As String
Dim StartTime As String
Dim StartDate As String
Dim StartTick As Long
Dim FinishTime As String
Dim FinishTick As Long
Dim siteLoc As String
Dim IPAddress As String
Dim ServicePack As String
Dim Computername As String
Dim ErrorNumber As Long
Dim CriticalError As Boolean, CriticalOnly As Boolean
Dim disconnectErrorflag As Integer, connectErrorflag As Integer, printerErrorflag As Integer
Dim prtlocError As Integer
Dim errorflag As Integer
Dim DebugMode As Boolean
Dim ErrorNum As Long
Dim ErrorMsg As String
Dim Version As Integer
Dim Traveller As Boolean
Dim Client As String
Dim APIMode As Boolean    ' Sets whether to use API for drive mappings or not
Dim API As Integer        ' 1 = old method (WNetAddConnection)
                          ' 2 = new method (WNetAddConnection2)

Dim InternalTesting As Boolean
Dim isLaptop As Boolean

Dim pNo As Integer
Dim rc As Long

Const WN_Success = &H0
Const WN_Not_Supported = &H1
Const WN_Net_Error = &H2
Const WN_Bad_Pointer = &H4
Const WN_Bad_NetName = &H32
Const WN_Bad_Password = &H6
Const WN_Bad_Localname = &H33
Const WN_Access_Denied = &H7
Const WN_Out_Of_Memory = &HB
Const WN_Already_Connected = &H34
Const WN_Server_Down = &H35
Const WN_Server_Down_2 = &H52E
Const WN_NOT_FOUND = &H43
Const WN_IN_USE = 85
Const WN_NOT_CONNECTED = 2250

' Windows Message constants
Const SMTO_ABORTIFHUNG = &H2
Const HWND_BROADCAST = &HFFFF
Const WM_SETTINGCHANGE = &H1A
Const WM_WININICHANGE = &H1A  ' <--- OBSOLETE, USE WM_SETTINGCHANGE
Const WM_USERCHANGED = &H54

' Constants for enable/disable desktop
Const GW_CHILD = 5
Const SW_HIDE = 0
Const SW_SHOW = 5

'Consts for return codes errors
Const ERROR_NO_ERROR = 0&
Const ERROR_ALREADY_ASSIGNED = 85&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_BAD_DEVICE_TYPE = 66&
Const ERROR_BAD_NET_NAME = 67&
Const ERROR_BAD_PROFILE = 1206&
Const ERROR_BAD_DEVICE = 1200&
Const ERROR_BAD_PROVIDER = 1204&
Const ERROR_BUSY = 170&
Const ERROR_CANCEL_VIOLATION = 173&
Const ERROR_CANNOT_OPEN_PROFILE = 1205&
Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
Const ERROR_EXTENDED_ERROR = 1208&
Const ERROR_INVALID_PASSWORD = 86&
Const ERROR_NO_NET_OR_BAD_PATH = 1203&
Const ERROR_SESSION_CREDENTIAL_CONFLICT = 1219&
Const ERROR_NO_NETWORK = 1222&
Const ERROR_CANCELLED = 1223&
Const ERROR_NO_CONNECTION = 8
Const ERROR_NO_DISCONNECT = 9
Const ERROR_DEVICE_IN_USE = 2404&
Const ERROR_NOT_CONNECTED = 2250&
Const ERROR_OPEN_FILES = 2401&
Const ERROR_MORE_DATA = 234

Const CONNECT_UPDATE_PROFILE = &H1
Const RESOURCETYPE_DISK = &H1
Const RESOURCETYPE_PRINT = &H2
Const RESOURCETYPE_ANY = &H0
Const RESOURCE_CONNECTED = &H1
Const RESOURCE_REMEMBERED = &H3
Const RESOURCE_GLOBALNET = &H2
Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Const RESOURCEDISPLAYTYPE_SERVER = &H2
Const RESOURCEDISPLAYTYPE_SHARE = &H3
Const RESOURCEUSAGE_CONNECTABLE = &H1
Const RESOURCEUSAGE_CONTAINER = &H2

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

Private Type NETCONNECTINFOSTRUCT
    cbStructure As Long
    dwFlags As Long
    dwSpeed As Long
    dwDelay As Long
    dwOptdataSize As Long
End Type

Dim lpNetResource As NETRESOURCE

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

' Declare API functions
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal Hostname As String, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal Hostname As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function AddPrinterConnection Lib "winspool.drv" Alias "AddPrinterConnectionA" (ByVal pName As String) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Integer
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Public Sub Form_Load()
    ' Set up variables with default variables, or dimension as appropriate
    Dim a As Integer
    Dim PRTLOC As String, retval As Long, strTemp1 As String, strTemp2 As String
    Dim ServerFlag As Boolean, lp As Integer, lp2 As Integer, msg As String
    Dim d As String, Source As String, Destination As String
    Dim erDesc As String, i As Integer
    Dim erNum As Long, cmd As String
    Dim GroupTempVar As String
    Dim varFilename As String, header As String, stuff As String, var As Integer
    Dim dr As String, LocalFile As String
    Dim fileinfo1 As Date, fileinfo2 As Date, varDiff As Long
    Dim drv As Integer, prn As Integer, drvLetter As String
    Dim Group As String, junk As String, HighEnd As Boolean
    Dim info As String, varInfo As Integer
    Dim dummy As Integer, varResult As Long
    Dim varPos1 As Integer, varPos2 As Integer
    Dim try As Integer, driveError As Boolean
    Dim groupNumber As Integer, printerNumber As Integer, E As String
    Dim travellerPDrive As String, travellerTDrive As String, travellerNDrive As String
    Dim travellerXDrive As String, travellerODrive As String, travellerQDrive As String
    Dim DrivePath(6) As String
    Dim printer1 As String, printer2 As String, printer3 As String, csvPrinter As String
    Dim t As String, pos As Integer, Servername As String, ServerIP As String
    Dim fso As Object, fsoDrive As Object, KIXpath As String
    Dim ModeChange As Boolean, fnum As Integer
    Dim Subnet As String, bResult As Boolean
    Dim f As Long, strTemp As String, lngTemp As Long
    Dim bReturnVal As Boolean, lngReturnval As Long
    Dim lParam As String, strWinDir As String
    
    ' Disable desktop
    EnableDesktop False
    If Dir("c:\temp\mapidone.txt") <> "" Then
        Kill "c:\temp\mapidone.txt"
    End If
    
    
    InternalTesting = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoDrive = fso.GetDrive(fso.GetDriveName("c:"))
        
    SocketsInitialize
        
    'Default to always be a traveller
    Traveller = True
    
    KIXpath = "c:\winnt\kix32.exe"
    ' Set up defaults - these can be overridden at runtime with commandline options
    DebugMode = False
    Version = 32
    HighEnd = False
    ' **********************************
    ' Logon33 Modification... ModeChange was False, APIMode was False
    ModeChange = True     ' Allows Logon to step between version 3.1 and 3.2 mode to correct errors
    APIMode = True        ' Determines whether to use 'NET USE' or API for drive connections
    ' End Logon33 Modification
    ' **********************************
    API = 2               ' 1 = old method (WNetAddConnection +  WNetCancelConnection)
                          ' 2 = new method (WNetAddconnection2 + WNetCancelConnection2)
                    
    ' The next value determines whether to display all errors, or Critical Errors only
    CriticalOnly = True
    
    ' *******************************************************************************************
    '                                        Some useful info
    '
    '  A.  If you can't read VB, then give up now.  Seriously, this is not the kind of program to
    '      learn with.  Good luck if you do persist.
    '  B.  If you can't code in VB, then these comments should help.
    '  C.  Naming conventions:
    '      Hungarian notation is used with the following prefixes
    '      frm = Form
    '      lst = listbox
    '      cmd = commandbutton
    '  D.  Even though VB is event oriented, this code is essentially initiated by the one event,
    '      this event being the Form_Load event.
    '  E.  Version 31 (3.1) is the older mode, which depends upon KIX32.exe being in c:\winnt
    '      Version 32 (3.2) has no external dependencies, except for a groups.txt on the
    '      logonserver.  This file is a list of NT Global Groups the user is a member of in the
    '      form:
    '              Group1
    '              Group2
    '              GroupN
    '      This file does not have to be alphabetically sorted, as this exe places each entry in
    '      a sorted listbox before using it.  (A quick and dirty sort routine)
    '  F.  DebugMode is used to determine whether we should output debugging info to the log file
    '  G.  Read the DebugMsg text as this will give you a good idea of what is happening
    '  H.  Ticks(x) where x = 1 to 5 is the same as Times(x), but expressed in milliseconds since
    '      Windows was started.  This is a replacement for the original Times(x).
    ' *******************************************************************************************
    
    
    ' Start of code
    ' Extract environment variables, and initialisation
    StartTick = GetTickCount
    StartTime = Time
    StartDate = Format(Year(Date), "0000") & "/" & Format(Month(Date), "00") & "/" & Format(Day(Date), "00")
    Me.Show
    frmMain.Caption = "Logon processing....."
        
    ' Commandline options
    If InStr(1, UCase(Trim(Command$)), "DEBUG") <> 0 Then
        DebugMode = True
    End If
    If InStr(1, UCase(Trim(Command$)), "V31") <> 0 Then
        Version = 31
    End If
    If InStr(1, UCase(Trim(Command$)), "V32") <> 0 Then
        Version = 32
    End If
    If InStr(1, UCase(Trim(Command$)), "M0") <> 0 Then
        ModeChange = False
    End If
    If InStr(1, UCase(Trim(Command$)), "M1") <> 0 Then
        ModeChange = True
    End If
    If InStr(1, UCase(Trim(Command$)), "A0") <> 0 Then
        APIMode = False
    End If
    If InStr(1, UCase(Trim(Command$)), "A1") <> 0 Then
        APIMode = True
        API = 1
    End If
    If InStr(1, UCase(Trim(Command$)), "A2") <> 0 Then
        APIMode = True
        API = 2
    End If
    If InStr(1, UCase(Trim(Command$)), "C0") <> 0 Then
        CriticalOnly = False
    End If
    If InStr(1, UCase(Trim(Command$)), "C1") <> 0 Then
        CriticalOnly = True
    End If
    If InStr(1, UCase(Trim(Command$)), "IT") <> 0 Then
        InternalTesting = True
    End If
    
    ' **********************************
    ' Logon33 modification:  debugmsg lines below were here originally
    ' **********************************
    Computername = UCase(Environ$("computername"))
    Username = UCase(Environ$("username"))
    LogonServer = UCase(Environ$("Logonserver"))
    
    ' Logon33 mod
    If InternalTesting Then LogonServer = "\\CBDXAAA"
    'siteServer = UCase(Environ$("Siteserver")) ' <---- Pre Logon 33
    'siteLoc = UCase(Environ$("siteloc"))       ' <---- Pre Logon 33
    PRTLOC = UCase(Environ$("prtloc"))
    Client = UCase(Environ$("client"))
    
    ' **********************************
    ' Logon33 modification:  these lines were located before computername= line
    DebugMsg "Commandline options: " & Command$
    DebugMsg "                     DEBUG - Debug Mode"
    DebugMsg "                     V31   - Version 3.1"
    DebugMsg "                     V32   - Version 3.2"
    DebugMsg "                     M0    - Modechange off"
    DebugMsg "                     M1    - Modechange on"
    DebugMsg "                     A0    - API Mode off"
    DebugMsg "                     A1    - API Mode 1"
    DebugMsg "                     A2    - API Mode 2"
    DebugMsg "                     C0    - CriticalOnly off"
    DebugMsg "                     C1    - CriticalOnly on"
    DebugMsg "                     IT    - Internal Testing"

    UpdateStatus "Logon Application v" & App.Major & "." & App.Minor & "." & App.Revision
    UpdateStatus "Mode: " & Version
    UpdateStatus "Logon started " & Format(Date, "dd/mm/yyyy")
    ' End Logon33 modification
    ' **********************************
    IPAddress = GetIPAddress
    
    ' ****************************************
    ' Determine Subnet (Logon 33 modification)
    i = InStrRev(IPAddress, ".")
    If i <> 0 Then
        Subnet = Left(IPAddress, i - 1)
        'bResult = GetSiteVariables(Subnet, siteLoc, siteServer)  ' <-- Access
        bResult = GetSiteVariables2(Subnet, siteLoc, siteServer) ' <-- Looping thru file
        If bResult = False Then ' GetSiteVariables failed!!!.. revert to pre-Logon 33 method
            siteServer = UCase(Environ$("Siteserver"))
            siteLoc = UCase(Environ$("siteloc"))
            UpdateStatus "Extracted environment variables"
        Else
            UpdateStatus "Extracted siteloc and siteserver from siteinfo.csv"
        End If
    Else ' IPAddress not determined!!!.. revert to pre-Logon 33 method
        siteServer = UCase(Environ$("Siteserver"))
        siteLoc = UCase(Environ$("siteloc"))
        UpdateStatus "Extracted environment variables"
    End If
    ' ****************************************
    ' Logon 33 modification
    ' 3-8-01 D. Robinson
    ' This code was moved from the EndProgram subroutine
    
    isLaptop = False
    
    ' Ensure neither error 256 or error 512
    '    Error 512 is logon cancelled (user clicked on 'Close')
    '    Error 256 is logon cancelled (server)
    If (ErrorNumber And 512) = 0 And (ErrorNumber And 256) = 0 Then
        '
        ' Laptop determination
        ' A desktop PC should not be able to write to D:, as it should be a CD ROM Drive
        On Error Resume Next
            f = FreeFile
            strTemp = "D:\Attempt to create flag file by Logon Application.txt"
            Err.Clear
            Open strTemp For Output As #f
                Print #f, "This file is safe to delete"
            Close f
            If Err.Number = 0 Then
                isLaptop = True
                UpdateStatus "Laptop detected"
                ' Remove temporary file
                Kill strTemp
            Else
                isLaptop = False
                UpdateStatus "Desktop detected"
            End If
        On Error GoTo 0
        
        If siteServer = "" Then siteServer = LogonServer
        
        If Not isLaptop Then
            ' This is for desktop users only
            bReturnVal = DriveExists("H:")
            If bReturnVal Then ' H: mapping exists.  Remove it.
                bReturnVal = RemoveDrive("H:", 0)      ' Mode 0 is SUBST
                If Not bReturnVal Then
                    bReturnVal = RemoveDrive("H:", 1)  ' Mode 1 is API
                    If bReturnVal Then
                        UpdateStatus "Removed H: with API Call"
                    Else
                        UpdateStatus "Unable to remove existing H:"
                    End If
                Else
                    UpdateStatus "Removed H: with SUBST H: /D"
                End If
            End If
        End If
        ' This is for desktop and laptop
        If FolderExists(siteServer & "\" & Username & "$") Then
            If isLaptop Then
                UpdateStatus "Mapping G: to " & siteServer & "\" & Username & "$ (API)"
                retval = Connect2("G:", siteServer & "\" & Username & "$", Username, "")
            Else
                UpdateStatus "Mapping H: to " & siteServer & "\" & Username & "$ (API)"
                retval = Connect2("H:", siteServer & "\" & Username & "$", Username, "")
            End If
        Else
            If Not isLaptop Then
                If FolderExists(siteServer & "\mobile$") Then
                    If FolderExists(siteServer & "\mobile$\" & Username) Then
                        ' Do nothing (yet)
                    Else
                        ' Create Directory
                        bReturnVal = CreateDir(siteServer & "\mobile$\" & Username)
                        If bReturnVal Then
                            ' Apply permissions to directory
                            UpdateStatus "Applying permissions to " & siteServer & "\mobile$\" & Username
                            Shell "cmd /c echo y|c:\winnt\system32\cacls.exe " & siteServer & "\mobile$\" & Username & " /C /T /G " & Username & ":C", vbHide
                        End If
                    End If
                    ' create drive mapping using SUBST
                    UpdateStatus "Mapping H: to " & siteServer & "\mobile$\" & Username & " (SUBST H: " & siteServer & "\mobile$\" & Username & ")"
                    bReturnVal = SubstPath("H:", siteServer & "\mobile$\" & Username)
                    If bReturnVal Then
                        UpdateStatus "SUBST of H: SUCCESSFUL"
                    Else
                        UpdateStatus "SUBST of H: UNSUCCESSFUL.. ERROR!"
                    End If
                    
                    On Error Resume Next
                    ' Create 'descriptive' text file in H: (zero length file)
                    Open "H:\This is not your normal Home Drive" For Output As #4
                    If Err.Number <> 0 Then
                        UpdateStatus "Unable to write to H:"
                    Else
                        Close 4
                    End If
                    On Error GoTo ErrorLog
                    
                    ' Display message explaining H:
                    msg = ""
                    msg = msg & "Traveler Warning:" & vbCrLf
                    msg = msg & vbCrLf
                    msg = msg & "This is not your designated home site." & vbCrLf
                    msg = msg & vbCrLf
                    msg = msg & "Do not store any important information" & vbCrLf
                    msg = msg & "on your H: drive at this location as it may" & vbCrLf
                    msg = msg & "not be available later." & vbCrLf
                    ' Logon 34 modification
                    'MsgBox msg, vbOKOnly + vbInformation, "Logon Application"
                    retval = UpdateFile("C:\temp\Traveler.exe", LogonServer & "\NETLOGON\Setup\Traveler.exe")
                    If bReturnVal Then
                        Shell "C:\temp\Traveler.exe " & msg, vbNormalFocus
                    Else
                        UpdateStatus "Unable to execute Traveler.exe"
                    End If
                    msg = ""
                    
                    Traveller = True
                    
                    ' This next message is for info only
                    UpdateStatus "---The current H: may only be disconnected by performing a SUBST H: /D from a command prompt---"
                Else
                    ' mobile$ share does not exist on siteserver.
                    UpdateStatus siteServer & "\mobile$ does not exist.  This is a CRITICAL error."
                    ErrorNumber = ErrorNumber Or 65536
                    CriticalError = True
                End If
            End If
        End If
    End If
    UpdateStatus "Saving Traveller info to registry"
                     
    If Traveller Then
        SaveSetting "WAPS", "Logon", "Traveller", ",1"
    Else
        SaveSetting "WAPS", "Logon", "Traveller", ",0"
    End If
    LogLocal
    ' ****************************************
    ' End Modification
    
    
    ' ****************************************
    'Logon 34 modification
    
    ' Modify environment variable HOMEDIR - 2 steps
    ' 1.  Modify registry
    lngReturnval = SetValueString("", "Environment", "HOMEDIR", "H\", REG_SZ)
    
    ' 2.  Send Windows Message to system, so that all top level windows update themselves.. currently this only works after logoff/logon
    '     This is not the desired result, but what was possible in the time given.
    lParam = "Environment"
    lngTemp = SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, StrPtr(lParam), SMTO_ABORTIFHUNG, 5000, lngReturnval)
    
    strWinDir = Environ$("WINDIR")
    ' ****************************************
    
    LogLocal

    UpdateStatus "Computername = " & Computername
    UpdateStatus "Username     = " & Username
    UpdateStatus "Logonserver  = " & Right(LogonServer, Len(LogonServer) - 2)
    If siteServer <> "" Then
        UpdateStatus "Siteserver   = " & Right(siteServer, Len(siteServer) - 2)
    Else
        UpdateStatus "Siteserver   = UNKNOWN"
    End If
    UpdateStatus "IP Address   = " & IPAddress
    UpdateStatus "Client Type  = " & Client
    LogLocal

    If DebugMode = False Then
        On Error GoTo ErrorLog
    End If
    UpdateStatus "APIMode      = " & APIMode
    UpdateStatus "API          = " & API
    UpdateStatus "Modechange   = " & ModeChange
    UpdateStatus "CriticalOnly = " & CriticalOnly
    DebugMsg "Debug Mode ON"
    ' If the groups.txt file already exists, delete it
    If Dir("c:\temp\groups.txt") <> "" Then
        Kill "c:\temp\groups.txt"
    End If
    
    ' Internal versioning as previously defined
V31:
    If Version = 31 Then
        ' Have to create enum.kix to enumerate groups
        If Dir("c:\temp\enum.kix") <> "" Then
            Kill "c:\temp\enum.kix"
            DebugMsg "Deleted existing enum.kix"
        End If
        DebugMsg "Creating enum.kix in c:\temp"
        Open "c:\temp\enum.kix" For Output As #2
            Print #2, "SETCONSOLE (""Hide"")"
            Print #2, "$tempfile = ""c:\temp\tempfile.txt"""
            Print #2, "$groupfile = ""c:\temp\groups.txt"""
            Print #2, "$CRLF = chr(13) + chr(10)"
            Print #2, "DEL $tempfile"
            Print #2, "$err = OPEN (1,$tempfile,5)"
            Print #2, "$Index = 0"
            Print #2, "DO"
            Print #2, "    $Group = UCASE(ENUMGROUP($Index))"
            Print #2, "    IF $group <> """""
            Print #2, "        $err = WRITELINE (1, $group)"
            Print #2, "    ENDIF"
            Print #2, "    IF $group <> """""
            Print #2, "        $err = WRITELINE (1, $CRLF)"
            Print #2, "    ENDIF"
            Print #2, "    $index = $index + 1"
            Print #2, "UNTIL LEN ($Group) = 0"
            Print #2, "$err = CLOSE (1)"
            Print #2, "COPY $tempfile $groupfile"
            Print #2, "DEL $tempfile"
        Close 2
        cmd = KIXpath & " C:\TEMP\ENUM.KIX"
        If Dir(KIXpath) <> "" Then
            Shell cmd, vbHide
            DebugMsg "Executed enum.kix with KIX32.exe"
        Else
            UpdateStatus "Cannot find " & KIXpath
            ErrorNumber = ErrorNumber Or 8192
            CriticalError = True
            If Dir("c:\temp\kix32.exe") <> "" Then      ' Found Kix on c:\temp
                UpdateStatus "Found KIX32.exe on c:\temp, copying to c:\winnt."
                FileCopy "c:\temp\kix32.exe", KIXpath   ' Try to copy it to c:\winnt directory
                If Dir(KIXpath) <> "" Then              ' Copy successful
                    UpdateStatus "Copy successful."
                    Shell cmd, vbHide
                    DebugMsg "Executed enum.kix with KIX32.exe"
                Else
                    UpdateStatus "Unable to copy KIX32.exe to c:\winnt"
                    If ModeChange = True Then
                        UpdateStatus "Upgrading to Version 3.2 mode"
                        Version = 32
                        ModeChange = False
                    Else
                        UpdateStatus "Unable to upgrade to Version 3.2 mode (Modechange FALSE)"
                        EndProgram
                    End If
                End If
            Else                                    ' Can't find KIX32.exe, so copy it to c:\temp
                UpdateStatus "KIX32.exe not found on c:\temp, copying from " & LogonServer & "\NETLOGON\SETUP"
                FileCopy LogonServer & "\NETLOGON\SETUP\KIX32.exe", "c:\temp\kix32.exe"
                If Dir("c:\temp\kix32.exe") <> "" Then
                    cmd = "c:\temp\kix32.exe c:\temp\enum.kix"
                    Shell cmd, vbHide
                    DebugMsg "Executed enum.kix with c:\temp\KIX32.exe"
                Else
                    UpdateStatus "Unable to copy KIX32.exe to c:\temp."
                    EndProgram
                End If
            End If
        End If
    End If
    
V32:
    If Version = 32 Then
        Source = LogonServer & "\Netlogon\Users\" & Username & ".txt"
        Destination = "c:\temp\groups.txt"
        d = Dir(Source)
        If d = "" Then
            CriticalError = True
            ErrorNumber = ErrorNumber Or 32768
            If ModeChange = True Then
                UpdateStatus "Reverting to Version 3.1 mode"
                Version = 31
                ModeChange = False
                GoTo V31
            Else
                UpdateStatus "Unable to revert to Version 3.1 mode"
                EndProgram
            End If
        End If
        FileCopy Source, Destination
    End If
    
    LogLocal
    
    If siteLoc = "SITELOC" Or siteLoc = "" Then
        MsgBox "This workstation has not been properly configured." & vbCrLf & vbCrLf & "The SITELOC environment variable is not set and this PC will not be able to logon correctly." & vbCrLf & vbCrLf & "Please contact the helpdesk on 1800 651 327.", vbCritical, "Logon Error"
        ErrorNumber = ErrorNumber Or 1
        UpdateStatus "Logon halted due to blank or default SITELOC environment variable :" & siteLoc
        LogLocal
        CriticalError = True
        EndProgram
    End If
    If Right(siteLoc, 1) <> "X" And Mid(siteLoc, Len(siteLoc) - 1, 1) <> "X" Then ' Error
        MsgBox "This workstation has not been properly configured." & vbCrLf & vbCrLf & "The SITELOC environment variable is not a valid format and this PC will not be able to logon correctly." & vbCrLf & vbCrLf & "Please contact the helpdesk on 1800 651 327.", vbCritical, "Logon Error"
        UpdateStatus "Logon halted due to incorrect configuration of SITELOC environment variable"
        ErrorNumber = ErrorNumber Or 2
        LogLocal
        CriticalError = True
        EndProgram
    End If
    
    UpdateStatus "SITELOC environment variable valid (" & siteLoc & ")"
    LogLocal
    
    If UCase(PRTLOC) = "PRTLOC" Or UCase(PRTLOC) = "" Then ' Error
        UpdateStatus "Invalid PRTLOC environment variable (" & PRTLOC & ")"
        ErrorNumber = ErrorNumber Or 4
    End If
    
    strTemp2 = Mid(UCase(Computername), Len(Computername) - 1, 1) ' Get secondlast character in computername
    If InStr(1, UCase(Computername), "X") > 0 Then                ' Does computername contain an X ?
        ServerFlag = True                                         ' Yes.  Computer is a server
    Else                                                          ' No.  Not a server
        If UCase(Left(Computername, 1)) <> "C" Then               ' Workstation name formatted correctly?
            If Dir("c:\temp\cscmachine.txt") = "" Then
                ErrorNumber = ErrorNumber Or 8
            Else
                'CSC computer.  NOT an error
            End If
        Else
        End If
        ServerFlag = False
    End If
    UpdateStatus "Completed computername diagnostics pass 1 "
    LogLocal
    If strTemp2 = "W" Then
        ServerFlag = False             ' Don't consider High-End workstations to be 'Servers'
        HighEnd = True
        UpdateStatus "This is a Highend Workstation"
    End If
    
    
    If ServerFlag = True Then
        retval = MsgBox("Do you wish to to run the Logon Application on this Server?" & vbCrLf & vbCrLf & "This may cause problems." & vbCrLf & vbCrLf, vbExclamation + vbYesNo + vbDefaultButton2, "Confirmation required")
        If retval = vbYes Then
            ' complete remainder of Logon Application
            UpdateStatus "User confirmed that they did want to run Logon on this server"
        Else
            UpdateStatus "User confirmed that they did NOT want to run Logon on this server"
            ErrorNumber = ErrorNumber Or 256
            LogLocal
            EndProgram ' OK then, REALLY stop Logon process.
        End If
    End If
    UpdateStatus "Completed computername diagnostics pass 2"
    LogLocal
    
    
    ' Logon33 modification
    DriveLetter(1) = "P:"
    DriveLetter(2) = "T:"
    DriveLetter(3) = "N:"
    DriveLetter(4) = "X:"
    DriveLetter(5) = "O:"
    DriveLetter(6) = "Q:"
        

    If Not Traveller Then
        ' lstGroups is used to temporarily hold the list of NT Global Groups the user is a member of
        ' Don't forget that this listbox is sorted
        lstGroups.Clear
        
        Ticks(1) = GetTickCount
        UpdateStatus "Time 1: " & Ticks(1) - StartTick & " mS"
        
        DebugMsg "Checking for existence of c:\temp\groups.txt"
        '
        ' It is possible, (but HIGHLY unlikely) that the .exe will just sit here and wait for the
        ' groups.txt file to appear - for up to 3 minutes.
        ' In this case, ensure that the c:\winnt\kix32.exe file is not corrupt.
        
        UpdateStatus "Waiting for c:\temp\groups.txt (max of 3 minutes)"
        d = Dir("c:\temp\groups.txt")
        t = Time
        Do Until d <> "" Or (ErrorNumber And 16384) = True
            d = Dir("c:\temp\groups.txt")
            'DoEvents
            If d = "" Then
                If Time > DateAdd("n", 3, t) Then ' Wait up to 3 minutes
                    ErrorNumber = ErrorNumber Or 16384
                    CriticalError = True
                    UpdateStatus "Disk Space on C: = " & FormatNumber(fsoDrive.FreeSpace / 1024, 0) & " Kb"
                    UpdateStatus "c:\temp\groups.txt NOT FOUND.  Logon cannot continue."
                    If Version = 31 Then
                        UpdateStatus "Ensure that C:\WINNT\KIX32.EXE exists, and is not corrupt."
                    End If
                    EndProgram
                End If
            End If
        Loop
        UpdateStatus "Retrieving group info from c:\temp\groups.txt"
        Open "c:\temp\groups.txt" For Input As #50
            Do Until EOF(50)
                Line Input #50, GroupTempVar
                GroupTempVar = UCase(Trim(GroupTempVar))
                lstGroups.AddItem GroupTempVar
                DebugMsg "Found " & GroupTempVar
            Loop
        Close 50
        
        If DebugMode = False Then
            UpdateStatus "Deleting c:\temp\groups.txt"
            Kill "c:\temp\groups.txt"
        Else
            DebugMsg "Keeping c:\temp\groups.txt for reference"
        End If
        
        LogLocal
        
        Ticks(2) = GetTickCount
        UpdateStatus "Time 2: " & Ticks(2) - Ticks(1) & " mS"
        
        varFilename = LogonServer & "\netlogon\sites\" & siteLoc & ".csv"
        
        UpdateStatus "Opening " & varFilename & " for processing"
        LogLocal
        
        UpdateStatus "Remote filename = " & varFilename
        ' Check if csv is already on local drive
        LocalFile = "c:\temp\" & siteLoc & ".csv"
        UpdateStatus "Local Filename = " & LocalFile
        dr = Dir(LocalFile)
        If dr = "" Then
            ' File not found, so copy it
            UpdateStatus "File not found - copying to local machine"
            UpdateStatus "Copy started"
            LogLocal
            'On Error Resume Next
            FileCopy varFilename, LocalFile
            UpdateStatus "Copy completed"
            E = Err.Number
            If E <> 0 Then
                UpdateStatus siteLoc & ".csv not found on Network.  No local version, so Logon cannot proceed."
                ErrorNumber = ErrorNumber Or 1024
                LogLocal
                CriticalError = True
                EndProgram
            End If
            'On Error GoTo 0
        Else
            UpdateStatus "Found file locally - checking filetime"
            ' file found, check if current
            dr = Dir(varFilename)
            If dr = "" Then
                ' Network version of SITELOC.csv NOT FOUND!!!
                UpdateStatus siteLoc & ".csv not found on Network.  Using local (possibly out of date) version."
                ErrorNumber = ErrorNumber Or 16
                LogLocal
            Else
                fileinfo1 = FileDateTime(varFilename)
                fileinfo2 = FileDateTime(LocalFile)
                varDiff = DateDiff("s", fileinfo1, fileinfo2)
                ' If the difference between these two files is greater than 30 seconds, then replace it
                If (varDiff >= 30 Or varDiff <= 30) And varDiff <> 0 Then
                    UpdateStatus "Local file is invalid - copying locally"
                    UpdateStatus "Copy started"
                    FileCopy varFilename, LocalFile
                    UpdateStatus "Copy completed"
                Else
                    UpdateStatus "Local file is up to date"
                End If
            End If
        End If
        
        ' Process file (on local machine) to collect drive mapping info
        UpdateStatus "Processing of local file (" & LocalFile & ") started"
        LogLocal
        
        Open LocalFile For Input As #60
            Line Input #60, header
            disconnectErrorflag = 0
            connectErrorflag = 0
            printerErrorflag = 0
            groupNumber = 1
            printerNumber = 1
            Do Until EOF(60)
                Input #60, csvGroupname(groupNumber)
                csvGroupname(groupNumber) = UCase(csvGroupname(groupNumber))
                If csvGroupname(groupNumber) <> "" Then
                    If csvGroupname(groupNumber) = "TRAVELLER" Then
                        Input #60, csvPDrive(groupNumber), csvTDrive(groupNumber), csvNDrive(groupNumber), csvXDrive(groupNumber), csvODrive(groupNumber), csvQDrive(groupNumber)
                        csvPDrive(groupNumber) = UCase(csvPDrive(groupNumber))
                        csvTDrive(groupNumber) = UCase(csvTDrive(groupNumber))
                        csvNDrive(groupNumber) = UCase(csvNDrive(groupNumber))
                        csvXDrive(groupNumber) = UCase(csvXDrive(groupNumber))
                        csvODrive(groupNumber) = UCase(csvODrive(groupNumber))
                        csvQDrive(groupNumber) = UCase(csvQDrive(groupNumber))
                        ' Additions for Traveller drive mappings (Logon 32 BUILD 29)
                        ' **************************************
                        'If UCase(csvPDrive(groupNumber)) <> "NA" Then travellerPDrive = UCase(csvPDrive(groupNumber))
                        'If UCase(csvTDrive(groupNumber)) <> "NA" Then travellerTDrive = UCase(csvTDrive(groupNumber))
                        'If UCase(csvNDrive(groupNumber)) <> "NA" Then travellerNDrive = UCase(csvNDrive(groupNumber))
                        'If UCase(csvXDrive(groupNumber)) <> "NA" Then travellerXDrive = UCase(csvXDrive(groupNumber))
                        'If UCase(csvODrive(groupNumber)) <> "NA" Then travellerODrive = UCase(csvODrive(groupNumber))
                        'If UCase(csvQDrive(groupNumber)) <> "NA" Then travellerQDrive = UCase(csvQDrive(groupNumber))
                        ' **************************************
                        Input #60, csvPrinter
                        travellerPrinter(1) = UCase(csvPrinter)
                        Input #60, csvPrinter
                        travellerPrinter(2) = UCase(csvPrinter)
                        Input #60, csvPrinter
                        travellerPrinter(3) = UCase(csvPrinter)
                        DebugMsg "Extracted traveller paths " & csvGroupname(groupNumber) & "," & csvPDrive(groupNumber) & "," & csvTDrive(groupNumber) & "," & csvNDrive(groupNumber) & "," & csvXDrive(groupNumber) & "," & csvODrive(groupNumber) & "," & csvQDrive(groupNumber) & "," & travellerPrinter(1) & "," & travellerPrinter(2) & "," & travellerPrinter(3)
                    Else
                        Input #60, csvPDrive(groupNumber), csvTDrive(groupNumber), csvNDrive(groupNumber), csvXDrive(groupNumber), csvODrive(groupNumber), csvQDrive(groupNumber)
                        csvPDrive(groupNumber) = UCase(csvPDrive(groupNumber))
                        csvTDrive(groupNumber) = UCase(csvTDrive(groupNumber))
                        csvNDrive(groupNumber) = UCase(csvNDrive(groupNumber))
                        csvXDrive(groupNumber) = UCase(csvXDrive(groupNumber))
                        csvODrive(groupNumber) = UCase(csvODrive(groupNumber))
                        csvQDrive(groupNumber) = UCase(csvQDrive(groupNumber))
                        Input #60, csvPrinter1(groupNumber), csvPrinter2(groupNumber), csvPrinter3(groupNumber)
                        csvPrinter1(groupNumber) = UCase(csvPrinter1(groupNumber))
                        csvPrinter2(groupNumber) = UCase(csvPrinter2(groupNumber))
                        csvPrinter3(groupNumber) = UCase(csvPrinter3(groupNumber))
                        DebugMsg "Extracted paths " & csvGroupname(groupNumber) & "," & csvPDrive(groupNumber) & "," & csvTDrive(groupNumber) & "," & csvNDrive(groupNumber) & "," & csvXDrive(groupNumber) & "," & csvODrive(groupNumber) & "," & csvQDrive(groupNumber) & "," & csvPrinter1(groupNumber) & "," & csvPrinter2(groupNumber) & "," & csvPrinter3(groupNumber)
                        groupNumber = groupNumber + 1
                    End If
                End If
            Loop
        Close 60
        UpdateStatus "Completed processing " & LocalFile
    
        Ticks(3) = GetTickCount
        UpdateStatus "Time 3: " & Ticks(3) - Ticks(2) & " mS"
        LogLocal
        ' Now match user's groups to .csv info
        groupNumber = groupNumber - 1
        printerNumber = printerNumber - 1
        pNo = 1
        ' Traveller = True ' Logon 33 modification
        For lp2 = 1 To groupNumber
            For lp = 0 To lstGroups.ListCount - 1
                'DoEvents
                Group = lstGroups.List(lp)
                If csvGroupname(lp2) = Group Then
                    UpdateStatus "User found in group (" & Group & ")"
                    'Traveller = False
                    If csvPDrive(lp2) <> "NA" Then DriveToMap(1) = csvPDrive(lp2)
                    If csvTDrive(lp2) <> "NA" Then DriveToMap(2) = csvTDrive(lp2)
                    If csvNDrive(lp2) <> "NA" Then DriveToMap(3) = csvNDrive(lp2)
                    If csvXDrive(lp2) <> "NA" Then DriveToMap(4) = csvXDrive(lp2)
                    If csvODrive(lp2) <> "NA" Then DriveToMap(5) = csvODrive(lp2)
                    If csvQDrive(lp2) <> "NA" Then DriveToMap(6) = csvQDrive(lp2)
                    PrinterToMap(pNo) = csvPrinter1(lp2)
                    PrinterToMap(pNo + 1) = csvPrinter2(lp2)
                    PrinterToMap(pNo + 2) = csvPrinter3(lp2)
                    pNo = pNo + 3
                End If
            Next lp
        Next lp2
        UpdateStatus "Completed matching user's groups to .csv info"
        LogLocal
        ' NOTE:  Traveller logic has changed BIG TIME!!!!
        ' ' A traveller is now defined as a user being at a site that does not host their REAL H: Drive
        ' ------ OLD INFO.....Traveller will still be 'True' (it is a boolean), if the user has no drives to map
        ' ------ OLD INFO.....ie the global groups he is a member of do not appear in this sites .csv file
        ' ***********************
        ' Logon 33 modifications
        'If Traveller = True Then
    Else
        DriveToMap(1) = ""
        DriveToMap(2) = siteServer & "\Corporate"
        DriveToMap(3) = siteServer & "\Applications"
        DriveToMap(4) = siteServer & "\Transfer"
        DriveToMap(5) = siteServer & "\General"
        DriveToMap(6) = ""
        PrinterToMap(1) = travellerPrinter(1)
        PrinterToMap(2) = travellerPrinter(2)
        PrinterToMap(3) = travellerPrinter(3)
        pNo = 3
        UpdateStatus "This user is a Traveller"
    End If
    ' ***********************
    
    Ticks(4) = GetTickCount
    UpdateStatus "Time 4: " & Ticks(4) - Ticks(3) & " mS"
    UpdateStatus "Mapping of Drives and printers started"
    
    ' Modify the following as you see fit
    'If InternalTesting = True Then
    '    ' ************************** TEST ONLY ************************
    '    DriveToMap(1) = "\\CBDXAAP\C$"
    '    DriveToMap(2) = "\\PERTHXWAA\C$"
    '    DriveToMap(3) = "\\CBDXAAG\Applications"
    '    DriveToMap(4) = "\\CBDXAAG\Corporate"
    '    DriveToMap(5) = "\\CBDXAAG\Transfer"
    '    ' *************************************************************
    'End If
    
    If DebugMode = True Then
        For lp = 1 To 6
            UpdateStatus "Drives to map - " & DriveLetter(lp) & " " & DriveToMap(lp)
        Next lp
    End If
    ' Resolve Hostname(s) using DNS / WINS
    UpdateStatus "Resolving Server names to IP Addresses"
    For lp = 1 To 6
        If DriveToMap(lp) <> "" Then
            pos = InStr(3, DriveToMap(lp), "\")
            Servername = Mid(DriveToMap(lp), 3, pos - 3)
            ServerIP = ResolveHostname(Servername)
            If ServerIP <> "" Then
                If APIMode = True Then
                    'No API mode option here - always use mode 2
                    UpdateStatus "Connecting to IPC$ on " & Servername & " (API)"
                    retval = Connect2("", "\\" & Servername & "\IPC$", Username, "")
                    If retval <> 0 Then UpdateStatus "" & msg & " - Error " & retval & " (" & ErrorMsg & ")"
                Else
                    cmd = "net use \\" & Servername & "\IPC$"
                    UpdateStatus "Connecting to IPC$ on " & Servername & " (Net use)"
                    Shell cmd, vbHide
                End If
            End If
        End If
    Next lp
    UpdateStatus "Done resolving names and creating sessions to server(s)."
    LogLocal
    If APIMode = False Then ' Net Use only
        UpdateStatus "Using 'Net Use' to map drives"
        For lp = 1 To 6
            If DriveToMap(lp) <> "" Then
                cmd = "echo Delete at any time > c:\temp\deleteme" & lp & ".txt"
                cmd = cmd & vbCrLf & "net use " & DriveLetter(lp) & " /d"
                cmd = cmd & vbCrLf & "net use " & DriveLetter(lp) & " " & DriveToMap(lp)
                cmd = cmd & vbCrLf & "del c:\temp\deleteme" & lp & ".txt"
                cmd = cmd & vbCrLf & "exit"
                fnum = FreeFile
                ' Remove previous files if they exist (from previous logon's)
                If Dir("c:\temp\mapdrive" & lp & ".bat") <> "" Then Kill "c:\temp\mapdrive" & lp & ".bat"
                Open "c:\temp\mapdrive" & lp & ".bat" For Output As #fnum
                    Print #fnum, cmd
                Close fnum
                ' The following loop is in place to ensure that the mapdriveX.bat file exists
                ' before trying to run it
                Do Until Dir("c:\temp\mapdrive" & lp & ".bat") <> ""
                    'DoEvents
                    UpdateStatus "Waiting for mapdrive" & lp & ".bat"
                Loop
                UpdateStatus "Disconnecting " & DriveLetter(lp) & " (Net use /D)"
                UpdateStatus "Connecting " & DriveLetter(lp) & " to " & DriveToMap(lp) & " (Net use)"
                retval = Shell("c:\temp\mapdrive" & lp & ".bat", vbHide)
            End If
        Next lp
    Else ' API
        UpdateStatus "Using API to map drives"
         ' Now that we have the paths to map, map them, disconnecting drives first
        For lp = 1 To 6
           try = 0
           driveError = False
            If DriveToMap(lp) <> "" And DriveToMap(lp) <> "NA" Then
                ' API call to disconnect
                If API = 1 Then
                    retval = DisConnect(DriveLetter(lp))
                End If
                If API = 2 Then
                    retval = DisConnect2(DriveLetter(lp), True)
                End If
                msg = "Disconnecting " & DriveLetter(lp)
                t = CStr(Time)
                Do Until fso.DriveExists(DriveLetter(lp)) = False Or DateDiff("s", t, Time) >= 5
                    'DoEvents
                Loop
                If retval <> 0 Then
                    If retval <> 2250 Then
                        UpdateStatus "" & msg & " - Error " & ErrorNum & " (" & ErrorMsg & ")"
                        disconnectErrorflag = 1
                    Else
                        ' This drive does not exist.  Don't worry about it
                    End If
                Else
                    UpdateStatus "" & msg & " - Success"
                End If
                LogLocal
                Do Until try > 2
                    retval = 0
                    If API = 1 Then
                        retval = Connect(DriveLetter(lp), DriveToMap(lp))
                    End If
                    If API = 2 Then
                        retval = Connect2(DriveLetter(lp), DriveToMap(lp), Username, "")
                    End If
                    'DebugMsg "Return Code from Connect = " & retval & " (" & ErrorMsg & ")"
                    msg = "Connecting " & DriveLetter(lp) & " to " & DriveToMap(lp)
                    If retval <> 0 Then
                        try = try + 1 ' Unsuccessful drive mapping attempt
                        driveError = True
                        If retval = ERROR_ALREADY_ASSIGNED Then
                            UpdateStatus "" & msg & " - Drive letter in use"
                        Else
                            UpdateStatus "" & msg & " - Error " & retval & " (" & ErrorMsg & ")"
                            connectErrorflag = 1
                        End If
                    Else
                        UpdateStatus "" & msg & " - Success"
                        driveError = False
                        try = 4 ' Successful drive mapping, so don't retry
                    End If
                    If driveError = True Then UpdateStatus "Retrying to map drive (" & try & " of 3)"
                    Err.Clear
                Loop
                LogLocal
                If driveError = True Then
                    ' Still can't map drive after 2 attempts using API so use a net use command
                    cmd = "net use " & DriveLetter(lp) & " " & DriveToMap(lp)
                    retval = Shell(cmd, vbHide)
                    UpdateStatus "Executed a net use command to map " & DriveToMap(lp)
                    End If
            End If
        Next lp
        If disconnectErrorflag = 1 Then ErrorNumber = ErrorNumber Or 32
        If connectErrorflag = 1 Then ErrorNumber = ErrorNumber Or 64
    End If
    LogLocal
    
    ' ****************************************************
    ' This is the Greg Appleby GUP additions
    ' GUP = Generate User Profile
    
    ' Logon34 mod 6-8-01
    
    ' NOTE... removed due to time constraints
    
    ' Set up base keys
'    CreateKey "", "Software\Microsoft", "Office"
'    CreateKey "", "Software\Microsoft\Office", "9.0"
'    CreateKey "", "Software\Microsoft\Office\9.0", "Access"
'    CreateKey "", "Software\Microsoft\Office\9.0", "Excel"
'    CreateKey "", "Software\Microsoft\Office\9.0", "Outlook"
'    CreateKey "", "Software\Microsoft\Office\9.0\Outlook", "Journal"
'    CreateKey "", "Software\Microsoft\Office\9.0\Outlook", "Setup"
'    CreateKey "", "Software\Microsoft\Office\9.0", "PowerPoint"
'    CreateKey "", "Software\Microsoft\Office\9.0", "Word"
'    CreateKey "", "Software\Microsoft\Office\9.0", "Common"
'    CreateKey "", "Software\Microsoft\Office\9.0\Common", "General"
'    CreateKey "", "Software\Microsoft\Office\9.0\Common", "UserInfo"
'
'    ' Insert values
'    SetValueString "", "Software\Microsoft\Office\9.0\Access", "UserData", "1", REG_DWORD
'    SetValueString "", "Software\Microsoft\Office\9.0\Common", "UserData", "1", REG_DWORD
'    SetValueString "", "Software\Microsoft\Office\9.0\Common\General", "FirstRun", "0", REG_DWORD
'    'SetValueString "", "Software\Microsoft\Office\9.0\Common\UserInfo", "UserInitials", "50000000", REG_BINARY
'    'SetValueString "", "Software\Microsoft\Office\9.0\Common\UserInfo", "UserName", "%USRINFO%", REG_BINARY
'    'SetValueString "", "Software\Microsoft\Office\9.0\Common\UserInfo", "Company", "57004100200050006F006C00690063006500200053006500720076006900630065000000", REG_BINARY
'    SetValueString "", "Software\Microsoft\Office\9.0\Excel", "UserData", "1", REG_DWORD
'    SetValueString "", "Software\Microsoft\Office\9.0\Outlook", "UserData", "1", REG_DWORD
'    'SetValueString "", "Software\Microsoft\Office\9.0\Outlook", "Settings", "1", REG_BINARY
'    SetValueString "", "Software\Microsoft\Office\9.0\Outlook", "Machine Name", Computername, REG_SZ
'    SetValueString "", "Software\Microsoft\Office\9.0\Outlook\Journal", strWinDir & "\Profiles\" & Username & "\Application Data\Microsoft\Outlook\outitems.log", "Outlook Item Log File", REG_SZ
'    SetValueString "", "Software\Microsoft\Office\9.0\Outlook\Journal", strWinDir & "\Profiles\" & Username & "\Application Data\Microsoft\Outlook\offitems.log", "Item Log File", REG_SZ
'    SetValueString "", "Software\Microsoft\Office\9.0\Outlook\Setup", "MailSupport", "1", REG_DWORD
'    SetValueString "", "Software\Microsoft\Office\9.0\PowerPoint", "UserData", 1, REG_DWORD
'    SetValueString "", "Software\Microsoft\Office\9.0\Word", "UserData", 1, REG_DWORD
    
    
    ' ****************************************************
    ' Moved from within EndProgram in Logon33 mod
    If Dir("N:\apps\MAPIProfile\Mapiprof.bat") <> "" Then ' Execute mapiprof.bat
        UpdateStatus "Shelled out to MAPIProf.bat"
        Shell "cmd /c N:\apps\MAPIProfile\Mapiprof.bat", vbHide
    Else
        ' FILE NOT FOUND!
        If DriveExists("N:") Then ' N: is mapped
            UpdateStatus "MAPIProf.bat not found"
        Else ' N: not mapped
            UpdateStatus "N: not mapped, unable to execute MAPIProf.bat"
        End If
    End If
                    
    prtlocError = 0
    For lp = 1 To pNo
        If PrinterToMap(lp) <> "NA" And PrinterToMap(lp) <> "" Then
            msg = "Mapping printer "
            If (PRTLOC = "" Or PRTLOC = "PRTLOC") And InStr(1, PrinterToMap(lp), "%PRTLOC%") <> 0 Then
                prtlocError = 1
                msg = msg & "- cannot map due to invalid PRTLOC (" & PRTLOC & ")"
            Else
                PrinterToMap(lp) = Replace(PrinterToMap(lp), "%PRTLOC%", PRTLOC)
                varResult = AddPrinterConnection(PrinterToMap(lp))
                If varResult = 0 Then
                    msg = msg & "to " & PrinterToMap(lp) & " - Error"
                    If HighEnd = False Then
                        printerErrorflag = 1
                    End If
                Else
                    msg = msg & "to " & PrinterToMap(lp) & " - Success"
                End If
            End If
                UpdateStatus msg
        End If
    Next lp
    
    If printerErrorflag = 1 Then ErrorNumber = ErrorNumber Or 128
    If prtlocError = 1 Then ErrorNumber = ErrorNumber Or 2048
    LogLocal
    ' This code added 26 Oct 2000 to overcome intermittent drives not being mapped by API
    For lp = 1 To 6
        If DriveToMap(lp) <> "" Then
            If Not (fso.DriveExists(DriveLetter(lp))) Then
                cmd = "net use " & DriveLetter(lp) & " " & DriveToMap(lp)
                retval = Shell(cmd, vbHide)
                UpdateStatus "Retried to map " & DriveLetter(lp) & " to " & DriveToMap(lp)
            End If
        End If
    Next lp
    
    Set fso = Nothing
    Ticks(5) = GetTickCount
    UpdateStatus "Time 5: " & Ticks(5) - Ticks(4) & " mS"
    
    ' Job done.  Logon complete at this point
    lblMessage = "Logon Complete"
    UpdateStatus "Logon process complete"
    FinishTime = Time
    LogLocal
    frmMain.Caption = "Logon Complete"
    DebugMsg "System BEEP"
    Beep
    DebugMsg "Playing 'The Microsoft Sound.wav'"
    PlayWav
    LogLocal
    EndProgram
    ' This next portion is only used if there is an internal error
ErrorLog:
    erDesc = Err.Description
    erNum = Err.Number
    Err.Clear
    On Error Resume Next
    If Dir("c:\temp\" & Username & ".log") <> "" Then
        Kill "c:\temp\" & Username & ".log"
    End If
    Open "c:\temp\" & Username & ".log" For Output As #1 ' Output log file on local machine
        For lp = 0 To lstStatus.ListCount - 1
            Print #1, lstStatus.List(lp)
        Next lp
        UpdateStatus "Internal Error " & erNum & " (" & erDesc & ")"
        ErrorNumber = ErrorNumber Or 4096
        CriticalError = True
    Close 1
    EndProgram
End Sub

Private Sub cmdClose_Click()
    DebugMsg "Started cmdClose_Click() Sub"
    ' User clicked the Close button, but we don't automatically close the application
    Dim retval As Long
    Dim msg As String
    retval = MsgBox("Do you really want to cancel your Windows NT Logon?" & vbCrLf & vbCrLf & "This will prevent all or some of your Network drives from being available.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation Required")
    If retval = vbYes Then
        ' The user has confirmed that they want to cancel their Logon.
        MsgBox "Logon cancelled at user's request"
        UpdateStatus "Logon cancelled at user's request"
        ErrorNumber = ErrorNumber Or 512
        LogLocal
        EndProgram
    End If
End Sub

Public Sub EndProgram()
    Dim TotalTime As String, msg As String, SlowLogon As Integer, TickNow As Long
    Dim retval As Long
    Dim t As Date, d As String
    ' This sub is kicked off when it's time to end the program for any reason
    DebugMsg "Started EndProgram() Sub"
    
    SlowLogon = 0
    
'    ' **********************
'    ' Logon 33 modification
'
'    isLaptop = False
'
'    ' Ensure neither error 256 or error 512
'    '    Error 512 is logon cancelled (user clicked on 'Close')
'    '    Error 256 is logon cancelled (server)
'    If (ErrorNumber And 512) = 0 And (ErrorNumber And 256) = 0 Then
'        UpdateStatus "Saving Traveller info to registry"
'         If Traveller Then
'            SaveSetting "WAPS", "Logon", "Traveller", ",1"  ' <-- NOTE: removal of version from 'Logon'
'         Else
'            SaveSetting "WAPS", "Logon", "Traveller", ",0"  ' <-- NOTE: removal of version from 'Logon'
'         End If
'
'        LogLocal
'
'        '
'        ' Laptop determination
'        ' A desktop PC should not be able to write to D:, as it should be a CD ROM Drive
'        On Error Resume Next
'            f = FreeFile
'            strTemp = "D:\Attempt to create flag file by Logon Application.txt"
'            Err.Clear
'            Open strTemp For Output As #f
'                Print #f, "This file is safe to delete"
'            Close f
'            If Err.Number = 0 Then
'                isLaptop = True
'                UpdateStatus "Laptop detected"
'                ' Remove temporary file
'                Kill strTemp
'            Else
'                isLaptop = False
'                UpdateStatus "Desktop detected"
'            End If
'        On Error GoTo 0
'
'        If siteServer = "" Then siteServer = LogonServer
'
'        If Not isLaptop Then
'            ' This is for desktop users only
'            bReturnVal = DriveExists("H:")
'            If bReturnVal Then ' H: mapping exists.  Remove it.
'                bReturnVal = RemoveDrive("H:", 0)      ' Mode 0 is SUBST
'                If Not bReturnVal Then
'                    bReturnVal = RemoveDrive("H:", 1)  ' Mode 1 is API
'                    If bReturnVal Then
'                        UpdateStatus "Removed H: with API Call"
'                    Else
'                        UpdateStatus "Unable to remove existing H:"
'                    End If
'                Else
'                    UpdateStatus "Removed H: with SUBST H: /D"
'                End If
'            End If
'        End If
'        ' This is for desktop and laptop
'        If FolderExists(siteServer & "\" & Username & "$") Then
'            If isLaptop Then
'                UpdateStatus "Mapping G: to " & siteServer & "\" & Username & "$ (API)"
'                retval = Connect2("G:", siteServer & "\" & Username & "$", Username, "")
'            Else
'                UpdateStatus "Mapping H: to " & siteServer & "\" & Username & "$ (API)"
'                retval = Connect2("H:", siteServer & "\" & Username & "$", Username, "")
'            End If
'        Else
'            If Not isLaptop Then
'                If FolderExists(siteServer & "\mobile$") Then
'                    If FolderExists(siteServer & "\mobile$\" & Username) Then
'                        ' Do nothing (yet)
'                    Else
'                        ' Create Directory
'                        bReturnVal = CreateDir(siteServer & "\mobile$\" & Username)
'                        If bReturnVal Then
'                            ' Apply permissions to directory
'                            UpdateStatus "Applying permissions to " & siteServer & "\mobile$\" & Username
'                            Shell "cmd /c echo y|c:\winnt\system32\cacls.exe " & siteServer & "\mobile$\" & Username & " /C /T /G " & Username & ":C", vbHide
'                        End If
'                    End If
'                    ' create drive mapping using SUBST
'                    UpdateStatus "Mapping H: to " & siteServer & "\mobile$\" & Username & " (SUBST H: " & siteServer & "\mobile$\" & Username & ")"
'                    bReturnVal = SubstPath("H:", siteServer & "\mobile$\" & Username)
'                    If bReturnVal Then
'                        UpdateStatus "SUBST of H: SUCCESSFUL"
'                    Else
'                        UpdateStatus "SUBST of H: UNSUCCESSFUL.. ERROR!"
'                    End If
'
'                    If Dir("N:\apps\MAPIProfile\Mapiprof.bat") <> "" Then ' Execute mapiprof.bat
'                        UpdateStatus "Shelled out to MAPIProf.bat"
'                        Shell "cmd /c N:\apps\MAPIProfile\Mapiprof.bat", vbHide
'                    Else
'                        ' FILE NOT FOUND!
'                        If DriveExists("N:") Then ' N: is mapped
'                            UpdateStatus "MAPIProf.bat not found"
'                        Else ' N: not mapped
'                            UpdateStatus "N: not mapped, unable to execute MAPIProf.bat"
'                        End If
'                    End If
'
'                    ' Create 'descriptive' text file in H: (zero length file)
'                    Open "H:\This is not your normal Home Drive" For Output As #4
'                    Close 4
'
'                    ' Display message explaining H:
'                    msg = ""
'                    msg = msg & "TRAVELLER DETECTED." & vbCrLf
'                    msg = msg & "Connecting to a substitute home drive." & vbCrLf
'                    msg = msg & "" & vbCrLf
'                    msg = msg & "Please note:" & vbCrLf
'                    msg = msg & "This is not your normal home drive but is a" & vbCrLf
'                    msg = msg & "temporary H: drive assigned to mobile users." & vbCrLf
'                    MsgBox msg, vbOKOnly + vbInformation, "Logon Application"
'                    msg = ""
'                    ' This next message is for info only
'                    UpdateStatus "---The current H: may only be disconnected by performing a SUBST H: /D from a command prompt---"
'                Else
'                    ' mobile$ share does not exist on siteserver.
'                    UpdateStatus siteServer & "\mobile$ does not exist.  This is a CRITICAL error."
'                    ErrorNumber = ErrorNumber Or 65536
'                    CriticalError = True
'                End If
'            End If
'        End If
'    End If
'    ' **********************
    TotalTime = (Ticks(5) - StartTick) / 1000
    If TotalTime >= 30 Then
        SlowLogon = 1
        DebugMsg "Slow logon, so create c:\temp\slowlogon.flg"
        ' This is a slow logon
        Open "c:\temp\slowlogon.flg" For Output As #22
            Print #22, "Slow logon flag.  The presence of this file indicates a slow logon (>30 seconds)."
        Close 22
    Else
        If Dir("c:\temp\slowlogon.flg") <> "" Then
            DebugMsg "Delete existing c:\temp\slowlogon.flg"
            Kill "c:\temp\slowlogon.flg"
        End If
    End If
    
    UpdateStatus "Logon Application time: " & TotalTime & " seconds"
    UpdateStatus "Errornumber      : " & ErrorNumber
    UpdateStatus "SlowLogon flag   : " & SlowLogon
    LogLocal
    'Write results to local log file for future copying to remote server - Appending as necessary
    DebugMsg "Write results to local log file"
    
    If siteServer <> "" Then siteServer = Right(siteServer, Len(siteServer) - 2)
    LogLocal
    Open "C:\Temp\" & Computername & ".log" For Append As #99
        Print #99, "1: " & Computername
        Print #99, "2: " & StartDate & " " & StartTime
        Print #99, "3: " & Format(Year(Date), "0000") & "/" & Format(Month(Date), "00") & "/" & Format(Day(Date), "00") & " " & FinishTime
        Print #99, "4: " & Username
        Print #99, "5: " & siteLoc
        Print #99, "6: " & Right(LogonServer, Len(LogonServer) - 2)
        Print #99, "7: " & (Ticks(1) - StartTick) / 1000
        Print #99, "8: " & (Ticks(2) - Ticks(1)) / 1000
        Print #99, "9: " & (Ticks(3) - Ticks(2)) / 1000
        Print #99, "A: " & (Ticks(4) - Ticks(3)) / 1000
        Print #99, "B: " & (Ticks(5) - Ticks(4)) / 1000
        Print #99, "C: " & ErrorNumber
        Print #99, "D: " & SlowLogon
        Print #99, "E: " & TotalTime
        Print #99, "F: " & App.Major & "." & App.Minor & "." & App.Revision
        Print #99, "G: " & siteServer
        Print #99, "H: " & Traveller
        Print #99, "I: " & IPAddress
        Print #99, "J: " & Client
        Print #99, "K: " & isLaptop
        Print #99, "--------------"
    Close 99
    '
    ' Here is where we display the logon error message.  Because of the way we do errors, this
    ' messagebox will only be shown once, but may contain information about multiple errors.
    ' Note that the boolean variable CriticalError determines whether to show all errors or critical only
    '
    DebugMsg "Check internal error numbers to display message to user (if appropriate)"
    If ErrorNumber > 0 Then
        DebugMsg "Internal error number " & ErrorNumber
        msg = "The logon Application encountered one or more errors whilst processing your logon." & vbCrLf & vbCrLf
        If CriticalError = True Then
            msg = msg & "Please contact the helpdesk on 1800 651 327 for assistance in rectifying this problem." & vbCrLf & vbCrLf
            msg = msg & "One or more of these errors are critical and require attention." & vbCrLf & vbCrLf
        Else
            msg = msg & "None of these errors are critical." & vbCrLf & vbCrLf
        End If
        
        ' Resolve error with users logging on locally, and getting error 4096 (Logon >33 mod)
        'If ErrorNumber And 4096 Then
        '    If Right(LogonServer, Len(LogonServer) - 2) = Computername Then
        '        msg = msg & vbCrLf & "Your workstation was not able to contact a Domain Controller in this session." & vbCrLf
        '        msg = msg & vbCrLf & "If you speak to the helpdesk, please inform them that you have 'Logged on with a cached profile'." & vbCrLf & vbCrLf
        '    End If
        'End If
        'msg = msg & "Error " & ErrorNumber & vbCrLf & vbCrLf
        'msg = msg & App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision & "  Mode = " & Version
        
        
        ' Logon34 mod
        ' Is it OK to enable the desktop?
        d = Dir("c:\temp\mapidone.txt")
        While d = ""
            ' wait until it disappears
            DoEvents
            d = Dir("c:\temp\mapidone.txt")
        Wend
        ' Enable desktop
        EnableDesktop True
        
        
        If CriticalOnly = True And DebugMode = False Then
            If CriticalError = True Then
                MsgBox msg, vbCritical, "Logon Application Error(s)"
            Else
                ' Dont display errors
            End If
        Else
            If CriticalError = True Then
                MsgBox msg, vbCritical, "Logon Application Error(s)"
            Else
                MsgBox msg, vbExclamation, "Logon Application Error(s)"
            End If
        End If
    Else
        UpdateStatus "No errors detected"
    End If
    SocketsCleanup
    
    ' Cleanup
    If Dir("c:\temp\mapdrive*.bat") <> "" Then
        Do Until Dir("c:\temp\deleteme*.bat") = ""
            UpdateStatus "Waiting for c:\temp\deletemeX.bat files to be removed"
            DoEvents
        Loop
        UpdateStatus "Removing c:\temp\mapdriveX.bat files"
        Kill "c:\temp\mapdrive*.bat"
    End If
        
    ' Write out to log file for the last time
    LogLocal
    
    ' ******** Logon 32 BUILD 29 Modification ***********
                                                         ' Execute Userstat.exe if present
    If Dir("C:\WinNT\stats\userstat.exe") <> "" Then
      Shell "C:\WinNT\stats\userstat.exe"
    Else                                                 ' Oops, file not present
      If Dir("C:\winnt\stats", vbDirectory) = "" Then    ' Check for Stats directory
        On Error Resume Next ' disregard errors
          MkDir ("C:\WinNT\Stats")                       ' Not present, make directory
          On Error GoTo 0    ' Reset error checking
      End If
                                                         ' Copy file to local machine
      FileCopy LogonServer & "\NETLOGON\SETUP\Userstat.exe", "C:\WinNT\Stats\Userstat.exe"
      Shell "C:\WinNT\stats\userstat.exe"                ' Execute newly copied userstat.exe
    End If
    End                                                  ' End program
End Sub

Function DebugMsg(Message As String)
    ' write out to the log file the debugging info, but only as appropriate.
    If DebugMode = False Then Exit Function
    UpdateStatus "**** " & Message
    LogLocal
End Function

Function UpdateStatus(Message As String)
    ' Update the listbox with current status for future writing to the log file
    lstStatus.AddItem Format(Time, "hh:mm:ss") & " " & Message
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lblStatus.Caption = Message
    frmMain.Refresh
End Function

Sub LogLocal()
    Dim lp As Integer
    'Update the log file on the workstation's c:\temp
    ' I know this is a bit slow (the delete/create option), but to ensure that the log file
    ' is updated during the logon process, we have to do this.
    '
    ' This sub MAY be the cause of VERY intermittent problems.
    ' On Error code placed in here just in case.
    '
    If Dir("c:\temp\" & Username & ".log") <> "" Then
        On Error Resume Next
            Kill "c:\temp\" & Username & ".log"
            If Err.Number <> 0 Then Exit Sub
        On Error GoTo 0
    End If
    On Error Resume Next
    Open "c:\temp\" & Username & ".log" For Output As #5 ' Output log file on local machine
        If Err.Number <> 0 Then Exit Sub
        For lp = 0 To lstStatus.ListCount - 1
            Print #5, lstStatus.List(lp)
        Next lp
    Close 5
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Stop the user from clicking on the windows close button (X) to end the program
    DebugMsg "Started Form_QueryUnload() Sub"
    If UnloadMode = 0 Then Cancel = True
    SocketsCleanup
End Sub

Sub PlayWav()
    ' As the name suggests, this sub plays a .wav file
    DebugMsg "Started Playwav() Sub"
    Dim sBuf As String, cSize As Long, retval As Long, WinDir As String, n As Long
    'Create a variable large enough to store the Windows path.
    sBuf = String(255, 0)
    cSize = 255
    retval = GetWindowsDirectoryA(sBuf, cSize)
    'Strip buffer from Windows directory
    WinDir = Left(sBuf, retval)
    'Load and Play the sound.
    n = sndPlaySound(WinDir + "\Media\The Microsoft Sound.wav", 0)
End Sub


Public Function Connect(sDrive, sService As String) As Long
    ' Connect a network drive routine - only works if already authenticated to target server
    ErrorNum = 0
    ErrorMsg = "SUCCESS"
    On Error GoTo Err_Connect
    rc = WNetAddConnection(sService & Chr(0), "" & Chr(0), sDrive & Chr(0))
    Connect = rc
    If rc <> 0 Then GoTo Err_Connect
    Exit Function
Err_Connect:
    ErrorNum = rc
    ErrorMsg = WnetError(rc)
    Connect = rc
End Function

Public Function Connect2(ByVal Localname As String, ByVal RemoteName As String, ByVal Username As String, ByVal Password As String) As Long
    Dim lpUserName As String
    Dim lpPassword As String

    On Error GoTo Err_Connect
    ErrorNum = 0
    ErrorMsg = "SUCCESS"
    lpNetResource.dwType = RESOURCETYPE_DISK
    lpNetResource.dwScope = RESOURCE_GLOBALNET
    lpNetResource.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    lpNetResource.dwUsage = RESOURCEUSAGE_CONNECTABLE
    lpNetResource.lpLocalName = Localname
    lpNetResource.lpRemoteName = RemoteName
    'lpPassword = Chr(0) ' THESE LINES SCREW THE CODE!!! LEAVE WELL ENOUGH ALONE
    'lpUserName = Chr(0) ' VB Requires these to be NULL, but the API doesn't accept
                         ' NULL, so don't assign values at all.
    rc = WNetAddConnection2(lpNetResource, lpPassword, lpUserName, 0)
    Connect2 = rc
    DebugMsg "Connecting drive returned error " & rc & " (" & WnetError(rc) & ")"
    If rc <> 0 Then GoTo Err_Connect
    Exit Function
Err_Connect:
    ErrorNum = rc
    ErrorMsg = WnetError(rc)
    Connect2 = rc
    DebugMsg "Connecting drive returned error " & rc & " (" & WnetError(rc) & ")"
End Function

Public Function DisConnect(sDrive As String) As Long
    ' Disconnect a network drive routine
    ErrorNum = 0
    ErrorMsg = "SUCCESS"
    On Error GoTo Err_DisConnect
    rc = WNetCancelConnection(sDrive + Chr(0), 0)
    DisConnect = rc
    If rc <> 0 Then GoTo Err_DisConnect
    Exit Function
Err_DisConnect:
    ErrorNum = rc
    ErrorMsg = WnetError(rc)
    DisConnect = rc
End Function

Public Function DisConnect2(ByVal Name As String, ByVal ForceOff As Boolean) As Long
    On Error GoTo Err_DisConnect
    ErrorNum = 0
    ErrorMsg = "SUCCESS"
    rc = WNetCancelConnection2(Name & Chr(0), CONNECT_UPDATE_PROFILE, ForceOff)
    DisConnect2 = rc
    If rc <> 0 Then GoTo Err_DisConnect
    Exit Function
Err_DisConnect:
    ErrorNum = rc
    ErrorMsg = WnetError(rc)
    DisConnect2 = rc
End Function


Private Function WnetError(Errcode As Long) As String
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
        Case Else:
           WnetError = "Unrecognized Error " + Str(Errcode) + "."
      End Select
End Function

Function ResolveHostname(Hostname As String) As String
    Dim strHostName As String * 256
    Dim lngHosetnAddr As Long
    Dim udtHOST As HOSTENT
    Dim lngHostIP As Long
    Dim arrTempIPAddr() As Byte
    Dim i As Integer
    Dim strIPAddress As String
    
    strHostName = Hostname
   
    If Len(strHostName) = 0 Then Exit Function
    
    lngHosetnAddr = gethostbyname(Trim(strHostName))
    If lngHosetnAddr = 0 Then
        UpdateStatus "Error getting IP. Could not resolve " & Hostname
        Exit Function
    End If

    'Exit Function
    
    RtlMoveMemory udtHOST, lngHosetnAddr, LenB(udtHOST)
    RtlMoveMemory lngHostIP, udtHOST.hAddrList, 4

    'get all of the IP addresses if machine is multi-homed
    Do
        ReDim arrTempIPAddr(1 To udtHOST.hLength)
        RtlMoveMemory arrTempIPAddr(1), lngHostIP, udtHOST.hLength

        For i = 1 To udtHOST.hLength
            strIPAddress = strIPAddress & arrTempIPAddr(i) & "."
        Next
        strIPAddress = Mid(strIPAddress, 1, Len(strIPAddress) - 1)
        
        UpdateStatus "Resolved " & Hostname & " to " & strIPAddress
        
        ResolveHostname = strIPAddress

        strIPAddress = ""
        udtHOST.hAddrList = udtHOST.hAddrList + LenB(udtHOST.hAddrList)
        RtlMoveMemory lngHostIP, udtHOST.hAddrList, 4
    Loop While (lngHostIP <> 0)
End Function

Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSADATA
    Dim lngRetVal As Integer
    Dim strLowByte As String
    Dim strHighByte As String
    Dim strMsg As String

    lngRetVal = WSAStartup(WS_VERSION_REQD, WSAD)

    If lngRetVal <> 0 Then
        'MsgBox "Winsock.dll is not responding."
        UpdateStatus "Winsock.dll is not responding."
        SocketsInitialize = False
        'End
    End If

    If LoByte(WSAD.wversion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wversion) = _
        WS_VERSION_MAJOR And HiByte(WSAD.wversion) < WS_VERSION_MINOR) Then

        strHighByte = Trim(Str(HiByte(WSAD.wversion)))
        strLowByte = Trim(Str(LoByte(WSAD.wversion)))
        strMsg = "Windows Sockets version " & strLowByte & "." & strHighByte
        strMsg = strMsg & " is not supported by winsock.dll "
        UpdateStatus strMsg
        SocketsInitialize = False
    End If

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        strMsg = "This application requires a minimum of "
        strMsg = strMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        UpdateStatus strMsg
        SocketsInitialize = False
    End If
    
    SocketsInitialize = True
End Function

Sub SocketsCleanup()
    Dim lngRetVal As Long

    lngRetVal = WSACleanup()

    If lngRetVal <> 0 Then
        'MsgBox "Socket error " & Trim(Str(lngRetVal)) & " occurred in Cleanup "
        UpdateStatus "Socket error " & Trim(Str(lngRetVal)) & " occurred in Cleanup "
        'End
    End If
End Sub

Public Function GetIPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim HOST As HOSTENT
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim i As Integer
    Dim sIPAddr As String
    
    If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        'MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)
    
    If lpHost = 0 Then
        GetIPAddress = ""
        'MsgBox "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    RtlMoveMemory HOST, lpHost, Len(HOST)
    RtlMoveMemory dwIPAddr, HOST.hAddrList, 4
    
    ReDim tmpIPAddr(1 To HOST.hLength)
    
    RtlMoveMemory tmpIPAddr(1), dwIPAddr, HOST.hLength
    
    For i = 1 To HOST.hLength
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
    
End Function

' *********************************
' Logon 33 modification
' *********************************

' NOTE: - Error -2147467259 in this sub is due to Database drivers.
' See MSDN Q209157 for further info
' The following is an extract of this KB article.
'
' This run-time error occurs only while using Automation.
' The occurrence varies based on whether or not Access is idle after the form opens,
' but before the Automation code attempts to set focus to a control. Variations in
' CPU speeds, program timing, and operating systems may cause this behavior to occur
' on some computers, but not others.
'
Function GetSiteVariables(Subnet As String, siteLoc As String, siteServer As String) As Boolean
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim DSN As String
    Dim SQL As String
    Dim dtLocal As Date, dtRemote As Date
    Dim LocalPath1 As String, RemotePath1 As String
    Dim LocalPath2 As String, RemotePath2 As String
    Dim t As Date, E As Long
    
    On Error GoTo Hell
    DebugMsg "Inside GetSiteVariables() Function"
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    
    LocalPath1 = "c:\temp\siteinfo.csv"
    RemotePath1 = LogonServer & "\netlogon\sites\siteinfo.csv"
    LocalPath2 = "c:\temp\schema.ini"
    RemotePath2 = LogonServer & "\netlogon\sites\schema.ini"
    
    DebugMsg "Local Path for siteinfo: " & LocalPath1
    DebugMsg "Remote Path for siteinfo: " & RemotePath1
    
    ' get most recent files
    If Dir(LocalPath1) <> "" Then
        DebugMsg "Found " & LocalPath1
        If Dir(RemotePath1) <> "" Then
            DebugMsg "Found " & RemotePath1
            dtRemote = FileDateTime(RemotePath1)
            dtLocal = FileDateTime(LocalPath1)
            If dtLocal < dtRemote Then
                DebugMsg "Copying " & RemotePath1 & " to " & LocalPath1
                FileCopy RemotePath1, LocalPath1
                FileCopy RemotePath2, LocalPath2
            Else
                DebugMsg LocalPath1 & " is up to date.. not copied locally"
            End If
        Else
            ' No siteinfo.csv on logonserver, so use local copy
            ' modify these next 2 lines to NOT use local file when remote file is absent
            ' GetSiteVariables = False ' <-- will force logon to use pre Logon 33 values
            ' Exit Function
            DebugMsg "Using " & LocalPath1 & ", because " & RemotePath1 & " not found"
        End If
    Else ' No siteinfo.csv on local machine
        DebugMsg LocalPath1 & " not found"
        If Dir(RemotePath1) <> "" Then
            DebugMsg "Copying " & RemotePath1 & " to " & LocalPath1
            FileCopy RemotePath1, LocalPath1
            FileCopy RemotePath2, LocalPath2
        Else
            ' No siteinfo.csv on logonserver!!
            GetSiteVariables = False ' <-- will force logon to use pre Logon 33 values
            DebugMsg LocalPath1 & " and " & RemotePath1 & " not found!  Using Pre Logon 33 values."
            Exit Function
        End If
    End If
    
    ' pause for two seconds to ensure objects are created
    t = Time
    Do Until DateDiff("s", t, Time) >= 2
        'DoEvents ' relinquish control to windows to allow processing of windows messages
    Loop
    
    DSN = "Driver=" & "{Microsoft Text Driver (*.txt; *.csv)}" & ";" & "DBQ=" & "C:\temp" & ";" & "DefaultDir=" & "C:\temp" & ";" & "Uid=Admin;Pwd=;"
    SQL = ""
    SQL = SQL & "SELECT * FROM siteinfo#csv WHERE "
    SQL = SQL & "Subnet = '" & Subnet & "'" ' <-- NOTE:  field names are case sensitive
    
    adoConn.Open DSN
    adoRS.Open SQL, adoConn
    If Not adoRS.EOF Then
        siteLoc = Trim(adoRS("Siteloc") & "")
        siteServer = "\\" & Trim(adoRS("Siteserver") & "")
    End If
    adoRS.Close
    adoConn.Close
    Set adoRS = Nothing
    Set adoConn = Nothing
    If siteLoc <> "" And siteServer <> "\\" Then
        GetSiteVariables = True
    Else
        UpdateStatus "Could not determine siteloc or siteserver using subnet of '" & Subnet & "'"
    End If
    Exit Function
Hell:
    E = Err.Number
    If E = -2147467259 Then ' Automation Error.. see note before this Function
        Resume
    Else
        DebugMsg "Error " & E & " inside GetSiteVariables() Function"
        Debug.Print "Error " & E & " inside GetSiteVariables() Function"
    End If
End Function

Function FolderExists(Foldername As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FolderExists(Foldername)) Then
        FolderExists = True
    Else
        FolderExists = False
    End If
    Set fso = Nothing
End Function

Function DriveExists(Drivename As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.DriveExists(Drivename)) Then
        DriveExists = True
    Else
        DriveExists = False
    End If
    Set fso = Nothing
End Function

Function RemoveDrive(Drivename As String, mode As Integer) As Boolean
    Dim t As Date, retval As Long
    ' mode 0 is SUBST
    ' mode 1 is API
    Select Case mode
        Case 0
            ' SUBST
            Shell "cmd /c SUBST H: /D", vbHide
            ' pause for five seconds to ensure Drive is removed, or drive is removed
            t = Time
            Do Until (DriveExists(Drivename) = False) Or (DateDiff("s", t, Time) >= 5)
                'DoEvents ' relinquish control to windows to allow processing of windows messages
            Loop
            If DriveExists(Drivename) = False Then
                RemoveDrive = True  ' SUCCESS
            Else
                RemoveDrive = False ' FAIL
            End If
        Case 1
            ' API
            retval = DisConnect2(Drivename, True)
            If retval = 0 Then
                RemoveDrive = True ' SUCCESS
            Else
                RemoveDrive = False ' FAIL
            End If
    End Select
End Function

Function CreateDir(Path As String) As Boolean
    Dim t As Date
    MkDir Path
    t = Time
    ' pause for up to five seconds to ensure path is created
    Do Until (DateDiff("s", t, Time) >= 5) Or (FolderExists(Path))
        'DoEvents ' relinquish control to windows to allow processing of windows messages
    Loop
    If FolderExists(Path) Then
        CreateDir = True
    Else
        CreateDir = False
    End If
End Function

Function SubstPath(Drive As String, Path As String) As Boolean
    Dim t As Date, cmd As String
    If Drive = "" Or Path = "" Then Exit Function
    cmd = "cmd /c SUBST " & Drive & " " & Path
    Shell cmd, vbHide
    t = Time
    ' pause for up to five seconds to ensure drive is SUBSTed
    Do Until (DateDiff("s", t, Time) >= 5) Or (DriveExists(Drive))
        'DoEvents ' relinquish control to windows to allow processing of windows messages
    Loop
    If DriveExists(Drive) Then
        SubstPath = True
    Else
        SubstPath = False
    End If
End Function

Function GetSiteVariables2(Subnet As String, siteLoc As String, siteServer As String) As Boolean
    Dim dtLocal As Date, dtRemote As Date
    Dim LocalPath As String, RemotePath As String
    Dim t As Date, E As Long, f As Integer
    Dim tmpSubnet As String, tmpSiteloc As String, tmpSiteserver As String
    Dim bReturnVal As Boolean
    
    DebugMsg "Inside GetSiteVariables2() Function"
    
    LocalPath = "c:\temp\siteinfo.csv"
    RemotePath = LogonServer & "\netlogon\sites\siteinfo.csv"
    bReturnVal = UpdateFile(LocalPath, RemotePath)
    'LocalPath2 = "c:\temp\schema.ini"
    'RemotePath2 = LogonServer & "\netlogon\sites\schema.ini"
    
    'DebugMsg "Local Path for siteinfo: " & LocalPath1
    'DebugMsg "Remote Path for siteinfo: " & RemotePath1
    
    ' get most recent files
    'If Dir(LocalPath1) <> "" Then
    '    DebugMsg "Found " & LocalPath1
    '    If Dir(RemotePath1) <> "" Then
    '        DebugMsg "Found " & RemotePath1
    '        dtRemote = FileDateTime(RemotePath1)
    '        dtLocal = FileDateTime(LocalPath1)
    '        If dtLocal < dtRemote Then
    '            DebugMsg "Copying " & RemotePath1 & " to " & LocalPath1
    '            FileCopy RemotePath1, LocalPath1
    '            'FileCopy RemotePath2, LocalPath2
    '        Else
    '            DebugMsg LocalPath1 & " is up to date.. not copied locally"
    '        End If
    '    Else
    '        ' No siteinfo.csv on logonserver, so use local copy
    '        ' modify these next 2 lines to NOT use local file when remote file is absent
    '        ' GetSiteVariables = False ' <-- will force logon to use pre Logon 33 values
    '        ' Exit Function
    '        DebugMsg "Using " & LocalPath1 & ", because " & RemotePath1 & " not found"
    '    End If
    'Else ' No siteinfo.csv on local machine
    '    DebugMsg LocalPath1 & " not found"
    '    If Dir(RemotePath1) <> "" Then
    '        DebugMsg "Copying " & RemotePath1 & " to " & LocalPath1
    '        FileCopy RemotePath1, LocalPath1
    '        'FileCopy RemotePath2, LocalPath2
    '    Else
    '        ' No siteinfo.csv on logonserver!!
    '        GetSiteVariables2 = False ' <-- will force logon to use pre Logon 33 values
    '        DebugMsg LocalPath1 & " and " & RemotePath1 & " not found!  Using Pre Logon 33 values."
    '        Exit Function
    '    End If
    'End If
    If bReturnVal Then
        
    Else
        GetSiteVariables2 = False
        Exit Function
    End If
    
    f = FreeFile
    If Dir(LocalPath, vbNormal) <> "" Then
        Open LocalPath For Input As #f
            Do Until EOF(f) Or (tmpSubnet = Subnet)
                Input #f, tmpSubnet, tmpSiteloc, tmpSiteserver
            Loop
        Close f
    Else
        UpdateStatus LocalPath & " not found."
        GetSiteVariables2 = False
    End If
    If tmpSubnet = Subnet Then
        siteLoc = tmpSiteloc
        siteServer = "\\" & tmpSiteserver
        GetSiteVariables2 = True
    Else
        GetSiteVariables2 = False
    End If
End Function

' New function for Logon 34+
Function UpdateFile(LocalPath As String, RemotePath As String) As Boolean
    Dim dtLocal As Date, dtRemote As Date
    DebugMsg "Inside UpdateFile() Function"
    If Dir(LocalPath) <> "" Then
        DebugMsg "Found " & LocalPath
        If Dir(RemotePath) <> "" Then
            DebugMsg "Found " & RemotePath
            dtRemote = FileDateTime(RemotePath)
            dtLocal = FileDateTime(LocalPath)
            If dtLocal <> dtRemote Then
                DebugMsg "Copying " & RemotePath & " to " & LocalPath
                FileCopy RemotePath, LocalPath
                UpdateFile = True
            Else
                DebugMsg LocalPath & " is up to date.. not copied locally"
                UpdateFile = True
            End If
        Else
            DebugMsg "Using " & LocalPath & ", because " & RemotePath & " not found"
            UpdateFile = True
        End If
    Else ' No localfile
        DebugMsg LocalPath & " not found"
        If Dir(RemotePath) <> "" Then
            DebugMsg "Copying " & RemotePath & " to " & LocalPath
            FileCopy RemotePath, LocalPath
            UpdateFile = True
        Else
            ' No remotefile OR localfile !!!
            UpdateFile = False
        End If
    End If
End Function

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
        
    RResult = RegConnectRegistry(Hostname, HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then
      GetValue = "No Data"
      Exit Function
    End If
    
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
        
    'RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
    RResult = RegConnectRegistry("", HKEY_CURRENT_USER, OpenHiveVal)
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
    
    'RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
    RResult = RegConnectRegistry(CMPTRName, HKEY_CURRENT_USER, OpenHiveVal)
    keytogo = Trim(Key)
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
    RegistryWriteValue Data, OpenKeyVal, Value, regType
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    'DisIPCConnection (Trim(Hostname))
        
End Function

Sub EnableDesktop(ByVal bEnable As Boolean)
    Dim hWnd_Desktop As Long
    hWnd_Desktop = GetWindow(FindWindow("Progman", "Program Manager"), GW_CHILD)
    If bEnable Then
        EnableWindow hWnd_Desktop, True
    Else
        EnableWindow hWnd_Desktop, False
    End If
End Sub


