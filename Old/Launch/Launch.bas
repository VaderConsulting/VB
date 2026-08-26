Attribute VB_Name = "modMain"
    Global gProgram(104) As String
    Global gFile(104) As String
    Global gDir(104) As String
    Global gIcon(104) As String
    Global gTip(104) As String
    Global gFileLocation As String
    Global gPersonal As String
    Global gPrevious(104) As Long
    Global gSurname As String
    Global gRank As String
    Global gComputerName As String
    Global gUserName As String
    Global gCriteria(12, 9) As String
    Global gDay(6) As String
    Global gPrinterFile As String
    Global gDefaultPrinter As String
    Global gSignalsDir As String
    
    Rem the following are temporary for use in the frmEdit portion
    Global tmpProgram(26) As String
    Global tmpFile(26) As String
    Global tmpDirs(26) As String
    Global tmpIcon(26) As String
    Global tmpTip(26) As String
    Global gError As Integer
    Global gTZFile As String
    Global gTZ As Long
    Global gDrive1 As String, gDrive2 As String
    
    Public Type LARGE_INTEGER
                lowpart As Long
                highpart As Long
    End Type
    
    'These are the two API functions we'll need to use here.  The first
    'is the one that really does the work.  The second function is used
    'to take action when the user clicks on the mouse icon (restores the
    'program and brings it to front of all other windows.)
    Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
              (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Public Declare Function SetForegroundWindow Lib "user32" _
              (ByVal hwnd As Long) As Long
    
    
    'UDT required by Shell_NotifyIcon API call
    Type NOTIFYICONDATA
        cbSize As Long             'size of this UDT
        hwnd As Long               'handle of the app
        uId As Long                'unused (set to vbNull)
        uFlags As Long             'Flags needed for actions
        uCallBackMessage As Long   'WM we are going to subclass
        hIcon As Long              'Icon we're going to use for the systray
        szTip As String * 64       'ToolTip for the mouse_over of the icon.
    End Type
    
    
    'Constants required by Shell_NotifyIcon API call:
    Private Const NIM_ADD = &H0             'Flag : "ALL NEW nid"
    Private Const NIM_MODIFY = &H1          'Flag : "ONLY MODIFYING nid"
    Private Const NIM_DELETE = &H2          'Flag : "DELETE THE CURRENT nid"
    Private Const NIF_MESSAGE = &H1         'Flag : "Message in nid is valid"
    Private Const NIF_ICON = &H2            'Flag : "Icon in nid is valid"
    Private Const NIF_TIP = &H4             'Flag : "Tip in nid is valid"
    Private Const WM_MOUSEMOVE = &H200      'This is our CallBack Message
    Private Const WM_LBUTTONDOWN = &H201    'LButton down
    Private Const WM_LBUTTONUP = &H202      'LButton up
    Private Const WM_LBUTTONDBLCLK = &H203  'LDouble-click
    Private Const WM_RBUTTONDOWN = &H204    'RButton down
    Private Const WM_RBUTTONUP = &H205      'RButton up
    Private Const WM_RBUTTONDBLCLK = &H206  'RDouble-click
    
    Public nid As NOTIFYICONDATA            'global UDT for the systray function

    Public Const HWND_BROADCAST = &HFFFF
    Public Const WM_WININICHANGE = &H1A
 
    Public Type OSVERSIONINFO
       dwOSVersionInfoSize As Long
       dwMajorVersion As Long
       dwMinorVersion As Long
       dwBuildNumber As Long
       dwPlatformId As Long
       szCSDVersion As String * 128
    End Type
 
    ' constants for DEVMODE structure
    Public Const CCHDEVICENAME = 32
    Public Const CCHFORMNAME = 32

    ' constants for DesiredAccess member of PRINTER_DEFAULTS
    Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
    Public Const PRINTER_ACCESS_ADMINISTER = &H4
    Public Const PRINTER_ACCESS_USE = &H8
    Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

    ' constant that goes into PRINTER_INFO_5 Attributes member to set it as default
    Public Const PRINTER_ATTRIBUTE_DEFAULT = 4

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
            dmICMMethod As Long        ' Windows 95 only
            dmICMIntent As Long        ' Windows 95 only
            dmMediaType As Long        ' Windows 95 only
            dmDitherType As Long       ' Windows 95 only
            dmReserved1 As Long        ' Windows 95 only
            dmReserved2 As Long        ' Windows 95 only
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
    
    ' the following is used in the caps lock code
    Public Const VK_CAPITAL = &H14
    Public Type KeyboardBytes
         kbByte(0 To 255) As Byte
    End Type
    Public kbArray As KeyboardBytes
    
    Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
    
    Public Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
    Public Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
    Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
 
    Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
 
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As String) As Long
    Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
    Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
    Public Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
    Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
    Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
    Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
    Public Declare Function GetLastError Lib "kernel32" () As Long
    Public Const MAX_COMPUTERNAME_LENGTH = 15
    Public Const spacer = "Not used"

    Const REG_DWORD = 4 ' 32-bit number
    Const HKEY_DYN_DATA = &H80000006
    '

Sub Main()
    Rem initialize variables
    Dim Department As String, num As String
    Dim Billet As String
    Dim lp As Integer
    Dim present As String
    Dim response As Integer
    Dim s$, cnt&, dl&, sz&
    
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim PrinterName As String
    Dim r As Long
    
    gDrive1 = "E:" ' H
    gDrive2 = "E:" ' J
    gFileLocation = gDrive2 & "\Bulletin\LAN\"
    gPersonal = gDrive1 & "\Personal.txt"
    gPrinterFile = gDrive1 & "\Printer.txt"
    gTZFile = gFileLocation & "Launcher.tz"
    gSignalsDir = gDrive2 & "\Bulletin\Signals"
    
    gDay(0) = "Sunday"
    gDay(1) = "Monday"
    gDay(2) = "Tuesday"
    gDay(3) = "Wednesday"
    gDay(4) = "Thursday"
    gDay(5) = "Friday"
    gDay(6) = "Saturday"
    
    Printer.TrackDefault = True
    
    'Get the user selected printer from \printer.txt
    On Error Resume Next
    present = Dir(gPrinterFile)
    On Error GoTo 0
    If present = "" Then
        Open gPrinterFile For Output As #2
        Close 2
    End If
    Open gPrinterFile For Input As #2
        Do Until EOF(2)
            Input #2, stuff
            If Left$(stuff, 8) = "Printer=" Then
                gDefaultPrinter = Right$(stuff, Len(stuff) - 8)
            End If
        Loop
    Close 2
    If Printer.DeviceName = gDefaultPrinter Then GoTo special ' required to fix printer errors
    If gDefaultPrinter = "" Then gDefaultPrinter = Printer.DeviceName
    
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    'set default printer according to file located on H:
    If osinfo.dwMajorVersion = 3 And osinfo.dwMinorVersion = 51 And osinfo.dwBuildNumber = 1057 And osinfo.dwPlatformId = 2 Then
        'Get the printer information for the currently selected
        'printer in the list. The information is taken from the
        'WIN.INI file.
        Buffer = Space(1024)
        PrinterName = gDefaultPrinter
        r = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))
    
        'Parse the driver name and port name out of the buffer
        frmPrinters.GetDriverAndPort Buffer, DriverName, PrinterPort
    
        If DriverName <> "" And PrinterPort <> "" Then
            frmPrinters.SetDefaultPrinter gDefaultPrinter, DriverName, PrinterPort
        End If
    ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 And osinfo.dwBuildNumber = 67109814 And osinfo.dwPlatformId = 1 Then
        Call frmPrinters.Win95SetDefaultPrinter
    ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 And osinfo.dwBuildNumber = 1381 And osinfo.dwPlatformId = 2 Then
        'Get the printer information for the currently selected
        'printer in the list. The information is taken from the
        'WIN.INI file.
        Buffer = Space(1024)
        PrinterName = gDefaultPrinter
        r = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))
    
        'Parse the driver name and port name out of the buffer
        frmPrinters.GetDriverAndPort Buffer, DriverName, PrinterPort
    
        If DriverName <> "" And PrinterPort <> "" Then
            frmPrinters.SetDefaultPrinter gDefaultPrinter, DriverName, PrinterPort
        End If
    End If
special:
    Rem check if run in special mode(s)
    If UCase$(Command$) = "VERSION" Then
        MsgBox "Version " & App.Major & "." & App.Minor & App.Revision
        End
    End If
    If UCase$(Command$) = "HELP" Then
        msg = "Version " & App.Major & "." & App.Minor & App.Revision & vbCrLf
        msg = msg & "Usage:        Launcher.exe [option]" & vbCrLf
        msg = msg & "Where [option] is one of the following:" & vbCrLf
        msg = msg & "VERSION" & vbCrLf
        msg = msg & "DEBUG" & vbCrLf
        msg = msg & "HELP" & vbCrLf
        msg = msg & vbCrLf
        MsgBox msg
        End
    End If
    Rem disregrard file errors
    On Error Resume Next
    Rem get computer name using API
    s$ = String$(MAX_COMPUTERNAME_LENGTH + 1, 0)
    sz& = MAX_COMPUTERNAME_LENGTH
    dl& = GetComputerName(s$, sz)
    gComputerName = s$
    
    Rem get username using API
    cnt& = 199
    s$ = String$(200, 0)
    dl& = GetUserName(s$, cnt)
    gUserName = Left$(s$, cnt)
    
    Rem remove last character in user name (chr$(0))
    gUserName = Left$(gUserName, Len(gUserName) - 1)
    Rem work out department
    If Len(gUserName) = 3 Then
        Billet = UCase$(Right$(gUserName, 1))
        If Billet = "A" Then Department = "Operations"
        If Billet = "B" Then Department = "Operations"
        If Billet = "C" Then Department = "Marine Engineering"
        If Billet = "D" Then Department = "Weapons Electrical"
        If Billet = "E" Then Department = "Support"
        If Billet = "F" Then Department = "Medical"
        If Billet = "G" Then Department = "Aviation"
    Else
        Department = UCase$(Left$(gUserName, 1)) & Right$(LCase$(gUserName), Len(gUserName) - 1)
    End If
    Rem do the following file stuff if this isnt in debug mode
    If UCase$(Command$) <> "DEBUG" Then
        For lp = 1 To 4
            Rem firstly load the stuff common to all users
            If lp = 1 Then file = gFileLocation & "Common1.TXT"
            Rem then load the Department specific stuff
            If lp = 2 Then file = gFileLocation & "Common2.TXT"
            Rem then load the Department specific stuff
            If lp = 3 Then file = gFileLocation & Department & ".TXT"
            Rem and lastly, load the user specific stuff
            If lp = 4 Then file = gPersonal
            present = Dir(file)
            If present = "" And lp <> 4 Then
                msg = "List of programs not found.  Launcher will now exit." & vbCrLf & vbCrLf
                msg = msg & "This may be because of either:" & vbCrLf
                msg = msg & "1.  The network connection cannot be found, or" & vbCrLf
                msg = msg & "2.  One of the program lists is unavailable." & vbCrLf & vbCrLf
                msg = msg & "     Check the availability of the following files:" & vbCrLf
                msg = msg & "      " & gFileLocation & Department & ".txt" & vbCrLf
                msg = msg & "      " & gFileLocation & "Common.txt" & vbCrLf
                msg = msg & "      " & gPersonal
                response = MsgBox(msg, vbOKOnly + vbCritical)
                End
            End If
            If present = "" And lp = 4 Then
                Rem there is no personal.txt:  Create personal.txt
                Open gPersonal For Output As #1
                    Rem Write #1, vbCrLf     ' required ??
                Close 1
            End If
            Rem offset refers to the icon number to refer to.
            If lp = 1 Then Offset = 0
            If lp = 2 Then Offset = 26
            If lp = 3 Then Offset = 52
            If lp = 4 Then Offset = 78
            If present <> "" Then
                Open file For Input As #1
                Do Until EOF(1)
                    Line Input #1, num
                    Line Input #1, gProgram(Val(num + Offset) - 1)
                    Line Input #1, gFile(Val(num + Offset) - 1)
                    Line Input #1, gDir(Val(num + Offset) - 1)
                    Line Input #1, gIcon(Val(num + Offset) - 1)
                    Line Input #1, gTip(Val(num + Offset) - 1)
                    frmStart.cmdButton(Val(num + Offset - 1)).Picture = LoadPicture(gIcon(Val(num + Offset) - 1))
                    frmStart.cmdButton(Val(num + Offset - 1)).ToolTipText = gTip(Val(num + Offset) - 1)
                Loop
                Close 1
            End If
            Close 1
        Next lp
        Else
            Rem set up the debug stuff
            num = InputBox("Enter a dummy icon number to use (0-9)", "Debug mode", "0")
            gProgram(Val(num)) = "D:\WINNT\Explorer.exe"
            gFile(Val(num)) = ""
            gDir(Val(num)) = "D:\WINNT"
            gIcon(Val(num)) = "C:\Program Files\Visual Basic\Graphics\Icons\Computer\Msgbox02.ico"
            gTip(Val(num)) = "Windows Explorer"
            frmStart.cmdButton(Val(num)).Picture = LoadPicture(gIcon(Val(num)))
            frmStart.cmdButton(Val(num)).ToolTipText = gTip(Val(num))
        End If
    frmStart.Show
    frmStart.tabLaunch.TabCaption(2) = Department
End Sub

Sub OnTop()
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    If SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS) = True Then
        success% = SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub
