VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make Def."
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   135
      TabIndex        =   1
      Top             =   675
      Width           =   6135
   End
   Begin VB.CommandButton cmdSetDef 
      Caption         =   "Make Default"
      Height          =   510
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   2220
   End
   Begin VB.Label Label1 
      Caption         =   "Select a printer and click on make that printer the default printer for all windows apps."
      Height          =   465
      Left            =   2475
      TabIndex        =   2
      Top             =   135
      Width           =   3795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function PtrCtoVbString(Add As Long) As String

    Dim sTemp As String * 512, x As Long

    x = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Private Sub Win95SetDefaultPrinter()
    Dim Handle As Long          'handle to printer
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim x As Long
    Dim need As Long            ' bytes needed
    Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
    Dim LastError As Long

    ' determine which printer was selected
    PrinterName = List1.List(List1.ListIndex)
    ' none - exit
    If PrinterName = "" Then
        Exit Sub
    End If

    ' set the PRINTER_DEFAULTS members
    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS

    ' Get a handle to the printer
    x = OpenPrinter(PrinterName, Handle, pd)
    ' failed the open
    If x = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' Make an initial call to GetPrinter, requesting Level 5
    ' (PRINTER_INFO_5) information, to determine how many bytes
    ' you need
    x = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    ' don't want to check GetLastError here - it's supposed to fail
    ' with a 122 - ERROR_INSUFFICIENT_BUFFER
    ' redim t as large as you need
    ReDim t((need \ 4)) As Long

    ' and call GetPrinter for keepers this time
    x = GetPrinter(Handle, 5, t(0), need, need)
    ' failed the GetPrinter
    If x = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' set the members of the pi5 structure for use with SetPrinter.
    ' PtrCtoVbString copies the memory pointed at by the two string
    ' pointers contained in the t() array into a Visual Basic string.
    ' The other three elements are just DWORDS (long integers) and
    ' don't require any conversion
    pi5.pPrinterName = PtrCtoVbString(t(0))
    pi5.pPortName = PtrCtoVbString(t(1))
    pi5.Attributes = t(2)
    pi5.DeviceNotSelectedTimeout = t(3)
    pi5.TransmissionRetryTimeout = t(4)

    ' this is the critical flag that makes it the default printer
    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

    ' call SetPrinter to set it
    x = SetPrinter(Handle, 5, pi5, 0)
    ' failed the SetPrinter
    If x = False Then
        MsgBox "SetPrinterFailed. Error code: " & GetLastError()
        Exit Sub
    End If

    ' and close the handle
    ClosePrinter (Handle)

End Sub
Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As String, PrinterPort As String)
    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    'The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

        'Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        'The port name is the second entry after the driver name
        'separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            'Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Private Sub ParseList(lstCtl As Control, ByVal Buffer As String)
    Dim i As Integer

    Dim s As String

    Do
        i = InStr(Buffer, Chr(0))
        If i > 0 Then
            s = Left(Buffer, i - 1)
            If Len(Trim(s)) Then lstCtl.AddItem s
            Buffer = Mid(Buffer, i + 1)
        Else
            If Len(Trim(Buffer)) Then lstCtl.AddItem Buffer
            Buffer = ""
        End If
    Loop While i > 0
End Sub

Private Sub WinNTSetDefaultPrinter()
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim PrinterName As String
    Dim r As Long
    If List1.ListIndex > -1 Then
        'Get the printer information for the currently selected
        'printer in the list. The information is taken from the
        'WIN.INI file.
        Buffer = Space(1024)
        PrinterName = List1.Text
        r = GetProfileString("PrinterPorts", PrinterName, "", _
            Buffer, Len(Buffer))

        'Parse the driver name and port name out of the buffer
        GetDriverAndPort Buffer, DriverName, PrinterPort

        If DriverName <> "" And PrinterPort <> "" Then
            SetDefaultPrinter List1.Text, DriverName, PrinterPort
        End If
    End If
End Sub

Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub


Private Sub cmdSetDef_Click()
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    If osinfo.dwMajorVersion = 3 And osinfo.dwMinorVersion = 51 And _
       osinfo.dwBuildNumber = 1057 And osinfo.dwPlatformId = 2 Then
        Call WinNTSetDefaultPrinter
    ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 _
       And osinfo.dwBuildNumber >= 67109814 And osinfo.dwPlatformId = 1 Then
        Call Win95SetDefaultPrinter
    ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 _
       And osinfo.dwBuildNumber = 1381 And osinfo.dwPlatformId = 2 Then
        Call WinNTSetDefaultPrinter
    End If

End Sub

Private Sub Form_Load()
    Dim r As Long
    Dim Buffer As String

    'Get the list of available printers from WIN.INI
    Buffer = Space(8192)
    r = GetProfileString("PrinterPorts", vbNullString, "", _
       Buffer, Len(Buffer))

    'Display the list of printer in the list box List1
    ParseList List1, Buffer
End Sub

Private Sub List1_Click()

End Sub
