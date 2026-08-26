VERSION 5.00
Begin VB.Form frmPrinters 
   Caption         =   "Printers"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   ControlBox      =   0   'False
   Icon            =   "frmPrinters.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   10680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Print"
      Height          =   375
      Left            =   8760
      TabIndex        =   28
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   11
      Left            =   9360
      Picture         =   "frmPrinters.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   10
      Left            =   7560
      Picture         =   "frmPrinters.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   9
      Left            =   5760
      Picture         =   "frmPrinters.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   8
      Left            =   3960
      Picture         =   "frmPrinters.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   7
      Left            =   2160
      Picture         =   "frmPrinters.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   6
      Left            =   360
      Picture         =   "frmPrinters.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   5
      Left            =   9360
      Picture         =   "frmPrinters.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   4
      Left            =   7560
      Picture         =   "frmPrinters.frx":1850
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   3
      Left            =   5760
      Picture         =   "frmPrinters.frx":1B5A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   2
      Left            =   3960
      Picture         =   "frmPrinters.frx":1E64
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   1
      Left            =   2160
      Picture         =   "frmPrinters.frx":216E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdPrinter 
      Enabled         =   0   'False
      Height          =   855
      Index           =   0
      Left            =   360
      Picture         =   "frmPrinters.frx":2478
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   9000
      TabIndex        =   27
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   7200
      TabIndex        =   25
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Shape shpDefault 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   240
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblDefault 
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   3360
      Width           =   6495
   End
   Begin VB.Label lblDef 
      Caption         =   "Default printer:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10560
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblNetwork 
      Alignment       =   2  'Center
      Caption         =   "Available Printers"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   10455
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   20
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   19
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   18
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   17
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9000
      TabIndex        =   16
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   15
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   13
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblPrinter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Open gPrinterFile For Output As #2
        Print #2, "Printer=" & gDefaultPrinter
    Close #2
    frmStart.Show
    Unload Me
End Sub

Public Sub cmdPrinter_Click(Index As Integer)
    Dim def As Printer
    Dim a As Integer
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    shpDefault.Left = cmdPrinter(Index).Left - 120
    shpDefault.Top = cmdPrinter(Index).Top - 120
    shpDefault.Visible = True
    lblDefault = lblPrinter(Index).Caption
    gDefaultPrinter = lblDefault
    If Printer.DeviceName = lblDefault Then Exit Sub ' required to fix printer errors

    For Each def In Printers
        If a = Index Then Set Printer = def
        a = a + 1
    Next
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    If osinfo.dwMajorVersion = 3 And osinfo.dwMinorVersion = 51 And osinfo.dwBuildNumber = 1057 And osinfo.dwPlatformId = 2 Then
        Call WinNTSetDefaultPrinter
    ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 And osinfo.dwBuildNumber = 67109814 And osinfo.dwPlatformId = 1 Then
        Call Win95SetDefaultPrinter
    ElseIf osinfo.dwMajorVersion = 4 And osinfo.dwMinorVersion = 0 And osinfo.dwBuildNumber = 1381 And osinfo.dwPlatformId = 2 Then
        Call WinNTSetDefaultPrinter
    End If
End Sub

Private Sub cmdTest_Click()
    PrintForm
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    Dim prt As Printer
    def = Printer.DeviceName
    For Each prt In Printers
        lblPrinter(a).Caption = prt.DeviceName
        cmdPrinter(a).ToolTipText = lblPrinter(a).Caption
        If lblPrinter(a) = def Then
            shpDefault.Left = cmdPrinter(a).Left - 120
            shpDefault.Top = cmdPrinter(a).Top - 120
            shpDefault.Visible = True
            lblDefault = lblPrinter(a).Caption
        End If
        a = a + 1
        cmdPrinter(a - 1).Enabled = True
    Next
    Dim r As Long
    Dim Buffer As String

    'Get the list of available printers from WIN.INI
    Buffer = Space(8192)
    r = GetProfileString("PrinterPorts", vbNullString, "", Buffer, Len(Buffer))

    'get the user selected printer from H:\printer.txt
    present = Dir(gPrinterFile)
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
    If gDefaultPrinter = "" Then gDefaultPrinter = Printer.DeviceName
End Sub

Public Function PtrCtoVbString(Add As Long) As String

     Dim sTemp As String * 512, X As Long

     X = lstrcpy(sTemp, Add)
     If (InStr(1, sTemp, Chr(0)) = 0) Then
          PtrCtoVbString = ""
     Else
          PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
     End If
End Function

Public Sub SetDefaultPrinter(ByVal PrinterName As String, ByVal DriverName As String, ByVal PrinterPort As String)
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
 
Public Sub Win95SetDefaultPrinter()
    Dim Handle As Long          'handle to printer
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim X As Long
    Dim need As Long            ' bytes needed
    Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
    Dim LastError As Long

    ' determine which printer was selected
    PrinterName = lblDefault.Caption
    ' none - exit
    If PrinterName = "" Then
        Exit Sub
    End If

    ' set the PRINTER_DEFAULTS members
    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS

    ' Get a handle to the printer
    X = OpenPrinter(PrinterName, Handle, pd)
    ' failed the open
    If X = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' Make an initial call to GetPrinter, requesting Level 5
    ' (PRINTER_INFO_5) information, to determine how many bytes
    ' you need
    X = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    ' don't want to check GetLastError here - it's supposed to fail
    ' with a 122 - ERROR_INSUFFICIENT_BUFFER
    ' redim t as large as you need
    ReDim t((need \ 4)) As Long

    ' and call GetPrinter for keepers this time
    X = GetPrinter(Handle, 5, t(0), need, need)
    ' failed the GetPrinter
    If X = False Then
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
    X = SetPrinter(Handle, 5, pi5, 0)
    ' failed the SetPrinter
    If X = False Then
        MsgBox "SetPrinterFailed. Error code: " & GetLastError()
        Exit Sub
    End If

    ' and close the handle
    ClosePrinter (Handle)

End Sub

Public Sub GetDriverAndPort(ByVal Buffer As String, DriverName As String, PrinterPort As String)
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

Public Sub WinNTSetDefaultPrinter()
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim PrinterName As String
    Dim r As Long
    'Get the printer information for the currently selected
    'printer in the list. The information is taken from the
    'WIN.INI file.
    Buffer = Space(1024)
    PrinterName = lblDefault.Caption
    r = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))

    'Parse the driver name and port name out of the buffer
    frmPrinters.GetDriverAndPort Buffer, DriverName, PrinterPort

    If DriverName <> "" And PrinterPort <> "" Then
        frmPrinters.SetDefaultPrinter lblDefault.Caption, DriverName, PrinterPort
    End If
End Sub
 
