VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon Client"
   ClientHeight    =   1890
   ClientLeft      =   210
   ClientTop       =   2700
   ClientWidth     =   3240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1890
   ScaleWidth      =   3240
   Begin VB.TextBox txtSendData 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Send"
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Output"
      Top             =   2760
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   120
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStep 
      Caption         =   "Mapping Printers"
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Image imgStep 
      Height          =   240
      Index           =   6
      Left            =   120
      Picture         =   "frmClient.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblStep 
      Caption         =   "Mapping Drives"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Image imgStep 
      Height          =   240
      Index           =   5
      Left            =   120
      Picture         =   "frmClient.frx":0442
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgStep 
      Height          =   240
      Index           =   4
      Left            =   120
      Picture         =   "frmClient.frx":0884
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblStep 
      Caption         =   "Disconnected"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblStep 
      Caption         =   "Disconnecting from Logon Server"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.Image imgStep 
      Height          =   240
      Index           =   3
      Left            =   120
      Picture         =   "frmClient.frx":0CC6
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblStep 
      Caption         =   "Received Logon Data"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.Image imgStep 
      Height          =   240
      Index           =   2
      Left            =   120
      Picture         =   "frmClient.frx":1108
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblStep 
      Caption         =   "Sending User Logon Request"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image imgStep 
      Height          =   240
      Index           =   1
      Left            =   120
      Picture         =   "frmClient.frx":154A
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblStep 
      Caption         =   "Contacting Logon Server"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image imgStep 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "frmClient.frx":198C
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StepNumber As Integer
Dim Hostname As String

Private Type LOGON_INFO
    DiskFree_C As String
    Domain As String
    Homedir As String
    Homeshare As String
    Hostname As String
    Hotfix() As String
    isLaptop As Boolean
    IPAddress As String
    LogonServer As String
    MACAddress As String
    Memory As String
    OperatingSystem As String
    ServicePack As String
    Username As String
End Type

Dim Logon As LOGON_INFO

Private Sub Form_Load()
    Dim WSHNetwork As IWshRuntimeLibrary.WSHNetwork
    'Dim WSHSystem As IWshRuntimeLibrary.WshEnvironment
    Dim oWMI As Object

    Me.Show
    Me.Refresh
    
    Set WSHNetwork = CreateObject("WScript.Network")
    'Set WSHSystem = CreateObject("WScript.Environment")
    
    On Error Resume Next
    ' Get User and Machine Logon Info
    
    ' Get Battery info
    Set oWMI = CreateObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT Status FROM Win32_PortableBattery")
    For Each battery In oWMI
        Logon.isLaptop = (battery.Status <> "")
    Next
    Set oWMI = Nothing
    
    ' Get MAC Address
    Set oWMI = CreateObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapter WHERE AdapterType = 'Ethernet 802.3'")
    For Each Card In oWMI
        Logon.MACAddress = Card.MACAddress
    Next
    
    ' Get IP Address
    Logon.IPAddress = GetIPAddress
    
    ' Get Username
    Logon.Username = Environ$("username") ' <-- DOES NOT WORK ON WIN9x
    Logon.Username = WSHNetwork.Username
    
    ' Get Hostname
    Logon.Hostname = Environ$("computername") ' <-- DOES NOT WORK ON WIN9x
    Logon.Hostname = WSHNetwork.ComputerName
    
    ' Get Homedir
    Logon.Homedir = Environ$("Homepath") ' <-- DOES NOT WORK ON WIN9x
    
    ' Get HomeShare
    Logon.Homeshare = Environ$("Homeshare") ' <-- DOES NOT WORK ON WIN9x
    
    ' Get Operating System
    Set oWMI = CreateObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT Caption FROM Win32_OperatingSystem WHERE Primary = TRUE")
    For Each os In oWMI
        Logon.OperatingSystem = os.Caption
    Next
    
    ' Get Service Pack
    Set oWMI = CreateObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT CSDVersion FROM Win32_OperatingSystem WHERE Primary = TRUE")
    For Each os In oWMI
        Logon.ServicePack = os.CSDVersion
    Next
    
    ' Get Domain
    Logon.Domain = Environ$("Userdomain") ' <-- DOES NOT WORK ON WIN9x
    Logon.Domain = WSHNetwork.UserDomain
    
    ' Get Memory
    Set oWMI = CreateObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
    For Each Sys In oWMI
        Logon.Memory = Sys.TotalPhysicalMemory
    Next
    
    ' Get Diskspace
    Set oWMI = CreateObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT FreeSpace, Size FROM Win32_LogicalDisk WHERE Caption = 'C:'")
    For Each Sys In oWMI
        Logon.DiskFree_C = Sys.FreeSpace
    Next
    
    ' Get Hotfixes
    Set oWMI = CreateObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT HotfixID FROM Win32_QuickFixEngineering")
    For Each Sys In oWMI
        i = i + 1
        ReDim Preserve Logon.Hotfix(i)
        Logon.Hotfix(i) = Sys.HotfixID
    Next
    
    Set oWMI = Nothing
    
    NextStep ' Sets form to 'Contacting Logon Server'
    
    
    'tcpClient.RemoteHost = Replace(Environ$("logonserver"), "\", "")
    tcpClient.RemoteHost = Replace(Logon.Hostname, "\", "")
    tcpClient.RemoteHost = "LAPTOP"
    tcpClient.RemotePort = 1001
    tcpClient.Connect
End Sub

Private Sub tcpClient_Close()
    StepNumber = 4
    NextStep       ' Sets form to 'Disconnected'
End Sub

Private Sub tcpClient_Connect()
    'Dim Username As String, Hostname As String, IP As String
    Dim TextToSend As String, lp As Integer
    Dim Hotfix As String
    
    NextStep ' Sets form to 'Sending User Logon Request'
    'Username = Environ$("username")
    'Hostname = Environ$("computername")
    'IP = IPAddress
    'TextToSend = username & "+" & Hostname & "+" & IP
    Hotfix = ""
    For lp = 1 To UBound(Logon.Hotfix())
        Hotfix = Hotfix & Logon.Hotfix(lp) & ","
    Next lp
    ' Remove final comma
    Hotfix = Left(Hotfix, Len(Hotfix) - 1)
    TextToSend = "CON:LOGONDATA"
    TextToSend = TextToSend & Logon.DiskFree_C & "+"
    TextToSend = TextToSend & Logon.Domain & "+"
    TextToSend = TextToSend & Logon.Homedir & "+"
    TextToSend = TextToSend & Logon.Homeshare & "+"
    TextToSend = TextToSend & Logon.Hostname & "+"
    TextToSend = TextToSend & Hotfix & "+"
    TextToSend = TextToSend & Logon.IPAddress & "+"
    TextToSend = TextToSend & Logon.isLaptop & "+"
    TextToSend = TextToSend & Logon.LogonServer & "+"
    TextToSend = TextToSend & Logon.MACAddress & "+"
    TextToSend = TextToSend & Logon.Memory & "+"
    TextToSend = TextToSend & Logon.OperatingSystem & "+"
    TextToSend = TextToSend & Logon.ServicePack & "+"
    TextToSend = TextToSend & Logon.Username

    tcpClient.SendData TextToSend
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, strPaths() As String
    tcpClient.GetData strData
    
    If Left(strData, 5) = "PATH:" Then
        strData = Right(strData, Len(strData) - 5) ' strip 'PATH:'
        NextStep ' Sets form to 'Received Logon Data'
        strPaths = Split(strData, "+")
        For lp = 0 To UBound(strPaths())
            Debug.Print strPaths(lp)
        Next lp
        NextStep ' Sets form to 'Disconnecting from Logon Server'
        tcpClient.SendData "***DISCONNECT***" ' Requesting to disconnect
    ElseIf strData = "***DISCONNECT***" Then  ' Disconnect Reply
        NextStep ' Sets form to 'Disconnected'
        tcpClient.Close
    ElseIf strData = "***AUTHORISATION***" Then
        tcpClient.SendData "AUTH:LOGON"
    End If
End Sub

Sub NextStep()
    If StepNumber > 0 Then
        imgStep(StepNumber - 1).Visible = False
    End If
    imgStep(StepNumber).Visible = True
    StepNumber = StepNumber + 1
    frmClient.Refresh
End Sub

Function GetIPAddress()
    Dim objLocator As WbemScripting.SWbemLocator
    Dim objService As WbemScripting.SWbemServices
    Dim objEnumerator As Object
    
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer("", "", "", "")
    strQuery = "Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE"

    Set objEnumerator = objService.ExecQuery(strQuery)

    For Each obj In objEnumerator
        If Not IsNull(obj.IPAddress) Then
            For i = LBound(obj.IPAddress) To UBound(obj.IPAddress)
                GetIPAddress = obj.IPAddress(i)
            Next
        End If
    Next
End Function

