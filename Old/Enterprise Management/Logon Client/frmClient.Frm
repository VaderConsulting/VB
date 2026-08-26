VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon Client"
   ClientHeight    =   1920
   ClientLeft      =   210
   ClientTop       =   2700
   ClientWidth     =   3450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1920
   ScaleWidth      =   3450
   Begin VB.TextBox txtSendData 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "Send"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "Output"
      Top             =   2280
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   3360
      Top             =   120
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

Private Sub Form_Load()
    Me.Show
    Me.Refresh
    NextStep ' Sets form to 'Contacting Logon Server'
    tcpClient.RemoteHost = Replace(Environ$("logonserver"), "\", "")
    'tcpClient.RemoteHost = Replace(Environ$("computername"), "\", "")
    tcpClient.RemotePort = 1001
    tcpClient.Connect
End Sub

Private Sub tcpClient_Close()
    StepNumber = 4
    NextStep       ' Sets form to 'Disconnected'
End Sub

Private Sub tcpClient_Connect()
    Dim Username As String, Hostname As String, IP As String
    Dim TextToSend As String
    NextStep ' Sets form to 'Sending User Logon Request'
    Username = Environ$("username")
    Hostname = Environ$("computername")
    IP = IPAddress
    TextToSend = Username & "+" & Hostname & "+" & IP
    tcpClient.SendData TextToSend
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String, strPaths() As String
    tcpClient.GetData strData
    
    If Mid(strData, 2, 1) = ":" Then
        NextStep ' Sets form to 'Received Logon Data'
        strPaths = Split(strData, "+")
        For lp = 0 To UBound(strPaths())
            Debug.Print strPaths(lp)
        Next lp
        NextStep ' Sets form to 'Disconnecting from Logon Server'
        tcpClient.SendData "***REQ:DIS***" ' Requesting to disconnect
    ElseIf strData = "***REP:DIS***" Then  ' Disconnect Reply
        NextStep ' Sets form to 'Disconnected'
        tcpClient.Close
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

