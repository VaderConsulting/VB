VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alerter Master"
   ClientHeight    =   4890
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "DistroTaskAgent"
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtToSend 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   4560
      Width           =   4095
   End
   Begin VB.CommandButton cmdGeneral 
      Caption         =   "Send Text"
      Default         =   -1  'True
      Height          =   375
      Index           =   5
      Left            =   4200
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdGeneral 
      Caption         =   "Close Agent"
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.ListBox lstHistory 
      Height          =   2010
      ItemData        =   "frmMain.frx":0442
      Left            =   0
      List            =   "frmMain.frx":0444
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.ListBox lstAgents 
      Height          =   2010
      ItemData        =   "frmMain.frx":0446
      Left            =   0
      List            =   "frmMain.frx":0448
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin MSWinsockLib.Winsock tcpMain 
      Index           =   0
      Left            =   4200
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   4920
      Top             =   3840
   End
   Begin VB.Label lblHistory 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4320
      Width           =   4095
   End
   Begin VB.Label lblAgents 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Agents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuQuitGrace 
         Caption         =   "&Quit Gracefully"
      End
      Begin VB.Menu N1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitNow 
         Caption         =   "Quit &Now"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGeneral_Click(Index As Integer)
  Dim p As Integer, AgentName As String, TCPConnID As Integer
  Dim oTCP As clsTCP, StringToSend As String
  Dim ListboxContent As String
  If lstAgents.ListCount = 0 Or lstAgents.ListIndex = -1 Then Exit Sub
  p = InStr(1, lstAgents.List(lstAgents.ListIndex), "(")
  
  ListboxContent = lstAgents.List(lstAgents.ListIndex)
  If p > 0 Then
    AgentName = Trim(Mid(ListboxContent, p + 1, Len(ListboxContent) - (p + 1)))
    TCPConnID = FindTCPConnIDfromAgentName(AgentName)
  End If
  
  Select Case Index
    Case 0 ' Shutdown Client
      StringToSend = "S105:"
    Case 5 ' Specific Text
      StringToSend = txtToSend.Text
      txtToSend.Text = ""
  End Select
  
  Set oTCP = TCPConns.Item(TCPConnID)
  oTCP.SendString StringToSend & vbCrLf
  Set oTCP = Nothing

End Sub

Private Sub Form_Load()
  Dim DoStart As Boolean
  
  Randomize Timer
    
  'Defaults
  DoStart = False
  LogSeverity = 0
  DebugMode = False ' ********** MAKE FALSE OUTSIDE DEVELOPMENT ENV.
  DoTCPEncrypt = True
  
  If DebugMode Then
    frmMain.Height = 5580
    frmMain.Width = 5715
  Else
    frmMain.Height = 5220
    frmMain.Width = 4245
  End If
  
  WriteToLog 0, "---------------------------"
  WriteToLog 0, App.EXEName & " Log Opened"
   
  'Parse Command Line
  If (InStr(UCase(Command$), "/LOG:1")) Then LogSeverity = 1
  If (InStr(UCase(Command$), "/LOG:2")) Then LogSeverity = 2
  If (InStr(UCase(Command$), "/LOG:3")) Then LogSeverity = 3
  If (InStr(UCase(Command$), "/LOG:4")) Then LogSeverity = 4
  If (InStr(UCase(Command$), "/INSECURE")) Then DoTCPEncrypt = False
  
  If Not InitObjects Then End
  tmrMain.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  WriteToLog 0, App.EXEName & " Log Closed"
  CloseObjects
End Sub

Private Sub mnuQuitNow_Click()
  Unload Me
End Sub

Private Sub tcpMain_Close(Index As Integer)
  Dim TCPID As Integer
  Dim RIP As String
  Dim RPort As String
  
  TCPID = FindTCPConnFromID(Index)
  If (TCPID <> -1) Then TCPConns.Remove TCPID
  RIP = tcpMain(Index).RemoteHostIP
  RPort = tcpMain(Index).RemotePort
  Unload tcpMain(Index)
  
  DoEvents
  
  WriteToLog 1, "TCP Connection Closed (" & RIP & ":" & RPort & ")"
  AddToHistory "Connection Closed (" & RIP & ":" & RPort & ")"
  RefreshAgentsList
End Sub

Private Function GetNextTCPInstance() As Integer
  Dim i As Integer
  Dim isFound As Boolean
  Dim oTCP As clsTCP
  
  i = 0
  Do
    i = i + 1
    isFound = False
    For Each oTCP In TCPConns
      isFound = isFound Or (oTCP.ID = i)
    Next
  Loop Until Not isFound
  GetNextTCPInstance = i
End Function


Private Sub tcpMain_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  On Error GoTo ConnectError
  Dim oTCP As clsTCP
  Dim newInstanceIndex As Integer
  
  ' Load a new instance of the control to service the connection.
  ' The variable newInstanceIndex maintains the current index of the
  ' next connection to load - this way we don't accidently use an
  ' index already created. Remember to Unload the control when the
  ' connection is closed
  newInstanceIndex = GetNextTCPInstance
  
  Load tcpMain(newInstanceIndex)
  ' Pass the value of the requestID parameter to the
  ' Accept method.
  tcpMain(newInstanceIndex).Accept requestID
  
  WriteToLog 1, "TCP Connection Opened (" & tcpMain(Index).RemoteHostIP & ":" & tcpMain(Index).RemotePort & ")"
  
  ' THESE LINES MUST BE LAST!!!!!
  Set oTCP = New clsTCP
  oTCP.SetupTCPInfo newInstanceIndex
  
  TCPConns.Add oTCP
  
  If (TCPConns.Count > MaxTCPConns) Then ' Too Many Agents Connected
    WriteToLog 0, "TCP Connection " & newInstanceIndex & " will force > maxTCPConns (" & MaxTCPConns & ") - disconnecting client " & newInstanceIndex
    oTCP.SendString "S1FF:Too Many Agents (" & MaxTCPConns & ")" & vbCrLf
'************************************
'* Need to address connection management i.e. close this connection
'************************************
  End If
  
  Exit Sub
ConnectError:
  WriteToLog 4, "TCP Connection Error (" & Err.Description & ")"
  Err.Clear
End Sub

Private Sub tcpMain_DataArrival(Index As Integer, ByVal BytesTotal As Long)
  Dim TCPID As Integer
  
  TCPID = FindTCPConnFromID(Index)
  If (TCPID <> -1) Then TCPConns.Item(TCPID).DataArrival BytesTotal
End Sub

Private Sub tcpMain_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  ' *******************************************************************************
  ' VERY POSSIBLE TO BE STUCK HERE, LOOPING THROUGH CONTINUOUS ERRORS!!!!!!!!!!!!!!
  ' *******************************************************************************
  WriteToLog 0, "TCP Error (" & Number & ") " & Description
  If (Number = 10053) Then ' TCP Timeout Error
    tcpMain.Item(Index).Close
  End If
  CancelDisplay = True
End Sub

Public Sub RefreshAgentsList()
  Dim oTCP As clsTCP
  lstAgents.Clear
  For Each oTCP In TCPConns
    If oTCP.isConfigured Then lstAgents.AddItem oTCP.AgentUsername & " (" & oTCP.AgentHostname & ")"
  Next
End Sub

Private Sub tmrMain_Timer()
  Dim SQL As String, strBroadcast
  Dim CSCDB1 As ADODB.Recordset, temp As Integer
  Set CSCDB1 = New ADODB.Recordset
  SQL = "SELECT * FROM tblAlerts WHERE Complete = 0"
  If DebugMode Then
    strBroadcast = "S1FC:" & Date & " " & Time & Chr(7) & "System" & Chr(7) & Int((Rnd(1) * 3)) + 1 & Chr(7)
    temp = Int(Rnd(1) * 4) + 1
    Select Case temp
      Case 1
        strBroadcast = strBroadcast & "NTSS"
      Case 2
        strBroadcast = strBroadcast & "EUC"
      Case 3
        strBroadcast = strBroadcast & "WAN"
      Case 4
        strBroadcast = strBroadcast & "Helpdesk"
    End Select
    strBroadcast = strBroadcast & Chr(7) & "Alert" & Chr(7) & "Hint"
    If BroadcastTCPMessage(strBroadcast) = True Then
      ' successfully broadcast to one or more clients
      
    End If
    'Debug.Print strBroadcast
  Else
    OpenDatabase
    CSCDB1.Open SQL, DBConn
      Do Until CSCDB1.EOF
        strBroadcast = "S1FC:" & CSCDB1("DateTime") & Chr(7) & CSCDB1("System") & Chr(7) & CSCDB1("Severity") & Chr(7) & CSCDB1("Group") & Chr(7) & CSCDB1("Alert") & Chr(7) & CSCDB1("Hint") & Chr(7)
        If BroadcastTCPMessage(strBroadcast) = True Then
          ' have successfully broadcast to one or more clients
          'Debug.Print "Successfully broadcast " & CSCDB1("ID")
          SQL = "UPDATE tblALerts SET Complete=1 WHERE ID = " & CSCDB1("ID")
          DBConn.Execute SQL
        End If
        'Debug.Print strBroadcast
        CSCDB1.MoveNext
      Loop
    'CSCDB1.Close
    CloseDatabase
  End If
End Sub
