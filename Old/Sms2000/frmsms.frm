VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Begin VB.Form frmSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMS2000"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11760
   Icon            =   "frmSMS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   11760
   StartUpPosition =   1  'CenterOwner
   Begin MSCommLib.MSComm comSMS 
      Left            =   120
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
      BaudRate        =   2400
      InputMode       =   1
   End
   Begin VB.Timer tmrCom 
      Interval        =   250
      Left            =   840
      Top             =   4560
   End
   Begin VB.Timer tmrHangup 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1440
      Top             =   4560
   End
   Begin VB.Frame fmeNumber 
      Caption         =   "Recipients Mobile Phone number"
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtMobileNumber 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblInternational 
         Caption         =   "International format.  ie 0419 111 222 becomes 61419111222"
         Height          =   615
         Left            =   1560
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fmeOptions 
      Caption         =   "Service Options"
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   3975
      Begin VB.OptionButton optNumber 
         Caption         =   "Telstra Development"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optNumber 
         Caption         =   "Telstra Basic Access"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox txtNumber 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Text            =   "0418707767"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblPhoneNumber 
         Alignment       =   2  'Center
         Caption         =   "Phone Number"
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fmeModem 
      Caption         =   "Modem Commands"
      Height          =   4215
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   7335
      Begin VB.ListBox lstTx 
         Height          =   2985
         ItemData        =   "frmSMS.frx":0442
         Left            =   4800
         List            =   "frmSMS.frx":0444
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   960
         Width           =   2295
      End
      Begin VB.ListBox lstRx 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         ItemData        =   "frmSMS.frx":0446
         Left            =   120
         List            =   "frmSMS.frx":0448
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSendString 
         Caption         =   "Send string"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdInit 
         Caption         =   "Init Modem"
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblOutgoing 
         Alignment       =   2  'Center
         Caption         =   "Outgoing Data"
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblIncoming 
         Alignment       =   2  'Center
         Caption         =   "Incoming Data"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   4575
      End
   End
   Begin VB.Frame fmeMessage 
      Caption         =   "Message"
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3975
      Begin VB.TextBox txtMessage 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.CommandButton cmdSendMsg 
         Caption         =   "Send Message"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblLen2 
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblLen1 
         Caption         =   "Length:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   615
      End
   End
   Begin AgentObjectsCtl.Agent agtSMS 
      Left            =   2040
      Top             =   4560
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   12360
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdInit_Click()
    InitModem
End Sub

Private Sub cmdSendMsg_Click()
    On Error Resume Next
    agtSMS.Characters("Merlin").Show
    agtSMS.Characters("Merlin").Speak "SMS Message Sent"
    txtData = "ATDT" & txtNumber
    cmdSendString_Click
    On Error GoTo 0
    'txtMessage = ""
End Sub

Private Sub cmdSendString_Click()
    'comSMS.PortOpen = True
    'If txtData > "" Then
    comSMS.Output = txtData & vbCrLf
    lstTx.AddItem txtData
    LogText "Message sent: " & txtData
    txtData = ""
    'End If
End Sub

Private Sub Form_Load()
    Me.Show
    comSMS.PortOpen = True
    comSMS.InputLen = 0
    comSMS.InputMode = comInputModeText
    LogText "Program Started"
    InitModem
End Sub

Sub LogText(Text As String)
    Open "c:\SMS.log" For Append As #1
        Print #1, Text
    Close 1
End Sub

Sub InitModem()
    comSMS.Output = "AT&FW&C1%C0\N0E0V1&K4" & Chr$(13)
    lstTx.AddItem "AT&FW&C1%C0\N0E0V1&K4"
    LogText "Modem initialized"
End Sub


Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    tmrHangup_Timer
    End
End Sub

Private Sub optNumber_Click(Index As Integer)
    Select Case Index
        Case 0
            txtNumber = "0418707767"
        Case 1
            txtNumber = "018018767"
    End Select
End Sub

Private Sub tmrCom_Timer()
    Static buffer As String
    Select Case comSMS.CommEvent
        ' Handle each event or error by placing
        ' code below each case statement
    
        ' Errors
        Case comEventBreak     ' A Break was received.
        Case comEventFrame     ' Framing Error
        Case comEventOverrun   ' Data Lost.
        Case comEventRxOver    ' Receive buffer overflow.
        Case comEventRxParity  ' Parity Error.
        Case comEventTxFull    ' Transmit buffer full.
        Case comEventDCB       ' Unexpected error retrieving DCB]
    
        ' Events
        Case comEvCD           ' Change in the CD line.
        Case comEvCTS          ' Change in the CTS line.
        Case comEvDSR          ' Change in the DSR line.
        Case comEvRing         ' Change in the Ring Indicator.
        Case comEvReceive      ' Received RThreshold # of chars.
        Case comEvSend         ' There are SThreshold number of characters in the transmit buffer.
        Case comEvEOF          ' An EOF charater was found in the input stream
    End Select
   
    buffer = buffer & comSMS.Input
    If Len(buffer) > 0 Then
        If InStr(buffer, Chr(13)) <> 0 Then ' EOL character indicating a valid message received.
            For lp = 1 To Len(buffer)
                valCode = Asc(Mid(buffer, lp, 1))
                If valCode > 31 Then
                    strBuffer = strBuffer & Chr(valCode)
                End If
            Next lp
            buffer = strBuffer
            LogText "Data received: " & buffer
            If buffer = "NODIALTONE" Then
                cmdSendMsg.Enabled = False
            Else
                If txtMessage.Text <> "" Then
                    cmdSendMsg.Enabled = True
                Else
                    cmdSendMsg.Enabled = False
                End If
            End If
            
            lstRx.AddItem buffer
                        
            ' Message parsing code here
            If InStr(1, buffer, "RETURN TO QUIT)") > 0 Then
                txtData = txtMobileNumber
                cmdSendString_Click
            End If
            
            If InStr(1, buffer, "AND HIT RETURN KEY") > 0 Then
                txtData = txtMessage
                cmdSendString_Click
                txtMessage = ""
            End If
                
            If InStr(1, buffer, "BEING SENT") > 0 Then
                txtData = ""
                cmdSendString_Click
                tmrHangup.Enabled = True ' Start 5 second countdown before hanging up.
            End If
            
            buffer = ""
        End If
    End If
End Sub

Private Sub tmrHangup_Timer()
    txtData = "ATH"
    cmdSendString_Click
    LogText "ATH sent"
    tmrHangup.Enabled = False
End Sub

Private Sub txtMessage_Change()
    If Len(txtMessage) > 0 Then
        cmdSendMsg.Enabled = True
    Else
        cmdSendMsg.Enabled = False
    End If
    lblLen2 = Len(txtMessage)
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If Len(txtMessage) > 159 Or KeyAscii = 10 Or KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub
