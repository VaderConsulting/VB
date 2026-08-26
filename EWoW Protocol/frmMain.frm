VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EWoW-C Protocol"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "clr"
      Height          =   250
      Left            =   3960
      TabIndex        =   32
      Top             =   5520
      Width           =   375
   End
   Begin VB.ListBox lstEWoW 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Frame fmePort 
      Caption         =   "Port"
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdClosePort 
         Caption         =   "Close"
         Height          =   315
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpenPort 
         Caption         =   "Open"
         Height          =   315
         Left            =   1560
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbPort 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Text            =   "Com1"
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSCommLib.MSComm comEWoW 
      Left            =   120
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   2400
   End
   Begin VB.Frame fmeDevice 
      Caption         =   "Device"
      Height          =   3135
      Left            =   3000
      TabIndex        =   18
      Top             =   960
      Width           =   1335
      Begin VB.CommandButton cmdDevice 
         Caption         =   "Flag 2"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "Flag 1"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "Dump 2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "Dump 1"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "Base 2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "Base 1"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "Mine 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdDevice 
         Caption         =   "Mine 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fmeClan 
      Caption         =   "Clans"
      Height          =   3135
      Left            =   1560
      TabIndex        =   9
      Top             =   960
      Width           =   1335
      Begin VB.CommandButton cmdShot 
         Caption         =   "Clan 8"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdShot 
         Caption         =   "Clan 7"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdShot 
         Caption         =   "Clan 6"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdShot 
         Caption         =   "Clan 5"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdShot 
         Caption         =   "Clan 4"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdShot 
         Caption         =   "Clan 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdShot 
         Caption         =   "Clan 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdShot 
         Caption         =   "Clan 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fmeReferee 
      Caption         =   "Referee"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1335
      Begin VB.CommandButton cmdReferee 
         Caption         =   "Respawn"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdReferee 
         Caption         =   "1 Health"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdReferee 
         Caption         =   "All Health"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdReferee 
         Caption         =   "1 clip"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdReferee 
         Caption         =   "All clips"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdReferee 
         Caption         =   "Abort"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdReferee 
         Caption         =   "Punish"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdReferee 
         Caption         =   "Freeze"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
    lstEWoW.Clear
End Sub

Private Sub cmdClosePort_Click()
    CloseCurrentPort
End Sub

Private Sub cmdDevice_Click(Index As Integer)
    Dim iDeviceCode As Integer
    
    Select Case Index
        Case 0 ' Mine 1
            iDeviceCode = 90
        Case 1 ' Mine 2
            iDeviceCode = 181
        Case 2 ' Base 1
            iDeviceCode = 130
        Case 3 ' Base 2
            iDeviceCode = 131
        Case 4 ' Dump 1
            iDeviceCode = 132
        Case 5 ' Dump 2
            iDeviceCode = 133
        Case 6 ' Flag 1
            iDeviceCode = 134
        Case 7 ' Flag 2
            iDeviceCode = 135
    End Select
    
    SendEWoW iDeviceCode
End Sub

Private Sub cmdOpenPort_Click()
    If comEWoW.PortOpen = False Then
        comEWoW.PortOpen = True
    End If
End Sub

Private Sub cmdReferee_Click(Index As Integer)
    Dim iRefereeCode As Integer
    
    Select Case Index
        Case 0 ' Freeze
            iRefereeCode = 113
        Case 1 ' Punish
            iRefereeCode = 114
        Case 2 ' Abort
            iRefereeCode = 115
        Case 3 ' All Clips
            iRefereeCode = 121
        Case 4 ' 1 Clip
            iRefereeCode = 122
        Case 5 ' All Health
            iRefereeCode = 124
        Case 6 ' 1 Health
            iRefereeCode = 125
        Case 7 ' Respawn
            iRefereeCode = 126
    End Select
    
    SendEWoW iRefereeCode
End Sub

Private Sub cmdShot_Click(Index As Integer)
    Dim iShotCode As Integer
    
    Select Case Index
        Case 0 ' Clan 1
            iShotCode = 170
        Case 1 ' Clan 2
            iShotCode = 42
        Case 2 ' Clan 3
            iShotCode = 74
        Case 3 ' Clan 4
            iShotCode = 154
        Case 4 ' Clan 5
            iShotCode = 85
        Case 5 ' Clan 6
            iShotCode = 37
        Case 6 ' Clan 7
            iShotCode = 149
        Case 7 ' Clan 8
            iShotCode = 165
        Case 8 ' Spare 1
            iShotCode = 106
        Case 9 ' Spare 2
            iShotCode = 101
    End Select
    
    SendEWoW iShotCode
End Sub

Private Sub comEWoW_OnComm()
    Dim strInput As String
    
    
    Select Case comEWoW.CommEvent
        ' Errors
        Case comEventBreak      ' A Break was received.
        Case comEventCDTO       ' CD (RLSD) Timeout.
        Case comEventCTSTO      ' CTS Timeout.
        Case comEventDSRTO      ' DSR Timeout.
        Case comEventFrame      ' Framing Error.
        Case comEventOverrun    ' Data Lost.
        Case comEventRxOver     ' Receive buffer overflow.
        Case comEventRxParity   ' Parity Error.
        Case comEventTxFull     ' Transmit buffer full.
        Case comEventDCB        ' Unexpected error retrieving DCB
        
        ' Events
        Case comEvCD            ' Change in the CD line.
        Case comEvCTS           ' Change in the CTS line.
        Case comEvDSR           ' Change in the DSR line.
        Case comEvRing          ' Change in the Ring Indicator.
        Case comEvReceive       ' Received RThreshold # of chars.
            strInput = comEWoW.Input
            HandleInput (strInput)
        Case comEvSend ' There are SThreshold number of characters in the transmit buffer.
        Case comEvEOF  ' An EOF character was found in the input stream.
     End Select
End Sub

Private Sub Form_Load()
    cmbPort.AddItem "Com1"
    cmbPort.AddItem "Com2"
    cmbPort.AddItem "Com3"
    cmbPort.AddItem "Com4"
    cmbPort.Text = "Com1"
    
    With comEWoW
       .CommPort = 1
       .Handshaking = 0 '- comNone
       .RThreshold = 3
       .RTSEnable = False
       .Settings = "2400,n,8,1"
       .SThreshold = 0
       .PortOpen = True
    End With

End Sub

Private Sub cmbPort_Click()
    Dim iPortNumber As Integer
    
    CloseCurrentPort
    
    iPortNumber = CInt(Right(cmbPort.Text, 1))
    comEWoW.CommPort = iPortNumber
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CloseCurrentPort
End Sub

Private Function SendEWoW(iCode As Integer)
    If Not comEWoW.PortOpen Then
        comEWoW.PortOpen = True
    End If
    
    comEWoW.Output = Chr(iCode) & Chr(iCode) & Chr(25)
End Function

Private Function CloseCurrentPort()
    If comEWoW.PortOpen Then
        comEWoW.PortOpen = False
    End If
    
    lstEWoW.Clear
End Function

Private Function HandleInput(strInputBuffer As String)
    Dim intInput() As Integer
    Dim strConvertedInput As String
    Dim strConvertedPortion As String
    
    ReDim intInput(Len(strInputBuffer))
    
    If strInputBuffer <> "" Then
        For i = 1 To Len(strInputBuffer)
            intInput(i - 1) = Asc(Mid(strInputBuffer, i, 1))
            strConvertedPortion = CStr(intInput(i - 1))
            Do Until Len(strConvertedPortion) = 3
                strConvertedPortion = "0" & strConvertedPortion
            Loop
            strConvertedInput = strConvertedInput & Trim(strConvertedPortion) & " "
        Next i
        strConvertedInput = Trim(strConvertedInput)
        
        lstEWoW.AddItem strConvertedInput & "  " & strInputBuffer, 0
    End If
End Function
