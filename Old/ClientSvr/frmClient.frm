VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2730
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   2730
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtShutdown 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtClientPort 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "2000"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtHostPort 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "3000"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtHost 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Text            =   "192.168.0.1"
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Server"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtReceive 
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskBroadcast 
      Left            =   2280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblClient 
      Alignment       =   2  'Center
      Caption         =   "Client"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblPorts 
      Alignment       =   2  'Center
      Caption         =   "IP Port Addressing"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      Caption         =   "Host"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblHostname 
      Alignment       =   2  'Center
      Caption         =   "Host IP Address"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUpdate_Click()
    Renew
End Sub

Private Sub Form_Load()
    wskClient.Protocol = sckTCPProtocol
    wskClient.Bind txtClientPort, "192.168.0.1"
    wskClient.Listen
    wskBroadcast.Connect "192.168.0.1", txtHostPort
    'wskBroadcast.Protocol = sckTCPProtocol
End Sub

Private Sub wskClient_ConnectionRequest(ByVal requestID As Long)
    If wskClient.State <> sckClosed Then wskClient.Close
    wskClient.Accept requestID
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    wskClient.GetData strData, vbString
    txtReceive.Text = strData
    If Left$(txtReceive.Text, 8) <> "SHUTDOWN" Then
        If txtReceive.MultiLine = True Then
            Shell Left$(txtReceive, Len(txtReceive) - 2), vbMinimizedNoFocus
        Else
        Shell txtReceive, vbMinimizedNoFocus
        End If
    End If
    If Left$(txtReceive.Text, 8) = "SHUTDOWN" Then
        txtShutdown = Right$(txtReceive, Len(txtReceive) - 8)
        frmShutdown.Show vbModal
    End If
End Sub

Private Sub Renew()
    'On Error Resume Next
        wskData = wskBroadcast.LocalHostName & " " & wskBroadcast.LocalIP
        wskBroadcast.SendData wskData
    'On Error GoTo 0
End Sub

