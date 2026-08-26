VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   3960
   ClientLeft      =   2700
   ClientTop       =   2310
   ClientWidth     =   6105
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeShutdown 
      Caption         =   "Shutdown"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   3840
      TabIndex        =   15
      Top             =   2760
      Width           =   1935
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "frmMain.frx":0442
         Top             =   240
         Width           =   375
      End
      Begin ComCtl2.UpDown updTime 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "txtTime"
         BuddyDispid     =   196610
         OrigLeft        =   1560
         OrigTop         =   240
         OrigRight       =   1800
         OrigBottom      =   615
         Max             =   60
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdShutdown 
         Caption         =   "Shutdown"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "Seconds"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.TextBox txtHostPort 
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Text            =   "3000"
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtClientPort 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Text            =   "2000"
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstClients 
      Height          =   2010
      Left            =   2640
      TabIndex        =   11
      Top             =   360
      Width           =   3135
   End
   Begin VB.ComboBox cmbTransmit 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   2415
   End
   Begin VB.Frame fmeApp 
      Caption         =   "Options"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2415
      Begin VB.OptionButton optApp 
         Caption         =   "Normal, without Focus"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   2175
      End
      Begin VB.OptionButton optApp 
         Caption         =   "Normal, with Focus"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin VB.OptionButton optApp 
         Caption         =   "Minimized, without Focus"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton optApp 
         Caption         =   "Minimized, with Focus"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.OptionButton optApp 
         Caption         =   "Maximized, with Focus"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optApp 
         Caption         =   "Hidden"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSWinsockLib.Winsock wskListen 
      Left            =   120
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Execute"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   600
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label lblClients 
      Alignment       =   2  'Center
      Caption         =   "Client Workstations"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   120
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblTransmit 
      Alignment       =   2  'Center
      Caption         =   "Application to execute"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuApps 
         Caption         =   "Applications..."
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigure 
         Caption         =   "Configure..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    If txtManual <> "" Then
        cmbTransmit.AddItem txtManual
        txtManual = ""
    End If
End Sub

Private Sub cmdShutdown_Click()
    wskServer.SendData "SHUTDOWN" & txtTime
End Sub

Private Sub Form_Load()
    wskServer.Connect "192.168.0.1", txtClientPort
    wskListen.Bind txtHostPort, "192.168.0.1"
    wskListen.Listen
    With cmbTransmit
        .AddItem "Notepad.exe"
        .AddItem "Winfile.exe"
        .AddItem "Calc.exe"
        .ListIndex = 0
    End With
End Sub

Private Sub lstClients_Click()
    cmdSend.Enabled = True
    fmeShutdown.Enabled = True
End Sub

Private Sub cmdSend_Click()
    wskServer.SendData cmbTransmit.List(cmbTransmit.ListIndex)
End Sub

Private Sub mnuApps_Click()
    frmApplications.Show
End Sub

Private Sub mnuConfigure_Click()
    frmConfigure.Show
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub wskListen_ConnectionRequest(ByVal requestID As Long)
    If wskListen.State <> sckClosed Then wskListen.Close
    wskListen.Accept requestID
End Sub

Private Sub wskListen_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    wskListen.GetData strData, vbString
    lstClients.AddItem strData
    lblTotal.Caption = "Total clients: " & lstClients.ListCount
End Sub

