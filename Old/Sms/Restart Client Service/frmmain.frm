VERSION 5.00
Object = "{B186D399-9B91-41CB-9242-E12B96CA99E1}#1.0#0"; "DRPing.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMS Client Restarter"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3450
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   3450
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwHosts 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5318
      View            =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hostname"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imlImages 
      Left            =   2400
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":625A
            Key             =   "Yes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BFBC
            Key             =   "No"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11D1E
            Key             =   "Question"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtHostname 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin Ping.drPing Ping 
      Left            =   3120
      Top             =   120
      _ExtentX        =   423
      _ExtentY        =   423
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function StartService(Hostname As String, ServiceName As String) As String
    Dim oService As Service.clsService
    Dim oStatus As Service.SERVICE_STATUS
    Set oService = CreateObject("Service.clsService")
    oService.vbStartService Hostname, ServiceName
    Set oService = Nothing
End Function

Function StopService(Hostname As String, ServiceName As String) As String
    Dim oService As Service.clsService
    Dim oStatus As Service.SERVICE_STATUS
    Set oService = CreateObject("Service.clsService")
    oService.vbStopService Hostname, ServiceName
    Set oService = Nothing
End Function

Function PauseService(Hostname As String, ServiceName As String) As String
    Dim oService As Service.clsService
    Dim oStatus As Service.SERVICE_STATUS
    Set oService = CreateObject("Service.clsService")
    oService.vbPauseService Hostname, ServiceName
    Set oService = Nothing
End Function

Function RestartService(Hostname As String, ServiceName As String) As String
    Dim oService As Service.clsService
    Dim oStatus As Service.SERVICE_STATUS
    Set oService = CreateObject("Service.clsService")
    StopService Hostname, ServiceName
    StartService Hostname, ServiceName
    Set oService = Nothing
End Function

Private Sub cmdAdd_Click()
    lvwHosts.ListItems.Add , txtHostname.Text & lvwHosts.ListItems.Count, txtHostname.Text, , "Question"
    txtHostname.Text = ""
    txtHostname.SetFocus
End Sub

Private Sub cmdStart_Click()
    Dim Hostname As String
    Dim Timeout As Long
    Dim IP As String
    Dim PingResult As String
    Timeout = 500
    For lp = 1 To lvwHosts.ListItems.Count
        Hostname = lvwHosts.ListItems(lp).Text
        IP = GetIPAddress(Hostname)
        PingResult = Ping.Ping(IP, Timeout)
        If IsNumeric(PingResult) Then
            If CLng(PingResult) > Timeout Then
                lvwHosts.ListItems(lp).SmallIcon = "No"
            Else
                lvwHosts.ListItems(lp).SmallIcon = "Yes"
                RestartService Hostname, "clisvc"
            End If
        Else
            lvwHosts.ListItems(lp).SmallIcon = "No"
        End If
    Next lp
End Sub

Private Sub Form_Load()
    Me.Show
    Me.Refresh
End Sub
