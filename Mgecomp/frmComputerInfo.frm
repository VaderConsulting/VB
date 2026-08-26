VERSION 5.00
Begin VB.Form frmComputerInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computer Information"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmComputerInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   615
   End
   Begin VB.Image imgStatus 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblRole 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Role:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblComputername 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmComputerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim strComputer As String
    Dim strComputerRole As String
    Dim objWMIService
    Dim colComputers As Object
    Dim objComputer As Object
    Dim strIPAddress As String
    Dim isUp As Boolean
    
    strComputer = lblComputername.Caption
    
    imgStatus.Picture = frmMain.imlImages.ListImages("UnknownComputer").Picture
    imgStatus.ToolTipText = "Unknown Status"
    
    lblInfo.Caption = "Please wait, determining status"
    Me.Refresh
    strIPAddress = ""
    strIPAddress = GetIPAddress(strComputer)
    
    isUp = Ping(strIPAddress)
    
    If isUp Then
        imgStatus.Picture = frmMain.imlImages.ListImages("UpComputer").Picture
        imgStatus.ToolTipText = "Computer is Up"
        Me.Refresh
        On Error Resume Next
        lblInfo.Caption = "Please wait, retrieving info..."
        
        Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        Set colComputers = objWMIService.ExecQuery("Select DomainRole from Win32_ComputerSystem")
        For Each objComputer In colComputers
            Select Case objComputer.DomainRole
                Case 0
                    strComputerRole = "Standalone Workstation"
                Case 1
                    strComputerRole = "Member Workstation"
                Case 2
                    strComputerRole = "Standalone Server"
                Case 3
                    strComputerRole = "Member Server"
                Case 4
                    strComputerRole = "Backup Domain Controller"
                Case 5
                    strComputerRole = "Primary Domain Controller"
            End Select
            lblRole.Caption = strComputerRole
        Next
    Else
        imgStatus.Picture = frmMain.imlImages.ListImages("DownComputer").Picture
        imgStatus.ToolTipText = "Computer is Down"
    End If
    lblInfo.Caption = "Idle"
End Sub

Private Sub Form_Load()
    Me.Refresh
End Sub

Private Sub Image1_Click()

End Sub

