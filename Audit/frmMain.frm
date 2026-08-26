VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PC Audit"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdServices 
      Caption         =   "Services"
      Height          =   495
      Left            =   1680
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   5400
      Width           =   1575
   End
   Begin VB.ListBox lstOutput 
      Height          =   5325
      Left            =   3240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   7695
   End
   Begin VB.CommandButton cmdExecutables 
      Caption         =   "Executables"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdApplications 
      Caption         =   "Applications"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdHotfixes 
      Caption         =   "Hotfixes"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdServicePack 
      Caption         =   "Service Pack"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOS 
      Caption         =   "OS"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdDomain 
      Caption         =   "Domain"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetName 
      Caption         =   "Set"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtComputername 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdGetMAC 
      Caption         =   "MAC Address"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetIP 
      Caption         =   "IP Address"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdLocalAccounts 
      Caption         =   "Local Accounts"
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Image imgUnknown 
      Height          =   240
      Left            =   4680
      Picture         =   "frmMain.frx":030A
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Left            =   4680
      Picture         =   "frmMain.frx":0454
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Left            =   4680
      Picture         =   "frmMain.frx":059E
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblComputername 
      Caption         =   "Computername:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bStable As Boolean ' Used to indicate when the application has finished loading

Private Sub cmdApplications_Click()
    Screen.MousePointer = vbHourglass
    frmMain.lstOutput.AddItem "Application List:"
    AppList = GetApplicationList(strComputername)
    frmMain.lstOutput.AddItem "----------------------------------------"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    frmMain.lstOutput.Clear
End Sub

Private Sub cmdDomain_Click()
    Dim strRegKey As String
    Dim strRegValue As String
    
    Screen.MousePointer = vbHourglass
    strRegKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    strRegValue = "DefaultDomainName"
    
    strDomainName = GetRegValue(strComputername, strRegKey, strRegValue)
    frmMain.lstOutput.AddItem "Domain: " & strDomainName
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExecutables_Click()
    Screen.MousePointer = vbHourglass
    frmMain.lstOutput.AddItem "Executables:"
    
    With FP
      .sFileRoot = "Z:\"
      .sFileNameExt = "*.exe"
      .bRecurse = True
      .bList = True
   End With
   
   'SearchForFiles (FP)
    
    frmMain.lstOutput.AddItem "----------------------------------------"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdHotfixes_Click()
    Screen.MousePointer = vbHourglass
    frmMain.lstOutput.AddItem "Hotfixes:"
    Hotfixes = GetHotfixes(strComputername)
    frmMain.lstOutput.AddItem "----------------------------------------"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOS_Click()
    Screen.MousePointer = vbHourglass
    GetOSInfo strComputername, PCInfo
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
    Dim Filename As String
    Dim App_Path As String
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    App_Path = App.Path
    
    If Right(App_Path, 1) = "\" Then
    Else
        App_Path = App_Path & "\"
    End If
    
    Filename = App_Path & frmMain.txtComputername & ".txt"
    
    Open Filename For Output As #1
        For i = 0 To frmMain.lstOutput.ListCount - 1
            Print #1, frmMain.lstOutput.List(i)
        Next i
    Close 1
    
    MsgBox "Saved to " & Filename, vbInformation + vbOKOnly, "Save complete"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdServicePack_Click()
    Dim strRegKey As String
    Dim strRegValue As String
    
    Screen.MousePointer = vbHourglass
    strRegKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    strRegValue = "CSDVersion"
    
    strServicePack = GetRegValue(strComputername, strRegKey, strRegValue)
    frmMain.lstOutput.AddItem strServicePack
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdLocalAccounts_Click()
    Screen.MousePointer = vbHourglass
    frmMain.lstOutput.AddItem "Local Accounts:"
    ListUsers (strComputername)
    frmMain.lstOutput.AddItem "----------------------------------------"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdGetIP_Click()
    Screen.MousePointer = vbHourglass
    strIPAddress = GetIPAddress(strComputername)
    frmMain.lstOutput.AddItem "IP address: " & strIPAddress
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdGetMAC_Click()
    Screen.MousePointer = vbHourglass
    strMACAddress = RemoteMACAddress(strIPAddress)
    frmMain.lstOutput.AddItem "MAC address: " & strMACAddress
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdServices_Click()
    Screen.MousePointer = vbHourglass
    Dim NumberOfServices As Long
    frmMain.lstOutput.AddItem "Services: "
    NumberOfServices = EnumSystemServices(SERVICE_STATE_ALL, strComputername)
    frmMain.lstOutput.AddItem "----------------------------------------"
    Screen.MousePointer = vbDefault
    txtComputername.SetFocus
End Sub

Private Sub cmdSetName_Click()
    Screen.MousePointer = vbHourglass
    
    Cleanup
    
    strComputername = frmMain.txtComputername.Text
    imgUp.Visible = False
    imgUnknown.Visible = True
    imgDown.Visible = False
    cmdClear_Click
    frmMain.lstOutput.AddItem "Computername: " & strComputername
    If DestinationReachable(strComputername) Then
        imgUp.Visible = True
        imgUnknown.Visible = False
        imgDown.Visible = False
    Else
        imgUp.Visible = False
        imgUnknown.Visible = False
        imgDown.Visible = True
    End If
    Screen.MousePointer = vbDefault
    txtComputername.SetFocus
End Sub

Private Sub Form_Load()
    txtComputername = Environ$("COMPUTERNAME")
    Disable
    bStable = True
End Sub

Private Sub txtComputername_Change()
    If bStable Then
        imgUp.Visible = False
        imgUnknown.Visible = True
        imgDown.Visible = False
        Disable
    End If
End Sub

Private Sub txtComputername_Click()
    txtComputername.SelStart = 0
    txtComputername.SelLength = Len(txtComputername.Text)
End Sub

Private Sub txtComputername_GotFocus()
    txtComputername.SelStart = 0
    txtComputername.SelLength = Len(txtComputername.Text)
End Sub

Private Sub txtComputername_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdSetName_Click
        txtComputername.SelStart = 0
        txtComputername.SelLength = Len(txtComputername.Text)
    End If
End Sub
