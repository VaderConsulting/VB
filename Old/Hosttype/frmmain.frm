VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Determine Hosttype..."
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtHostname 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblOS 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   5655
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   5655
   End
   Begin VB.Label lblHostname 
      Caption         =   "Hostname"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DoUnload As Boolean

Private Function GetHostType(Host As Computer) As Boolean
    Dim strFilename As String, tmpString As String
    Dim retval As Long
    strFilename = "\\" & Host.Name & "\c$\boot.ini"
    lblStatus = "Determining OS on " & Host.Name
    lblStatus.Refresh
    retval = Connect2("", "\\" & Host.Name & "\IPC$", txtUsername, txtPassword)
    On Error Resume Next
    
    Host.OS.Version = ""
    Host.OS.Name = ""
    Host.OS.Primary = False
    Host.OS.ServicePack = ""
    Host.OS.Version = ""
    Host.OS.Server = False
    Host.OS.Workstation = False
    
    Open strFilename For Input As #1
        If Err.Number = 0 Then
            Do Until EOF(1) Or (Host.OS.Version <> "")
                Line Input #1, tmpString
                If InStr(1, UCase(tmpString), "WORKSTATION") > 0 And InStr(1, UCase(tmpString), "3.50") <> 0 Then
                    Host.OS.Workstation = True
                    Host.OS.Name = "Windows NT Workstation"
                    Host.OS.Version = "3.50"
                End If
                If InStr(1, UCase(tmpString), "WORKSTATION") > 0 And InStr(1, UCase(tmpString), "3.51") <> 0 Then
                    Host.OS.Workstation = True
                    Host.OS.Name = "Windows NT Workstation"
                    Host.OS.Version = "3.51"
                End If
                If InStr(1, UCase(tmpString), "WORKSTATION") > 0 And InStr(1, UCase(tmpString), "4") <> 0 Then
                    Host.OS.Workstation = True
                    Host.OS.Name = "Windows NT Workstation"
                    Host.OS.Version = "4"
                End If
                If InStr(1, UCase(tmpString), "PROFESSIONAL") > 0 Then
                    Host.OS.Workstation = True
                    Host.OS.Name = "Windows 2000 Professional"
                    Host.OS.Version = "5"
                End If
                If InStr(1, UCase(tmpString), "SERVER") > 0 And InStr(1, UCase(tmpString), "3.50") <> 0 Then
                    Host.OS.Server = True
                    Host.OS.Name = "Windows NT Server"
                    Host.OS.Version = "3.50"
                End If
                If InStr(1, UCase(tmpString), "SERVER") > 0 And InStr(1, UCase(tmpString), "3.51") <> 0 Then
                    Host.OS.Server = True
                    Host.OS.Name = "Windows NT Server"
                    Host.OS.Version = "3.51"
                End If
                If InStr(1, UCase(tmpString), "SERVER") > 0 And InStr(1, UCase(tmpString), "4") <> 0 And InStr(1, UCase(tmpString), "ADVANCED") = 0 And InStr(1, UCase(tmpString), "TERMINAL") = 0 Then
                    Host.OS.Server = True
                    Host.OS.Name = "Windows NT Server"
                    Host.OS.Version = "4"
                End If
                If InStr(1, UCase(tmpString), "SERVER") > 0 And InStr(1, UCase(tmpString), "4") <> 0 And InStr(1, UCase(tmpString), "ADVANCED") <> 0 And InStr(1, UCase(tmpString), "TERMINAL") = 0 Then
                    Host.OS.Server = True
                    Host.OS.Name = "Windows NT Advanced Server"
                    Host.OS.Version = "4"
                End If
                If InStr(1, UCase(tmpString), "SERVER") > 0 And InStr(1, UCase(tmpString), "4") <> 0 And InStr(1, UCase(tmpString), "TERMINAL") <> 0 Then
                    Host.OS.Server = True
                    Host.OS.Name = "Windows NT Terminal Server"
                    Host.OS.Version = "4"
                End If
                If InStr(1, UCase(tmpString), "SERVER") > 0 And InStr(1, UCase(tmpString), "2000") <> 0 And InStr(1, UCase(tmpString), "ADVANCED") = 0 And InStr(1, UCase(tmpString), "TERMINAL") = 0 Then
                    Host.OS.Server = True
                    Host.OS.Name = "Windows 2000 Server"
                    Host.OS.Version = "5"
                End If
                If InStr(1, UCase(tmpString), "SERVER") > 0 And InStr(1, UCase(tmpString), "2000") <> 0 And InStr(1, UCase(tmpString), "ADVANCED") <> 0 And InStr(1, UCase(tmpString), "TERMINAL") = 0 Then
                    Host.OS.Server = True
                    Host.OS.Name = "Windows 2000 Advanced Server"
                    Host.OS.Version = "5"
                End If
                If InStr(1, UCase(tmpString), "SERVER") > 0 And InStr(1, UCase(tmpString), "2000") <> 0 And InStr(1, UCase(tmpString), "TERMINAL") <> 0 Then
                    Host.OS.Server = True
                    Host.OS.Name = "Windows 2000 Terminal Server"
                    Host.OS.Version = "5"
                End If
            Loop
        End If
    Close 1
    DoEvents
    retval = DisConnect2("\\" & Host.Name & "\IPC$", True)
    If Host.OS.Version <> "" Then
        GetHostType = True
    Else
        GetHostType = False
    End If
End Function

Private Sub cmdQuery_Click()
    Dim wmiLocator As SWbemLocator
    Dim oServer As SWbemServices
    Dim oWMI As Object, bRes As Boolean
    Dim E As Long
    Dim OS As Variant
    Dim DSN As String
    Dim SQL As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim Retries As Integer
    
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    
    ' Remove spaces from Domain and Server names
    'txtDomainName.Text = UCase(Trim(txtDomainName.Text))
    'txtDCName.Text = UCase(Trim(txtDCName.Text))
    
    'DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents\VB\SMS\ADDiscovery\Domain.mdb;Persist Security Info=False"
    DSN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Domain Control;Data Source=(Local)"
    
    adoConn.Open DSN
    
    Host.OS.Version = ""
    Host.OS.Name = ""
    Host.OS.Primary = False
    Host.OS.ServicePack = ""
    Host.OS.Version = ""
    Host.OS.Server = False
    Host.OS.Workstation = False
    
    lblOS = ""
    
    ' Attempt to use WMI
    Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")
    
    On Error Resume Next
    
    lblStatus = "Connecting to " & txtHostname
    frmMain.Refresh
    Set oServer = wmiLocator.ConnectServer(txtHostname, "root/CIMV2", txtUsername, txtPassword)
    E = Err.Number
    DoEvents
    If E <> 0 Then
        lblStatus = "WMI error connecting to " & txtHostname & " (" & E & ")"
    Else
        ' Get Operating System
        lblStatus = "Determining Operating System (WMI Query)"
        frmMain.Refresh
        Set oWMI = oServer.ExecQuery("SELECT * FROM Win32_OperatingSystem WHERE Primary = TRUE")
        DoEvents
        E = Err.Number
        If E <> 0 Then
            lblStatus = "WMI error querying OS on " & txtHostname & " (" & E & ")"
            frmMain.Refresh
        Else
            For Each OS In oWMI
                
                Host.Name = txtHostname
                Host.OS.Name = OS.Caption
                Host.OS.Version = OS.Version
                If OS.servicepackmajorversion <> "" Then
                    Host.OS.ServicePack = OS.CSDVersion
                Else
                    Host.OS.ServicePack = ""
                End If
                If OS.Primary = True Then
                    Host.OS.Primary = True
                Else
                    Host.OS.Primary = False
                End If
                lblOS = Host.OS.Name & " " & Host.OS.Version & " " & Host.OS.ServicePack
                lblStatus = "Done"
                frmMain.Refresh
            Next
        End If
    End If
    ' Attempt to read boot.ini
    If Trim(lblOS) = "" Then
        'lblOS = "Unable to determine OS on " & txtHostname & " (WMI)"
        lblStatus = "Attempting to open boot.ini"
        frmMain.Refresh
        Host.Name = UCase(txtHostname.Text)
        DoEvents
        On Error GoTo 0
        bRes = GetHostType(Host)
        DoEvents
        If bRes Then
            lblStatus = "Done"
            lblOS = Host.OS.Name & " Version " & Host.OS.Version
        Else
            lblStatus = "Unable to determine OS on " & txtHostname & " (boot.ini)"
            lblOS = ""
        End If
    End If
    
    On Error Resume Next
    
    SQL = "SELECT Retries FROM tblHosts WHERE Hostname = '" & txtHostname.Text & "'"
    adoRS.Open SQL, adoConn
    Retries = adoRS("Retries") + 0
    
    If Retries > 3 Then
        adoRS.Close
        adoConn.Close
        Set adoRS = Nothing
        Set adoConn = Nothing
        Exit Sub
    End If
    
    If lblOS.Caption <> "" Then
        SQL = "UPDATE tblHosts SET Type = '" & lblOS.Caption & "' WHERE Hostname = '" & txtHostname.Text & "'"
        adoConn.Execute SQL
    Else
        SQL = "UPDATE tblHosts SET Retries = " & Retries + 1 & ", Type = '' WHERE Hostname = '" & txtHostname.Text & "'"
        adoConn.Execute SQL
    End If
    
    adoRS.Close
    adoConn.Close
    
    frmMain.Refresh
    txtHostname.SelStart = 0
    txtHostname.SelLength = Len(txtHostname)
    txtHostname.SetFocus
    
    Set adoRS = Nothing
    Set adoConn = Nothing
End Sub

Private Sub Form_Load()
    Dim DSN As String
    Dim SQL As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim adoRS2 As ADODB.Recordset
    Dim i, j
    
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    Set adoRS2 = CreateObject("ADODB.Recordset")
    
    Randomize Timer
    ' Remove spaces from Domain and Server names
    'txtDomainName.Text = UCase(Trim(txtDomainName.Text))
    'txtDCName.Text = UCase(Trim(txtDCName.Text))
    
    'DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents\VB\SMS\ADDiscovery\Domain.mdb;Persist Security Info=False"
    DSN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Domain Control;Data Source=(Local)"
    
    Me.Show
    Me.Refresh
    On Error GoTo erra
    
    adoConn.Open DSN
    
    txtUsername = Environ$("USERDOMAIN") & "\" & Environ$("USERNAME")
    'If Command$ = "DEBUG" Then
    txtUsername = "moj_master\robinsod"
    txtPassword.Text = "password"
    'End If
    
    SQL = "SELECT Hostname, inProgress FROM tblHosts WHERE isUP = 1 AND (Type = '' OR Type = NULL) AND Retries < 4 ORDER BY Hostname"
    adoRS.Open SQL, adoConn
    Do Until adoRS.EOF
        txtHostname.Text = adoRS("Hostname")
        If adoRS("inProgress") = 0 Then
            SQL = "SELECT * FROM tblHosts WHERE Hostname = '" & txtHostname.Text & "'"
            adoRS2.Open SQL, adoConn
            If adoRS2("inProgress") = 0 Then
                SQL = "UPDATE tblHosts SET inProgress = 1 WHERE Hostname = '" & txtHostname.Text & "'"
                adoConn.Execute SQL
                cmdQuery_Click
                SQL = "UPDATE tblHosts SET inProgress = 0 WHERE Hostname = '" & txtHostname.Text & "'"
                adoConn.Execute SQL
            End If
            adoRS2.Close
        End If
        DoEvents
        
        If DoUnload Then
            Unload Me
            End
        End If
        i = Int(Rnd(1) * 10000) + 1
        For j = 1 To i
            DoEvents
        Next
        adoRS.MoveNext
    Loop
    
    adoRS.Close
    Set adoRS2 = Nothing
    Set adoRS = Nothing
    Set adoConn = Nothing
    Exit Sub
erra:
    Shell App.Path & "\hosttype.exe"
    Unload Me
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = False
    DoUnload = True
    frmMain.lblStatus = "Cancel pending"
    frmMain.Refresh
End Sub
