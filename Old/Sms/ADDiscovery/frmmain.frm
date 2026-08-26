VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AD Discovery"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLogonTime 
      Caption         =   "Get IP's"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox lstDC 
      Height          =   4935
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdEnum 
      Caption         =   "Enum Domain"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.ListBox lstHosts 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtDomainName 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "MOJ_MASTER"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtDCName 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "MOJADMIN"
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6390
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDomainName 
      Caption         =   "Domain Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblDCName 
      Caption         =   "DC Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DoCancel As Boolean

Private Sub cmdEnum_Click()
    Dim DSN As String
    Dim SQL As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    
    ' Remove spaces from Domain and Server names
    txtDomainName.Text = UCase(Trim(txtDomainName.Text))
    txtDCName.Text = UCase(Trim(txtDCName.Text))
    
    lstDC.Clear
    
    'DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Domain.mdb;Persist Security Info=False"
    DSN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Domain Control;Data Source=(Local)"
    adoConn.Open DSN
    
    ' Domain Controllers
    GetInfo txtDCName.Text, txtDomainName.Text, SV_TYPE_DOMAIN_CTRL Or SV_TYPE_DOMAIN_BAKCTRL, lstDC
    
    For l = 0 To lstDC.ListCount - 1
        SQL = "SELECT * FROM tblHosts WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoRS.Open SQL, adoConn
        If adoRS.EOF Then
            SQL = "INSERT INTO tblHosts (Hostname, Domainname) VALUES ('" & lstDC.List(l) & "', '" & txtDomainName.Text & "')"
            adoConn.Execute SQL
        End If
        SQL = "UPDATE tblHosts SET isDC = 1 WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoConn.Execute SQL
        adoRS.Close
    Next l
    
    ' SQL Servers
    GetInfo txtDCName.Text, txtDomainName.Text, SV_TYPE_SQLSERVER, lstDC
    
    For l = 0 To lstDC.ListCount - 1
        SQL = "SELECT * FROM tblHosts WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoRS.Open SQL, adoConn
        If adoRS.EOF Then
            SQL = "INSERT INTO tblHosts (Hostname, Domainname) VALUES ('" & lstDC.List(l) & "', '" & txtDomainName.Text & "')"
            adoConn.Execute SQL
        End If
        SQL = "UPDATE tblHosts SET isSQL = 1 WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoConn.Execute SQL
        adoRS.Close
    Next l
    
    ' All
    GetInfo txtDCName.Text, txtDomainName.Text, SV_TYPE_WORKSTATION, lstDC
    
    For l = 0 To lstDC.ListCount - 1
        SQL = "SELECT * FROM tblHosts WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoRS.Open SQL, adoConn
        If adoRS.EOF Then
            SQL = "INSERT INTO tblHosts (Hostname, Domainname) VALUES ('" & lstDC.List(l) & "', '" & txtDomainName.Text & "')"
            adoConn.Execute SQL
        End If
        adoRS.Close
    Next l
    
    ' Print Queue
    GetInfo txtDCName.Text, txtDomainName.Text, SV_TYPE_PRINTQ_SERVER, lstDC
    
    For l = 0 To lstDC.ListCount - 1
        SQL = "SELECT * FROM tblHosts WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoRS.Open SQL, adoConn
        If adoRS.EOF Then
            SQL = "INSERT INTO tblHosts (Hostname, Domainname) VALUES ('" & lstDC.List(l) & "', '" & txtDomainName.Text & "')"
            adoConn.Execute SQL
        End If
        SQL = "UPDATE tblHosts SET isPrint = 1 WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoConn.Execute SQL
        adoRS.Close
    Next l
    
    ' RAS
    GetInfo txtDCName.Text, txtDomainName.Text, SV_TYPE_DIALIN_SERVER, lstDC
    
    For l = 0 To lstDC.ListCount - 1
        SQL = "SELECT * FROM tblHosts WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoRS.Open SQL, adoConn
        If adoRS.EOF Then
            SQL = "INSERT INTO tblHosts (Hostname, Domainname) VALUES ('" & lstDC.List(l) & "', '" & txtDomainName.Text & "')"
            adoConn.Execute SQL
        End If
        SQL = "UPDATE tblHosts SET isRAS = 1 WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoConn.Execute SQL
        adoRS.Close
    Next l
    
    ' UNIX
    GetInfo txtDCName.Text, txtDomainName.Text, SV_TYPE_SERVER_UNIX, lstDC
    
    For l = 0 To lstDC.ListCount - 1
        SQL = "SELECT * FROM tblHosts WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoRS.Open SQL, adoConn
        If adoRS.EOF Then
            SQL = "INSERT INTO tblHosts (Hostname, Domainname) VALUES ('" & lstDC.List(l) & "', '" & txtDomainName.Text & "')"
            adoConn.Execute SQL
        End If
        SQL = "UPDATE tblHosts SET isUNIX = 1 WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoConn.Execute SQL
        adoRS.Close
    Next l
    
    ' Browsers
    GetInfo txtDCName.Text, txtDomainName.Text, SV_TYPE_BACKUP_BROWSER Or SV_TYPE_MASTER_BROWSER, lstDC
    
    For l = 0 To lstDC.ListCount - 1
        SQL = "SELECT * FROM tblHosts WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoRS.Open SQL, adoConn
        If adoRS.EOF Then
            SQL = "INSERT INTO tblHosts (Hostname, Domainname) VALUES ('" & lstDC.List(l) & "', '" & txtDomainName.Text & "')"
            adoConn.Execute SQL
        End If
        SQL = "UPDATE tblHosts SET isBrowser = 1 WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoConn.Execute SQL
        adoRS.Close
    Next l
    
    ' Active Directory Extract
    GetDomain txtDCName.Text, txtDomainName.Text, lstDC
    
    For l = 0 To lstDC.ListCount - 1
        SQL = "SELECT * FROM tblHosts WHERE Hostname LIKE '" & lstDC.List(l) & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
        adoRS.Open SQL, adoConn
        If adoRS.EOF Then
            SQL = "INSERT INTO tblHosts (Hostname, Domainname) VALUES ('" & lstDC.List(l) & "', '" & txtDomainName.Text & "')"
            adoConn.Execute SQL
        End If
        adoRS.Close
    Next l
    
    adoConn.Close
    Set adoRS = Nothing
    Set adoConn = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub GetDomain(DCName As String, DomainName As String, Lst As ListBox)
    Dim domain As IADs
    Dim DSN As String
    Dim SQL As String
    Dim PCName As String
    
    Set domain = GetObject("WinNT://" & DomainName & ",domain")
    domain.Filter = Array("computer")
    
    For Each Computer In domain
        Lst.AddItem Computer.Name
    Next
    Set domain = Nothing
End Sub

Private Sub cmdExtract_Click()
    Dim domain As IADs
    Dim Computer As IADsComputer
    Dim DSN As String
    Dim SQL As String
    Dim PCName As String
    Dim bRet As Boolean
    Dim OS As String
    Dim IP As String
    Dim Role As String
'    Dim adoConn As ADODB.Connection
'    Dim adoRS As ADODB.Recordset
    
'    Set adoConn = CreateObject("ADODB.Connection")
'    Set adoRS = CreateObject("ADODB.Recordset")
    
    ' Remove spaces from Domain and Server names, and convert to uppercase
    txtDomainName.Text = UCase(Trim(txtDomainName.Text))
    txtDCName.Text = UCase(Trim(txtDCName.Text))
    
    Select Case cmdExtract.Caption
        Case "Extract"
            cmdExtract.Caption = "Cancel"
            DoCancel = False
            sbrStatus.SimpleText = "Extracting Computers from " & txtDomainName.Text
            sbrStatus.Refresh
        Case "Cancel"
            cmdExtract.Caption = "Extract"
            DoCancel = True
            sbrStatus.SimpleText = "Cancelling..."
            sbrStatus.Refresh
    End Select
    
    'DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\UMigrate.mdb;Persist Security Info=False"
    
    'adoConn.Open DSN

    Set domain = GetObject("WinNT://" & txtDomainName.Text & ",domain")
    domain.Filter = Array("computer")
    
    For Each Computer In domain
        PCName = Computer.Name
        sbrStatus.SimpleText = PCName
        DoEvents
        res = ""
        
        On Error Resume Next
        sbrStatus.SimpleText = PCName & ": IP"
        sbrStatus.Refresh
        'IP = ResolveHostname(PCName)
        'Info = GetInfo(PCname)
        DoEvents
        'If IP <> "" Then
            'sbrStatus.SimpleText = PCName & ": OS"
            'sbrStatus.Refresh
            'OS = Computer.OperatingSystem
            DoEvents
            'sbrStatus.SimpleText = PCName & ": Role"
            'sbrStatus.Refresh
            'Role = Computer.Role
            DoEvents
            If OS <> "" Then OS = " " & OS
            If Role <> "" Then Role = " " & Role
            lstHosts.AddItem PCName ' & ": " & IP & OS & Role
        'End If
        On Error GoTo 0
        SQL = ""
        frmMain.Refresh
        DoEvents
        If DoCancel Then
            sbrStatus.SimpleText = "Extract Cancelled"
            sbrStatus.Refresh
            Exit Sub
        End If
    Next
    
'    Set adoConn = Nothing
'    Set obj1 = Nothing
    sbrStatus.SimpleText = "Extract complete"
    sbrStatus.Refresh
'    frmMain.sbrStatus.SimpleText = "Extract complete"
'    frmMain.sbrStatus.Refresh
End Sub

Private Sub cmdLogonTime_Click()
    Dim DSN As String
    Dim SQL As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim adoRS2 As ADODB.Recordset
    Dim PCName As String
    Dim DCName As String
    Dim User As IADsUser
    
    Dim DC As IADsComputer
    
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    Set adoRS2 = CreateObject("ADODB.Recordset")
    
    ' Remove spaces from Domain and Server names
    txtDomainName.Text = UCase(Trim(txtDomainName.Text))
    txtDCName.Text = UCase(Trim(txtDCName.Text))
    
    'DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Domain.mdb;Persist Security Info=False"
    DSN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Domain Control;Data Source=(Local)"
    adoConn.Open DSN
    
    SQL = "SELECT * FROM tblHosts WHERE Domainname LIKE '" & txtDomainName.Text & "' ORDER BY Hostname"
    adoRS.Open SQL, adoConn
    Do Until adoRS.EOF
        PCName = adoRS("Hostname")
        IP = ""
        IP = ResolveHostname(PCName)
        DoEvents
        sbrStatus.SimpleText = PCName & ": " & IP
        frmMain.Refresh
        If IP <> "" Then
            SQL = "UPDATE tblHosts SET IP = '" & IP & "' WHERE Hostname LIKE '" & PCName & "' AND Domainname LIKE '" & txtDomainName.Text & "'"
            adoConn.Execute SQL
        End If
        adoRS.MoveNext
    Loop
    adoRS.Close
    adoConn.Close
    
    Set DC = Nothing
    Set User = Nothing
    Set adoRS2 = Nothing
    Set adoRS = Nothing
    Set adoConn = Nothing
End Sub

Private Sub Form_Load()
    bRet = SocketsInitialize()
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SocketsCleanup
End Sub

Private Sub txtDCName_Change()
    If txtDCName.Text <> "" And txtDomainName.Text <> "" Then
        cmdExtract.Enabled = True
    Else
        cmdExtract.Enabled = False
    End If
End Sub

Private Sub txtDomainName_Change()
    If txtDCName.Text <> "" And txtDomainName.Text <> "" Then
        cmdExtract.Enabled = True
    Else
        cmdExtract.Enabled = False
    End If
End Sub

