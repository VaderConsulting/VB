VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Policy Distribution Applet"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6825
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstErrors 
      Height          =   1230
      Left            =   120
      TabIndex        =   26
      Top             =   6240
      Width           =   6615
   End
   Begin VB.Frame fmeMode 
      Caption         =   "Mode"
      Height          =   975
      Left            =   2040
      TabIndex        =   23
      Top             =   3480
      Width           =   1935
      Begin VB.OptionButton optChange 
         Caption         =   "Change settings"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optChange 
         Caption         =   "Revert to defaults"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame fmeDirection 
      Caption         =   "Direction"
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1815
      Begin VB.OptionButton optDirection 
         Caption         =   "Start at top"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Start at bottom"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "Randomise"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkDisregardLaptops 
      Caption         =   "Disregard Laptops"
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "Save Results in SQL"
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   4320
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CommandButton cmdGetWorkstations 
      Caption         =   "Get Workstations"
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Errors"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CheckBox chkAutomated 
      Caption         =   "Automated Siteloc determination"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   3600
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog cdlBrowse 
      Left            =   4440
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstTemp 
      Height          =   450
      Left            =   7200
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton btnDelAll 
      Caption         =   "Remove A&ll"
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame fmeNew 
      Caption         =   "New Settings"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   6615
      Begin VB.TextBox txtSiteServer 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Text            =   "Applications\Apps\NTPolicy\NTConfig.POL"
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label lblSiteServer 
         Caption         =   "Site Server"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPolicyPath 
         Caption         =   "Policy Path"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton btnDelTarget 
      Caption         =   "&Remove"
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton btnAddTarget 
      Caption         =   "&Add"
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtAddTarget 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
   End
   Begin VB.ListBox lstTargets 
      Height          =   2595
      ItemData        =   "frmMain.frx":030A
      Left            =   120
      List            =   "frmMain.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblSiteloc 
      Caption         =   "Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label lblProgress 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   7560
      Width           =   6975
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2160
      Picture         =   "frmMain.frx":030E
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   5400
      Picture         =   "frmMain.frx":25B0
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   7920
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblTargets 
      Alignment       =   2  'Center
      Caption         =   "Targets"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7080
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Deploy"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Bail"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExecuteMode As Boolean
Dim SiteServer As String
Dim SiteListFile As String

Sub LogLocal(Message As String)
    lblProgress = Message
    lblProgress.Refresh
    Open App.Path & "\RPolStatus" & App.ThreadID & ".txt" For Append As #3
        Print #3, Message
    Close 3
End Sub

Public Sub DoTargets()
    Dim i As Integer, Target As String, Data As Long, ECode As Boolean, j As Integer, msg As String
    Dim adoConn As ADODB.Connection
    Dim adoRs As ADODB.Recordset
    Set adoConn = New ADODB.Connection
    Set adoRs = New ADODB.Recordset
    Randomize Timer
    DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=" & "CBDXAAI"
    adoConn.Open DSN
    
    frmLogin.Show vbModal
    lstTemp.Clear
    
    LogLocal "----------------------------------------------"
    LogLocal "Starting Target List"
    If chkDisregardLaptops.Value = vbChecked Then LogLocal "* Skipping Laptops"
    If chkSave.Value = vbChecked Then LogLocal "* Saving to SQL"
    If chkAutomated.Value = vbChecked Then LogLocal "* Automated SITESERVER Determination ON"
    If optDirection(0).Value = True Then LogLocal "* Starting at top"
    If optDirection(1).Value = True Then LogLocal "* Starting at bottom"
    If optDirection(2).Value = True Then LogLocal "* Using random determination of targets"
    LogLocal "* Policy path: " & txtPath
    If optChange(0).Value = True Then
        LogLocal "* Modifying targets to use custom path"
    Else
        LogLocal "* Modifying targets to use default path"
    End If
    
    For i = 0 To lstTargets.ListCount - 1
redo:
        If optDirection(0).Value = True Then
            Target = lstTargets.List(i)
        End If
        If optDirection(1).Value = True Then
            Target = lstTargets.List(lstTargets.ListCount - i - 1)
        End If
        If optDirection(2).Value = True Then
            r = (Rnd(1) * lstTargets.ListCount - 1) + 1
            Target = lstTargets.List(r)
            'NextDateTime = adoRs.Fields("NextDateTime")
            'adoRs.Close
            'If NextDateTime <> Null Then
            '    Do Until Int(NextDateTime) < Int(Time)
            '        r = (Rnd(1) * lstTargets.ListCount - 1) + 1
            '        Target = lstTargets.List(r)
            '        SQL = "SELECT * FROM tblWorkstations WHERE Hostname = '" & Target & "'"
            '        adoRs.Open SQL, adoConn
            '        NextDateTime = adoRs.Fields("NextDateTime")
            '        Debug.Print NextDateTime
            '        adoRs.Close
            '    Loop
            'End If
        End If
        ' Check to ensure this target isn't already done
        SQL = "SELECT * FROM tblWorkstations WHERE Hostname = '" & Target & "'"
        adoRs.Open SQL, adoConn
        If Val(adoRs.Fields("isDone")) = "1" Then
            i = i + 1
            adoRs.Close
            GoTo redo
        End If
        adoRs.Close
        LogLocal Target & " " & Time
        DoEvents
        txtAddTarget.Text = Target + " ..."
        txtAddTarget.Refresh
        
        ResetModuleIOResult
        
        ' determine if Target is up
        dr = ""
        On Error Resume Next
            dr = Dir("\\" & Target & "\c$\*.*")
        On Error GoTo 0
        If dr <> "" Then GoTo continue
        lstErrors.AddItem Target & " unreachable"
        LogLocal Target & " unreachable"
        i = i + 1
        GoTo redo
continue:
        LogLocal Target & " (Reading Registry)"
        oldsiteloc = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Siteloc")
        OldSiteserver = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Siteserver")
        OldNetworkPath = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Update", "NetworkPath")
        OldUpdateMode = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Update", "UpdateMode")
        OldVerbose = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Update", "Verbose")
        DoEvents
        
        ' Get current siteloc
        siteloc = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Siteloc")
        
        ' Get client type
        client = ""
        client = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Client")
        
        If UCase(client) = "CPQLAPTOP" And chkDisregardLaptops.Value = vbChecked Then
            ' Leave it alone
            LogLocal Target & " **** Laptop found - skipping"
        Else
            If UCase(siteloc) = "SITELOC" Or siteloc = "" Then
                LogLocal Target & " **** Incorrect SITELOC"
                'lstSiteloc.AddItem Target & " (" & siteloc & ")"
            End If
            If chkAutomated.Value = vbChecked Then
                    LogLocal Target & " (Determining Siteserver)"
                    ' Work through siteloc's for this computer, and determine site server
                    j = 1
                    SiteServer = ""
                    doExit = False
                    If siteloc <> "" Then
                        Do Until CSVServer(j) = "" Or doExit = True
                            If UCase(siteloc) = UCase(CSVSiteloc(j)) Then
                                SiteServer = "\\" & CSVServer(j)
                                doExit = True
                            End If
                            j = j + 1
                        Loop
                    Else
                        LogLocal Target & " **** unable to determine SITESERVER"
                    End If
                End If
                LogLocal Target & " (Siteserver = " & SiteServer & ")"
                
                ' Change the target to use these new values.
                If optChange(0).Value = True And SiteServer <> "" Then
                    If (GetModuleIOResult = 0) Then
                        ' Set registry
                        LogLocal Target & " (Writing Registry to set new values)"
                        Retval = SetValueString(Target, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "SiteServer", SiteServer, REG_SZ)
                        Retval = SetValueString(Target, "SYSTEM\CurrentControlSet\Control\Update", "NetworkPath", "%SITESERVER%\" & txtPath, REG_SZ)
                        Retval = SetValueLong(Target, "SYSTEM\CurrentControlSet\Control\Update", "UpdateMode", 2, REG_DWORD)
                        Retval = SetValueLong(Target, "SYSTEM\CurrentControlSet\Control\Update", "Verbose", 0, REG_DWORD)
                        
                        ' Check registry
                        LogLocal Target & " (Reading Registry for confirmation)"
                        'NewSiteloc = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Siteloc")
                        NewSiteserver = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Siteserver")
                        NewNetworkPath = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Update", "NetworkPath")
                        NewupdateMode = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Update", "UpdateMode")
                        NewVerbose = GetValue(Target, "SYSTEM\CurrentControlSet\Control\Update", "Verbose")
                        DoEvents
                        
                        ' Check that changes have been applied
                        If (NewSiteserver <> SiteServer) Or (NewNetworkPath <> ("%SITESERVER%\" & txtPath)) Then
                            lstTemp.AddItem ("(ERR) " + lstTargets.List(i))
                            LogLocal Target & " (Error!)"
                            Open App.Path & "\RPolErrors.txt" For Append As #2
                                Print #2, Target
                            Close 2
                            If chkSave.Value = vbChecked Then
                                LogLocal Target & " (Updating SQL)"
                                SQL = "UPDATE tblWorkstations SET HasPolicyUpdated = '0' Where (Hostname = '" & Target & "')"
                                adoRs.Open SQL, adoConn
                                ' Done saving to SQL Database
                            End If
                        Else
                            Open App.Path & "\RPolLog.txt" For Append As #2
                                Print #2, Target
                            Close 2
                            If chkSave.Value = vbChecked Then
                                LogLocal Target & " (Updating SQL)"
                                SQL = "UPDATE tblWorkstations SET HasPolicyUpdated = '1' Where (Hostname = '" & Target & "')"
                                adoRs.Open SQL, adoConn
                                ' Done saving to SQL Database
                            End If
                        End If
                    Else
                        lstTemp.AddItem ("(ERR) " + lstTargets.List(i))
                    End If
            End If
            DoEvents
            
            ' Change the target to use default values, ie revert back to normal
            If optChange(1).Value = True Or SiteServer = "" Then
                msg = Target & " **** Unknown or blank Siteserver.  Siteloc = " & oldsiteloc
                lstErrors.AddItem msg
                LogLocal msg
                If GetModuleIOResult = 0 Then
                    LogLocal Target & " (Writing Registry to set defaults)"
                    Retval = SetValueString(Target, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "SiteServer", "", REG_SZ)
                    Retval = SetValueString(Target, "SYSTEM\CurrentControlSet\Control\Update", "NetworkPath", "", REG_SZ)
                    Retval = SetValueLong(Target, "SYSTEM\CurrentControlSet\Control\Update", "UpdateMode", 1, REG_DWORD)
                    Retval = SetValueLong(Target, "SYSTEM\CurrentControlSet\Control\Update", "Verbose", 0, REG_DWORD)
                    ' Update database
                    SQL = "UPDATE tblWorkstations SET tblWorkstations.HasPolicyUpdated = 0 Where (Hostname = '" & Target & "')"
                    adoRs.Open SQL, adoConn
                Else
                    lstTemp.AddItem ("(ERR) " + lstTargets.List(i))
                End If
            End If
            DoEvents
        End If
    Next
    lstTargets.Clear
    For i = 0 To lstTemp.ListCount - 1
        lstTargets.AddItem (lstTemp.List(i))
    Next
    txtAddTarget.Text = ""
    If lstTemp.ListCount > 0 Then
        cmdExport.Enabled = True
    Else
        cmdExport.Enabled = False
    End If
End Sub

Private Function GetValue(Hostname As String, Key As String, Value As String) As String
    Dim OpenKeyVal As Long
    Dim RResult As Long
    Dim OpenHiveVal As Long
    Dim strValue, CMPTRName, keytogo As String, InfoTextStr As String
    Dim x As Integer
    
    ECode = False
    
    GetIPCConnection (Trim(Hostname))
    If (GetModuleIOResult <> 0) Then Exit Function
    CMPTRName = Trim(Hostname)
    
    RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
    keytogo = Trim(Key)
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
    InfoTextStr = RegistryQueryValue(OpenKeyVal, Trim(Value), REG_SZ)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    DisIPCConnection (Trim(Hostname))
    
    Retval = Trim(InfoTextStr)
    GetValue = Retval
    ECode = True
End Function

Private Function SetValueString(Hostname As String, Key As String, Value As String, Data As String, regType As Long) As Long
    Dim OpenKeyVal As Long
    Dim RResult As Long
    Dim OpenHiveVal As Long
    Dim strValue, CMPTRName, keytogo As String, InfoTextStr As String
    Dim x As Integer
    
    GetIPCConnection (Trim(Hostname))
    
    CMPTRName = Trim(Hostname)
    
    RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
    keytogo = Trim(Key)
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
    RegistryWriteValue Data, OpenKeyVal, Value, regType
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    DisIPCConnection (Trim(Hostname))
        
End Function

Private Function SetValueLong(Hostname As String, Key As String, Value As String, Data As Long, regType As Long) As Long
    Dim OpenKeyVal As Long
    Dim RResult As Long
    Dim OpenHiveVal As Long
    Dim strValue, CMPTRName, keytogo As String, InfoTextStr As String
    Dim x As Integer
    
    GetIPCConnection (Trim(Hostname))
    
    CMPTRName = Trim(Hostname)
    
    RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
    keytogo = Trim(Key)
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
    RegSetValueEx OpenKeyVal, Value, 0&, REG_DWORD, Chr(Data), 4
    '                                               ^^^^^^^^^  = VERY Strange!

    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    DisIPCConnection (Trim(Hostname))
        
End Function

Private Sub btnAddTarget_Click()
  If (txtAddTarget.Text <> "") Then
    lstTargets.AddItem (txtAddTarget.Text)
  End If
  txtAddTarget.Text = ""
  txtAddTarget.SetFocus
End Sub

Private Sub btnDelAll_Click()
  lstTargets.Clear
End Sub

Private Sub btnDelTarget_Click()
  Dim i As Integer
  i = lstTargets.ListIndex
  If (i >= 0) Then
    lstTargets.RemoveItem (i)
  End If
End Sub

Sub Disable()
  btnAddTarget.Enabled = False
  btnDelTarget.Enabled = False
  btnDelAll.Enabled = False
  txtAddTarget.Enabled = False
  txtSiteServer.Enabled = False
  lblTargets.Enabled = False
  lblSiteServer.Enabled = False
  txtPath.Enabled = False
  lblPolicyPath.Enabled = False
  fmeNew.Enabled = False
End Sub

Sub Enable()
  btnAddTarget.Enabled = True
  btnDelTarget.Enabled = True
  btnDelAll.Enabled = True
  txtAddTarget.Enabled = True
  lblTargets.Enabled = True
  If optChange(0).Value = True Then ' Change
    txtSiteServer.Enabled = True
    lblSiteServer.Enabled = True
    txtPath.Enabled = True
    lblPolicyPath.Enabled = True
    fmeNew.Enabled = True
  End If
End Sub

Private Sub chkAutomated_Click()
    If chkAutomated.Value = vbChecked Then
        With cdlBrowse
            .CancelError = True
            .InitDir = "\\CBDXAAA\Netlogon\Sites"
            .Filter = "Comma Delimited Files (*.csv)|*.csv)"
            .FileName = "sitelist.csv"
            .DialogTitle = "Open Sitelist file"
            On Error Resume Next
            .ShowOpen
            On Error GoTo 0
            If Err.Number = 0 Then
                SiteListFile = cdlBrowse.FileName
                Open SiteListFile For Input As #1
                    i = 1
                    Do Until EOF(1)
                        Input #1, CSVSiteloc(i), CSVServer(i)
                        'Debug.Print CSVSiteloc(i)
                        'Debug.Print CSVServer(i)
                        i = i + 1
                    Loop
                Close 1
            End If
        End With
        txtSiteServer.Enabled = False
        lblSiteServer.Enabled = False
    Else
        txtSiteServer.Enabled = True
        lblSiteServer.Enabled = True
    End If
End Sub

Private Sub cmdExport_Click()
    Open App.Path & "\RPolErrors.txt" For Output As #1
        For lp = 0 To lstTemp.ListCount - 1
            Print #1, Right(lstTemp.List(lp), Len(lstTemp.List(lp)) - 6)
        Next lp
    Close 1
End Sub

Private Sub cmdGetWorkstations_Click()
    Dim adoConn As ADODB.Connection, DSN As String, SQL As String, adoData As ADODB.Recordset
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoData = CreateObject("ADODB.Recordset")
    
    lstTargets.Clear
    DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=CBDXAAI"
    adoConn.Open DSN
    SQL = ""
    SQL = "SELECT * FROM tblworkstations WHERE (isPing = 1) AND (HasPolicyUpdated = '0')"
    If chkDisregardLaptops.Value = vbChecked Then SQL = SQL & " AND (isLaptop = '0')"
    adoData.Open SQL, adoConn
    Do Until adoData.EOF
        lstTargets.AddItem adoData.Fields("Hostname")
        adoData.MoveNext
        'UpdateDestination
    Loop
    adoConn.Close
    Set adoConn = Nothing
    lblProgress = lstTargets.ListCount - 1 & " Targets found"
End Sub

Private Sub cmdImport_Click()
    With cdlBrowse
        .InitDir = App.Path
        .DefaultExt = ".txt"
        .CancelError = True
        .Filter = "Text Files (*.txt)|*.txt|Comma Delimited Files (*.csv)|*.csv|All Files (*.*)|*.*"
        .FileName = ""
        On Error Resume Next
            Do Until .FileName <> "" Or Err.Number <> 0
                .ShowOpen
            Loop
            If .FileName <> "" Then
                Open cdlBrowse.FileName For Input As #1
                    Do Until EOF(1)
                        Input #1, Data
                        If Data <> "" Then lstTargets.AddItem Data
                    Loop
                Close 1
            End If
        On Error GoTo 0
    End With
End Sub

Private Sub Form_Load()
  ExecuteMode = False
End Sub


Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuStart_Click()
    If txtSiteServer.Text = "" And optChange(0).Value = True And chkAutomated.Value = vbUnchecked Then
        MsgBox "No SiteServer entered."
        Exit Sub
    End If
    If (lstTargets.ListCount < 1) Then
        MsgBox "No Targets specified"
        Exit Sub
    End If
    SiteServer = txtSiteServer
    If Left(SiteServer, 2) <> "\\" Then
        SiteServer = "\\" & txtSiteServer.Text
    Else
        'SiteServer = txtSiteServer.Text
    End If
    Disable
    DoTargets
    Enable
    lblProgress = "Idle"
    lblProgress.Refresh
    msg = "Errors have been logged to " & App.Path & "\RPolErrors.log" & vbCrLf
    msg = msg & "Successes have been logged to " & App.Path & "\RPolLog.txt"
    MsgBox msg
End Sub

Private Sub optChange_Click(Index As Integer)
    Select Case Index
        Case 0
            fmeNew.Enabled = True
            txtSiteServer.Enabled = True
            lblSiteServer.Enabled = True
            txtPath.Enabled = True
            lblPolicyPath.Enabled = True
        Case 1
            fmeNew.Enabled = False
            txtSiteServer.Enabled = False
            lblSiteServer.Enabled = False
            txtPath.Enabled = False
            lblPolicyPath.Enabled = False
    End Select
End Sub

Private Sub txtSiteServer_KeyPress(KeyAscii As Integer)
    ok = False
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ' Ensure text field contains valid characters
    If (Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z") Then ok = True
    If (KeyAscii = 8) Then ok = True
    If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") Then ok = True
    If (Chr(KeyAscii)) = "_" Then ok = True
    
    If ok = True Then
        ' Ensure text field is no greater than 15 characters
        If Len(txtSiteServer) < 15 Then
            ok = True
        Else
            If KeyAscii = 8 Then ' Backspace key only key allowed now
                ok = True
            Else
                ok = False
            End If
        End If
    End If
    
    If ok = True Then
        KeyAscii = KeyAscii ' Everything OK
    Else
        KeyAscii = 0        ' Nup, don't let this char through
    End If
End Sub
