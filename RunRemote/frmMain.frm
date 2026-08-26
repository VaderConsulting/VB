VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run Remote"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   13620
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstProgress 
      Height          =   2790
      Left            =   7560
      TabIndex        =   8
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdGetServers 
      Caption         =   "Get List"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox lstServers 
      Columns         =   4
      Height          =   1035
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   7335
   End
   Begin VB.TextBox txtSource2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   6495
   End
   Begin VB.TextBox txtSource1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label lblInfo 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "Source 2"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblSource 
      Caption         =   "Source 1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetServers_Click()
    Dim strFilename As String
    Dim strData As String
    Dim strDefault As String
    
    lstServers.Clear
    
    If Right(App.Path, 1) <> "\" Then
        strDefault = App.Path & "\Servers.txt"
    Else
        strDefault = App.Path & "Servers.txt"
    End If
    
    strFilename = InputBox("Enter the filename to import", "Input", strDefault)
    
    If strFilename <> "" Then
        If Dir(strFilename) <> "" Then
            Open strFilename For Input As #1
                Do Until EOF(1)
                    Line Input #1, strData
                    If strData <> "" Then
                        lstServers.AddItem strData
                    End If
                Loop
            Close 1
        Else
            MsgBox "The filename entered was not found.", vbExclamation + vbOKOnly, "Problem"
        End If
    End If
    
End Sub

Private Sub cmdStart_Click()
    Dim strComputername As String
    Dim strUsername As String
    Dim strPassword As String
    Dim strSource1 As String
    Dim strSource2 As String
    Dim strDestination As String
    Dim strNewDir As String
    Dim i As Integer
    Dim bResult As Boolean
    
    lstProgress.Clear
        
    strSource1 = Me.txtSource1
    strSource2 = Me.txtSource2
    
    For i = 0 To lstServers.ListCount - 1
        strComputername = lstServers.List(i)
        
        lblInfo.Caption = strComputername
        Me.Refresh
        
        '1.  Create subdirectories
            On Error Resume Next
'                strNewDir = "\\" & strComputername & "\c$\temp"
'                lblInfo.Caption = "Creating Temp directory on " & strComputername
'                frmMain.Refresh
'                MkDir strNewDir
'                If Err.Number <> 0 Then
'                    UpdateProgress Err.Description & " creating Temp directory on " & strComputername
'                Else
'                    UpdateProgress "Created Temp directory on " & strComputername
'                End If
'                Err.Clear
'                strNewDir = "\\" & strComputername & "\c$\temp\ServerScripts"
'                lblInfo.Caption = "Creating ServerScripts directory on " & strComputername
'                frmMain.Refresh
'                MkDir strNewDir
'                If Err.Number <> 0 Then
'                    UpdateProgress Err.Description & " creating Server Scripts directory on " & strComputername
'                Else
'                    UpdateProgress "Created Server Scripts directory on " & strComputername
'                End If
'                Err.Clear
            On Error GoTo 0
        ' Copy files to destination server
'            '2.  MSI
'                lblInfo.Caption = "Copying MSI to " & strComputername
'                frmMain.Refresh
'                bResult = APIFileCopy(strSource1 & "\Setup.msi", strNewDir & "\setup.msi", False)
'                UpdateProgress strComputername & ": MSI Copy result - " & bResult
'            '3.  SetupPurge
'                lblInfo.Caption = "Copying SetupPurge to " & strComputername
'                frmMain.Refresh
'                bResult = APIFileCopy(strSource2 & "\SetupPurge.exe", strNewDir & "\SetupPurge.exe", False)
'                UpdateProgress strComputername & ": SetupPurge Copy result - " & bResult
'            '4.  SetupRegBack
'                lblInfo.Caption = "Copying SetupRegBack to " & strComputername
'                frmMain.Refresh
'                bResult = APIFileCopy(strSource2 & "\SetupRegBack.exe", strNewDir & "\SetupRegBack.exe", False)
'                UpdateProgress strComputername & ": SetupRegBack Copy result - " & bResult
'            '5.  SetupELBackups
'                lblInfo.Caption = "Copying SetupELBackups to " & strComputername
'                frmMain.Refresh
'                bResult = APIFileCopy(strSource2 & "\SetupELBackups.exe", strNewDir & "\SetupELBackups.exe", False)
'                UpdateProgress strComputername & ": SetupELBackups Copy result - " & bResult
'            '6.  Batch file
'                lblInfo.Caption = "Copying InstallSS to " & strComputername
'                frmMain.Refresh
'                bResult = APIFileCopy(strSource2 & "\InstallSS.cmd", strNewDir & "\InstallSS.cmd", False)
'                UpdateProgress strComputername & ": InstallSS Copy result - " & bResult
            '7.  MSI
                lblInfo.Caption = "Copying Scripts to " & strComputername
                frmMain.Refresh
                strNewDir = "\\" & strComputername & "\c$\winnt\scripts"
                bResult = APIFileCopy(strSource2 & "\backupeventlogs.vbe", strNewDir & "\backupeventlogs.vbe", False)
                UpdateProgress strComputername & ": Script 1 result - " & bResult
                bResult = APIFileCopy(strSource2 & "\backupregistry.vbe", strNewDir & "\backupregistry.vbe", False)
                UpdateProgress strComputername & ": Script 2 result - " & bResult
'           '8.  Create scheduled task
'            oSchedule.TargetComputer = "\\" & strComputername
'            oSchedule.Refresh
'            'On Error Resume Next
'                Set oJob = oSchedule.CreateTask("Server Scripts - Installation")
'                If Err.Number = 0 Then ' No problem, continue
'                    oJob.ApplicationName = "C:\Temp\ServerScripts\InstallSS.cmd"
'                    oJob.Creator = Environ$("USERNAME")
'                    oJob.CommandLine = ""
'                    oJob.WorkingDirectory = "C:\Temp\ServerScripts"
'                    Set oTrigger = oJob.Triggers.Add
'                    oTrigger.TriggerType = ttOnce
'                    oTrigger.BeginDay = Date
'                    oTrigger.StartTime = DateAdd("n", 1, Time)
'                    'oTrigger.Update
'                    oJob.Save
'                    oJob.SetAccountInfo strUsername, strPassword
'                    oJob.Save
'                    Set oTrigger = Nothing
'                    Set oJob = Nothing
'                End If
'            'On Error GoTo 0
        lstProgress.AddItem "------- " & strComputername & " complete. -------"
        bResult = False
    Next i
    lblInfo.Caption = "Idle"
    Me.Refresh
End Sub

Private Sub Form_Load()
    txtUsername = Environ$("USERNAME")
    txtSource1.Text = "\\wde.woodside.com.au\shares\WDE\Apps\Optional\Information Management and Technology\A0001244_SERVSCRIPT_V05"
    txtSource2.Text = "\\wde.woodside.com.au\shares\WDE\Admin\Startup and Logon Scripts"
End Sub

' InstallSS.cmd:
'  c:\temp\ServerScripts\SetupELBackups.exe
'  c:\temp\ServerScripts\SetupPurge.exe
'  c:\temp\ServerScripts\SetupRegBack.exe
'  msiexec /i "c:\temp\ServerScripts\Setup.msi" ALLUSERS=1 /qn

Private Sub UpdateProgress(strMessage As String)
    lstProgress.AddItem strMessage
    Open App.Path & "\Results.txt" For Append As #1
        Print #1, Time & " " & strMessage
    Close 1
    frmMain.Refresh
    DoEvents
End Sub
