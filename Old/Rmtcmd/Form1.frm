VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Scheduling"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmeFTP 
      Caption         =   "FTP Commands"
      Height          =   1455
      Left            =   5520
      TabIndex        =   29
      Top             =   1440
      Width           =   1935
      Begin VB.TextBox txtPeriod3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   32
         Text            =   "10"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtFTP 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Text            =   "c:\temp\ftpcopy.cmd"
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkRun 
         Caption         =   "Run"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CheckBox chkSchedule 
      Caption         =   "Start Schedule Service"
      Height          =   255
      Left            =   5520
      TabIndex        =   28
      Top             =   240
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Frame fmeCommand 
      Caption         =   "Remote Command"
      Height          =   735
      Left            =   5520
      TabIndex        =   26
      Top             =   600
      Width           =   1935
      Begin VB.CheckBox chkInstall 
         Caption         =   "Install"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CheckBox chkShare 
      Caption         =   "Share CD ROM Drive"
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   240
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Frame fmeImport 
      Caption         =   "Server List import"
      Height          =   1575
      Left            =   2880
      TabIndex        =   21
      Top             =   3720
      Width           =   2535
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   495
         Left            =   720
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Text            =   "c:\temp\servers.csv"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblPath 
         Alignment       =   2  'Center
         Caption         =   "Import path"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fmeServers 
      Caption         =   "Servers"
      Height          =   5175
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   20
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtAdd 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.ListBox lstServers 
         Height          =   3765
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame fmeOptions2 
      Caption         =   "SP6A Upgrade"
      Height          =   1455
      Left            =   2880
      TabIndex        =   10
      Top             =   2160
      Width           =   2535
      Begin VB.CheckBox chkInteractive2 
         Caption         =   "Interactive"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkSP 
         Caption         =   "Perform SP6A Upgrade"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtPeriod2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Text            =   "30"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Wait Period:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "minutes"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fmeOptions1 
      Caption         =   "CD Copy"
      Height          =   1455
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   2535
      Begin VB.CheckBox chkInteractive1 
         Caption         =   "Interactive"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkCopy 
         Caption         =   "Perform CD Copy"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.TextBox txtPeriod1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "3"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblWait1 
         Caption         =   "Wait Period:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "minutes"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fmeResults 
      Caption         =   "Results"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   7335
      Begin VB.ListBox lstResults 
         Height          =   1425
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7095
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   7095
      End
   End
   Begin VB.CommandButton cmdSchedule 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   7680
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkCopy_Click()
    If chkCopy.Value = vbChecked Then
        txtPeriod1.Enabled = True
        chkInteractive1.Enabled = True
    Else
        txtPeriod1.Enabled = False
        chkInteractive1.Enabled = False
    End If
End Sub

Private Sub chkShare_Click()
    If chkShare.Value = vbChecked Then
        fmeOptions1.Enabled = True
        fmeOptions2.Enabled = True
    Else
        fmeOptions1.Enabled = False
        fmeOptions2.Enabled = False
    End If
    
End Sub

Private Sub chkSP_Click()
    If chkSP.Value = vbChecked Then
        txtPeriod2.Enabled = True
        chkInteractive2.Enabled = True
    Else
        txtPeriod2.Enabled = False
        chkInteractive2.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    lstServers.AddItem txtAdd.Text
End Sub

Private Sub cmdImport_Click()
    lstServers.Clear
    Open txtPath For Input As #1
    Do Until EOF(1)
        Input #1, varLine
        'If InStr(1, varLine, "WA") <> 0 Then
            lstServers.AddItem varLine
        'End If
    Loop
    Close 1
    If lstServers.ListCount <> 0 Then cmdSchedule.Enabled = True
End Sub

Private Sub cmdRemove_Click()
    If lstServers.ListIndex = -1 Then Exit Sub
    lstServers.RemoveItem lstServers.ListIndex
End Sub

Private Sub cmdSchedule_Click()
    'lstServers.Enabled = False
    For lp = 0 To lstServers.ListCount - 1
        lstServers.ListIndex = lp
         varSuccess = "Service is pending start on"
        varDisabled = "Error code 1058"
              varOK = "Service is already running or in a paused/pending state."
        
        Dim t As Date
        Dim tt As Date
        varStat1 = ""
        varStat2 = ""
        varStat3 = ""
        varStat4 = ""
        varStat5 = ""
        varStat6 = ""
        varHost = lstServers.List(lstServers.ListIndex)
        
        If chkSchedule = vbChecked Then
            varFile = "C:\temp\startsvr.bat"
            If Dir("c:\temp\netsvc.txt") <> "" Then
                Kill "c:\temp\netsvc.txt"
            End If
            varcmd = "netsvc schedule \\" & varHost & " /start >c:\temp\netsvc.txt"
            Open varFile For Output As #1
                Print #1, varcmd
            Close 1
            Shell varFile, vbHide
            lblStatus = "Executing command on remote host"
            varStat1 = lblStatus
            lblStatus.Refresh
            Do Until Dir("c:\temp\netsvc.txt") <> ""
                DoEvents
                t = Time
                Do Until Time > DateAdd("s", 2, t)
                    DoEvents
                Loop
                Do Until FileLen("c:\temp\netsvc.txt") > 0
                    lblStatus = "Waiting for command to complete"
                    varStat2 = lblStatus
                    DoEvents
                Loop
            Loop
            varFile = "C:\temp\netsvc.txt"
            varRet = ""
            Open varFile For Input As #1
                Input #1, varRet
            Close 1
            lblStatus = ""
            If InStr(1, varRet, "1058") <> 0 Then lblStatus = "Schedule Service Disabled"
            If InStr(1, varRet, "already") <> 0 Then lblStatus = "Schedule Service was already running"
            If InStr(1, varRet, "start") <> 0 Then lblStatus = "Schedule Service has been started"
            If lblStatus = "" And chkCopy.Value = vbChecked Then
                lblStatus = "Unknown error - " & varRet
            End If
            If lblStatus = "" And chkCopy.Value = vbUnchecked Then
                lblStatus = varRet
            End If
            lblStatus.Refresh
            varStat3 = lblStatus
            lstResults.AddItem varHost & " - " & lblStatus
        End If
        
        On Error Resume Next
        If chkInstall = vbChecked Then
            lblStatus = "copying instsrv.exe"
            lstResults.AddItem varHost & " - " & lblStatus
            FileCopy "c:\reskit\instsrv.exe", "\\" & varHost & "\c$\winnt\system32\instsrv.exe"
            lblStatus = "copying rcmdsvc.exe"
            lstResults.AddItem varHost & " - " & lblStatus
            FileCopy "c:\winnt\system32\rcmdsvc.exe", "\\" & varHost & "\c$\winnt\system32\rcmdsvc.exe"
            lblStatus = "copying instsrv.bat"
            lstResults.AddItem varHost & " - " & lblStatus
            FileCopy "c:\inst.bat", "\\" & varHost & "\c$\temp\inst.bat"
            t = Time
            tt = DateAdd("n", 10, t)
            lstResults.AddItem varHost & " - " & lblStatus
            varcmd = "at.exe \\" & varHost & " " & tt & " " & "c:\temp\inst.bat"
            varFile = "c:\temp\rcmdi.bat"
            Open varFile For Output As #1
                Print #1, varcmd
            Close 1
            Shell varFile, vbHide
            lblStatus = "Scheduled " & "c:\temp\inst.bat to run at " & tt
            lblStatus.Refresh
            varStat5 = lblStatus
            lstResults.AddItem varHost & " - " & lblStatus
            Form1.Refresh
        End If
        
        If chkRun = vbChecked Then
            t = Time
            tt = DateAdd("n", Val(txtPeriod3), t)
            lstResults.AddItem varHost & " - " & lblStatus
            varcmd = "at.exe \\" & varHost & " " & tt & " " & txtFTP
            varFile = "c:\temp\ftpcopy.bat"
            Open varFile For Output As #1
                Print #1, varcmd
            Close 1
            Shell varFile, vbHide
            lblStatus = "Scheduled " & txtFTP & " to run at " & tt
            lblStatus.Refresh
            varStat5 = lblStatus
            lstResults.AddItem varHost & " - " & lblStatus
        End If
        
        
        If chkShare.Value = vbChecked Then
            If lblStatus <> "Schedule Service Disabled" And Left(lblStatus, 7) <> "Unknown" Then
                varcmd = "rmtshare \\" & varHost & "\CDROM=d:\"
                Shell varcmd, vbHide
                varcmd = "\\" & varHost & "\CDROM\cdcopy.exe"
                e = ""
                On Error Resume Next
                    varResult = Dir(varcmd)
                    e = Error
                On Error GoTo 0
                If e = "" Then
                    If varResult <> "" Then
                        If chkCopy.Value = vbChecked Then
                            lblStatus = "Found file on shared CD ROM Drive"
                            lblStatus.Refresh
                            varStat4 = lblStatus
                            t = Time
                            tt = DateAdd("s", Val(txtPeriod1) * 60, t)
                            lstResults.AddItem varHost & " - " & lblStatus
                            If chkInteractive1.Value = vbChecked Then
                                varInter1 = "/interactive"
                            Else
                                varInter1 = ""
                            End If
                            varcmd = "at.exe \\" & varHost & " " & tt & " " & varInter1 & " d:\cdcopy.exe"
                            varFile = "c:\temp\cdcopy.bat"
                            Open varFile For Output As #1
                                Print #1, varcmd
                            Close 1
                            Shell varFile, vbHide
                            lblStatus = "Scheduled D:\CDCopy.exe to run at " & tt
                            lblStatus.Refresh
                            varStat5 = lblStatus
                            lstResults.AddItem varHost & " - " & lblStatus
                        End If
                        ' *************************************************************
                        If chkSP.Value = vbChecked Then
                            ' This is the portion that executes SP6a,
                            ' BUT will only execute when the user next logs on!!!
                            '
                            t = Time
                            tt = DateAdd("s", Val(txtPeriod2) * 60, t)
                            If chkInteractive2.Value = vbChecked Then
                                varInter2 = "/interactive"
                            Else
                                varInter2 = ""
                            End If
                            varcmd = "at.exe \\" & varHost & " " & tt & " " & varInter2 & " e:\wapssoe\updates\wep1\wep1stub.exe"
                            varFile = "c:\temp\sp6.bat"
                            Open varFile For Output As #1
                                Print #1, varcmd
                            Close 1
                            Shell varFile, vbHide
                            lblStatus = "Scheduled SP6a to run at " & tt
                            lblStatus.Refresh
                            varStat6 = lblStatus
                            lstResults.AddItem varHost & " - " & lblStatus
                        End If
                    Else
                        lblStatus = "File not found"
                        varStat4 = lblStatus
                    End If
                Else
                    lblStatus = "File not found - CDROM Drive empty"
                    lblStatus.Refresh
                    varStat4 = lblStatus
                    lstResults.AddItem varHost & " - " & lblStatus
                End If
            End If
        End If
        Open "c:\temp\copystat.csv" For Append As #2
            Print #2, varHost & "," & varStat1 & "," & varStat2 & "," & varStat3 & "," & varStat4 & "," & varStat5 & "," & varStat6
        Close 2
        lstResults.ListIndex = lstResults.ListCount - 1
    Next lp
    lstServers.Enabled = True
End Sub

Private Sub Form_Load()
    'lstServers.AddItem "CSC0100"
    'lstServers.AddItem "satsitexawa"
    'lstServers.AddItem "brooktonxawa"
    varFile = "C:\temp\copystat.csv"
    If Dir(varFile) <> "" Then
        Kill varFile
    End If
End Sub

Private Sub lstServers_Click()
    If lstServers.ListIndex <> -1 Then
        cmdSchedule.Enabled = True
        cmdRemove.Enabled = True
    Else
        cmdSchedule.Enabled = False
        cmdRemove.Enabled = False
    End If
End Sub

Private Sub txtAdd_Change()
    If txtAdd.Text <> "" Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
End Sub

