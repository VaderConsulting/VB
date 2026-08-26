VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rolling Ping"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstFiles 
      Height          =   840
      Left            =   4440
      TabIndex        =   21
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox txtFind 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6720
      TabIndex        =   20
      Text            =   "2"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CheckBox chkCopy 
      Caption         =   "Copy files"
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Text            =   "c:\temp"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtEnd1 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "10"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtEnd2 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Text            =   "40"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtEnd3 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Text            =   "240"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtEnd4 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Text            =   "224"
      Top             =   600
      Width           =   495
   End
   Begin VB.ListBox lstIP 
      Height          =   3375
      Left            =   480
      TabIndex        =   12
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtIP4 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtIP3 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtIP2 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtIP1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "10"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblCopy 
      Alignment       =   2  'Center
      Caption         =   "Files to copy:"
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblNumber 
      Caption         =   "Number of PC's to find per subnet"
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5400
      Width           =   7095
   End
   Begin VB.Label lblPath 
      Caption         =   "Output path"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblDevices 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      Caption         =   "End"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExtract_Click()
    lblStatus = "Extracting..."
    lblStatus.Refresh
    Open txtOutput & "\ipaddresses.txt" For Output As #1
        For a = 0 To lstIP.ListCount - 1
            Print #1, lstIP.List(a)
        Next a
    Close 1
    lblStatus = "Ready"
    lblStatus.Refresh
End Sub

Private Sub cmdStart_Click()
    gStop = 0
    cmdStart.Visible = False
    cmdStop.Visible = True
    Dim t As Date
    Dim tt As Date
    'cmdStart.Enabled = False
    cmdExtract.Enabled = False
    lstIP.Clear
    Dim fMatch As Boolean, sRTT As String, sHost As String
    startIP1 = txtIP1
    startIP2 = txtIP2
    startIP3 = txtIP3
    startIP4 = txtIP4
    ip1 = startIP1
    varCopied = 0
    Do
        ip2 = startIP2
        Do
            ip3 = startIP3
            Do
                ip4 = startIP4
                Do
                    DoEvents
                    txtIP1 = ip1
                    txtIP2 = ip2
                    txtIP3 = ip3
                    txtIP4 = ip4
                    sHost = ip1 & "." & ip2 & "." & ip3 & "." & ip4
                    fMatch = True
                    lblStatus = "Pinging " & sHost
                    lblStatus.Refresh
                    If Ping(sHost, sRTT, fMatch) Then
                        lblStatus = "No reply from " & sHost
                        lblStatus.Refresh
                        DoEvents
                        txtIP1.Refresh
                        txtIP2.Refresh
                        txtIP3.Refresh
                        txtIP4.Refresh
                        'If fMatch Then
                        'Else
                        'End If
                    Else
                        ' Found a host
                        lblStatus = "Found a host"
                        lblStatus.Refresh
                        varDir = txtOutput & "\mifs\" & sHost
                        varSource1 = "\\" & sHost & "\c$\winnt\ms\sms\noidmifs\compaq0.mif"
                        varSource2 = "\\" & sHost & "\c$\winnt\ms\sms\noidmifs\compaq2.mif"
                        varSource3 = "\\" & sHost & "\c$\temp\PC.mif"
                        varSource4 = "\\" & sHost & "\c$\temp\Monitor.mif"
                        varSource5 = "\\" & sHost & "\c$\temp\*.*"
                        lblStatus = "Checking for PC or Device"
                        lblStatus.Refresh
                        varDir5 = ""
                        On Error Resume Next
                            varDir5 = Dir(varSource5)
                        On Error GoTo 0
                        DoEvents
                        If varDir5 <> "" Then ' found a computer
                            lblStatus = "Found a computer"
                            lblStatus.Refresh
                            varDir1 = Dir(varSource1)
                            DoEvents
                            varDir2 = Dir(varSource2)
                            DoEvents
                            varDir3 = Dir(varSource3)
                            DoEvents
                            varDir4 = Dir(varSource4)
                            DoEvents
                            MkDir txtOutput & "\mifs"
                            MkDir varDir
                            varDestination = txtOutput & "\mifs\" & sHost & "\*.*"
                            lblStatus = "Retrieving files"
                            lblStatus.Refresh
                            FileCopy varSource1, txtOutput & "\mifs\" & sHost & "\compaq0.mif"
                            DoEvents
                            FileCopy varSource2, txtOutput & "\mifs\" & sHost & "\compaq2.mif"
                            DoEvents
                            FileCopy varSource3, txtOutput & "\mifs\" & sHost & "\PC.mif"
                            DoEvents
                            FileCopy varSource4, txtOutput & "\mifs\" & sHost & "\Monitor.mif"
                            On Error GoTo 0
                            DoEvents
                            If chkCopy.Value = vbChecked Then
                                If varCopied < Val(txtFind) + 1 Then
                                    lblStatus = "Copying files to " & sHost
                                    lblStatus.Refresh
                                    On Error Resume Next
                                        FileCopy lstFiles.List(0), "\\" & sHost & "\c$\temp\wake.exe"
                                        lblStatus = "Wake.exe copied"
                                        lblStatus.Refresh
                                        FileCopy lstFiles.List(1), "\\" & sHost & "\c$\temp\dhcp client info.exe"
                                        lblStatus = "DHCP Client info copied"
                                        lblStatus.Refresh
                                        varCopied = varCopied + 1
                                    On Error GoTo 0
                                End If
                            End If
                        End If
                        sIP = sHost
                        If varDir1 <> "" Or varDir2 <> "" Or varDir3 <> "" Or varDir4 <> "" Then
                            sHost = sHost & ",with files"
                        Else
                            
                        End If
                        varCmd = "nbtstat -A " & sHost & " >c:\temp\nbtstat.txt"
                        lblStatus = "Creating NBTSTAT"
                        lblStatus.Refresh
                        Open "c:\temp\nbtinfo.bat" For Output As #1
                            Print #1, varCmd
                        Close 1
                        DoEvents
                        lblStatus = "Running NBTSTAT"
                        lblStatus.Refresh
                        Shell "c:\temp\nbtinfo.bat", vbHide
                        t = Time
                        If varDir5 <> "" Then
                            duration = 10 ' this is a pc.. give it more time
                        Else
                            duration = 1  ' Not a pc.. less time
                        End If
                        
                        tt = DateAdd("s", duration, t)
                                                
                        Do Until Time > tt
                            DoEvents
                        Loop
                        lblStatus = "Extracting NBTSTAT info"
                        lblStatus.Refresh
                        Open "c:\temp\nbtstat.txt" For Input As #1
                            Do Until EOF(1)
                                Line Input #1, varData
                                If InStr(varData, "MAC Address = ") > 0 Then
                                    Mac = Right(varData, 17)
                                    varOut = varDir & "\" & Mac
                                    If varDir5 <> "" Then
                                        Open varOut For Output As #4
                                        Close 4
                                    End If
                                End If
                            Loop
                        Close 1
                        lstIP.AddItem sHost & "," & Mac
                        Mac = ""
                    End If
                    DoEvents
                    lblDevices = lstIP.ListCount & " Devices"
                    frmMain.Refresh
                    ip4 = ip4 + 1
                    If ip4 > txtEnd4 Then
                        ip4 = 1
                        ip3 = ip3 + 1
                    End If
                    If varDir5 <> "" And chkCopy.Value = vbChecked And varCopied > Val(txtFind) Then
                        ip4 = txtEnd4
                        varCopied = 0
                    End If
                Loop Until gStop = 1
                ip4 = 1
                ip3 = ip3 + 1
            Loop Until gStop = 1
            ip3 = 1
            ip2 = ip2 + 1
        Loop Until gStop = 1
        ip2 = 1
        ip1 = ip1 + 1
    Loop Until (ip1 > txtEnd1 And ip2 > txtEnd2 And ip3 > txtEnd3 And ip4 > txtEnd4) Or gStop = 1
    cmdExtract.Enabled = True
    cmdStop.Visible = False
    cmdStart.Visible = True
End Sub

Private Sub cmdStop_Click()
    gStop = 1
    cmdStart.Visible = True
    cmdStop.Visible = False
End Sub

Private Sub Form_Load()
    PING_TIMEOUT = 200
    gStop = 0
    lstFiles.AddItem "f:\tools\wake.exe"
    lstFiles.AddItem "f:\tools\dhcp client info.csv"
End Sub

