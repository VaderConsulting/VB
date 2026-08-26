VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   4545
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Text            =   "c:\temp"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtEnd1 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Text            =   "10"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtEnd2 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Text            =   "40"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtEnd3 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Text            =   "240"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtEnd4 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Text            =   "224"
      Top             =   600
      Width           =   495
   End
   Begin VB.ListBox lstIP 
      Height          =   3960
      Left            =   1080
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblPath 
      Caption         =   "Output path"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblDevices 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      Caption         =   "End"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   10
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
    Open txtOutput & "\ipaddresses.txt" For Output As #1
        For a = 0 To lstIP.ListCount - 1
            Print #1, lstIP.List(a)
        Next a
    Close 1
    
End Sub

Private Sub cmdStart_Click()
    Dim t As Date
    Dim tt As Date
    cmdStart.Enabled = False
    cmdExtract.Enabled = False
    lstIP.Clear
    Dim fMatch As Boolean, sRTT As String, sHost As String
    startIP1 = txtIP1
    startIP2 = txtIP2
    startIP3 = txtIP3
    startIP4 = txtIP4
    For ip1 = startIP1 To txtEnd1
        For ip2 = startIP2 To txtEnd2
            For ip3 = startIP3 To txtEnd3
                For ip4 = startIP4 To txtEnd4
                    DoEvents
                    txtIP1 = ip1
                    txtIP2 = ip2
                    txtIP3 = ip3
                    txtIP4 = ip4
                    sHost = ip1 & "." & ip2 & "." & ip3 & "." & ip4
                    fMatch = True
                    If Ping(sHost, sRTT, fMatch) Then
                        If txtIP4 = "1" Then
                            txtIP3 = txtIP3 + 1
                            If txtIP3 > txtEnd2 Then
                                txtIP3 = txtIP2
                                txtIP2 = txtIP2 + 1
                                If txtIP2 > txtEnd1 Then
                                    txtIP2 = txtIP1
                                    txtIP1 = txtIP1 + 1
                                    If txtIP1 > txtEnd1 Then txtIP1 = txtEnd1
                                End If
                            End If
                        End If
                        DoEvents
                        txtIP1.Refresh
                        txtIP2.Refresh
                        txtIP3.Refresh
                        txtIP4.Refresh
                        If fMatch Then
                        Else
                        End If
                    Else
                        ' Found a host
                        varDir = txtOutput & "\mifs\" & sHost
                        varSource1 = "\\" & sHost & "\c$\winnt\ms\sms\noidmifs\compaq0.mif"
                        varSource2 = "\\" & sHost & "\c$\winnt\ms\sms\noidmifs\compaq2.mif"
                        varSource3 = "\\" & sHost & "\c$\temp\PC.mif"
                        varSource4 = "\\" & sHost & "\c$\temp\Monitor.mif"
                        varSource5 = "\\" & sHost & "\c$\temp\*.*"
                        On Error Resume Next
                        varDir1 = Dir(varSource1)
                        varDir2 = Dir(varSource2)
                        varDir3 = Dir(varSource3)
                        varDir4 = Dir(varSource4)
                        varDir5 = Dir(varSource5)
                        On Error GoTo 0
                        If varDir1 <> "" Or varDir2 <> "" Or varDir3 <> "" Or varDir4 <> "" Then
                            On Error Resume Next
                                MkDir txtOutput & "\mifs"
                                MkDir varDir
                                varDestination = txtOutput & "\mifs\" & sHost & "\*.*"
                                'FileCopy varSource1, txtOutput & "\mifs\" & sHost & "\compaq0.mif"
                                'FileCopy varSource2, txtOutput & "\mifs\" & sHost & "\compaq2.mif"
                                'FileCopy varSource3, txtOutput & "\mifs\" & sHost & "\PC.mif"
                                'FileCopy varSource4, txtOutput & "\mifs\" & sHost & "\Monitor.mif"
                            On Error GoTo 0
                            DoEvents
                        End If
                        sIP = sHost
                        If varDir1 <> "" Or varDir2 <> "" Or varDir3 <> "" Or varDir4 <> "" Then
                            sHost = sHost & ",with files"
                        Else
                            
                        End If
                        varCmd = "nbtstat -A " & sHost & " >c:\temp\nbtstat.txt"
                        Open "c:\temp\nbtinfo.bat" For Output As #1
                            Print #1, varCmd
                        Close 1
                        DoEvents
                        Shell "c:\temp\nbtinfo.bat", vbHide
                        t = Time
                        tt = DateAdd("s", 6, t)
                        Do Until tt > t
                            DoEvents
                        Loop
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
                Next ip4
            Next ip3
        Next ip2
    Next ip1
    cmdExtract.Enabled = True
End Sub

Private Sub Form_Load()
    PING_TIMEOUT = 200
End Sub
