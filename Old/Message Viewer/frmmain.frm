VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon Status Message Viewer - (c) 2000 CSC + D. Robinson"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12225
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12225
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLogonStatus 
      Caption         =   "View Logonstatus.txt"
      Height          =   255
      Left            =   12000
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox rtbStatus 
      Height          =   4095
      Left            =   6240
      TabIndex        =   2
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7223
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0442
   End
   Begin RichTextLib.RichTextBox rtbKix 
      Height          =   8085
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   14261
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":051C
   End
   Begin VB.FileListBox filKix 
      Height          =   8085
      Left            =   120
      Pattern         =   "*.log"
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin MSChart20Lib.MSChart chtLogs 
      Height          =   3855
      Left            =   6240
      OleObjectBlob   =   "frmMain.frx":05F6
      TabIndex        =   7
      Top             =   4680
      Width           =   5895
   End
   Begin VB.Label lblEntries 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   8640
      Width           =   2175
   End
   Begin VB.Label lblUnable 
      Alignment       =   2  'Center
      Caption         =   "Unable to chart less than 2 entries"
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
      Left            =   6240
      TabIndex        =   9
      Top             =   6360
      Width           =   5895
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   12240
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblLogonStatus 
      Alignment       =   2  'Center
      Caption         =   "Logonstatus.txt (on remote computer)"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblServer 
      Alignment       =   2  'Center
      Caption         =   "Status file contents"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblFiles 
      Alignment       =   2  'Center
      Caption         =   "Computer names"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   8640
      Width           =   9735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuLogonStatus 
         Caption         =   "View LogonStatus.txt"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuExtract 
         Caption         =   "Extract to .csv"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Base 1
    Dim WANAccessTime(3000, 1000) As Integer
    Dim ProcessingTime(3000, 1000) As Integer
    Dim Workstations As Long
    Dim Username(3000) As String
    Dim Sitename(3000) As String
    Dim Server(3000) As String
    Dim Entries(3000) As Integer

Sub Setup()
    'filKix.Path = "\\CBDXAAI\Kixlog$"
    On Error Resume Next
    filKix.Path = LogDir
    On Error GoTo 0
    If Err <> 0 Then
        MsgBox "Error accessing directory."
        Exit Sub
    End If
    Logs = Dir(filKix.Path)
    While Logs <> ""
        filKix.AddItem Logs
        Logs = Dir()
    Wend
    If filKix.ListCount > 0 Then
        For lst = 0 To filKix.ListCount - 1
            logfilename = filKix.List(lst)
            Workstations = 1
            Open filKix.Path & "\" & logfilename For Input As #1
                While Not EOF(1)
                    On Error Resume Next                 ' required to get around files still open
                    Line Input #1, StartTime
                    Line Input #1, WANCompleteTime
                    Line Input #1, FinishTime
                    Line Input #1, user
                    Line Input #1, site
                    Line Input #1, LogonServer
                    Line Input #1, IPAddress
                    Line Input #1, LogonDate
                    
                    StartTime = Mid(StartTime, 14, Len(StartTime) - 13)
                    WANCompleteTime = Mid(WANCompleteTime, 22, Len(WANCompleteTime) - 21)
                    FinishTime = Mid(FinishTime, 15, Len(FinishTime) - 14)
                    
                    'If DateDiff("s", StartTime, WANCompleteTime) > 0 And DateDiff("s", WANCompleteTime, FinishTime) > 0 Then
                        WANAccessTime(lst + 1, Workstations) = Abs(DateDiff("s", StartTime, WANCompleteTime))
                        ProcessingTime(lst + 1, Workstations) = Abs(DateDiff("s", WANCompleteTime, FinishTime))
                        Username(lst + 1) = Mid(user, 15, Len(user) - 12)
                        Sitename(lst + 1) = Mid(site, 15, Len(site) - 9)
                        Server(lst + 1) = Mid(LogonServer, 15, Len(LogonServer) - 9)
                        Workstations = Workstations + 1
                        If Workstations > 1000 Then Close (1)     ' Max of 1000 entries per workstation
                    'End If
                Wend
            Entries(lst + 1) = Workstations
            Close 1
        Next lst
    End If
    lblEntries = filKix.ListCount - 1 & " entries"
    lblEntries.Refresh
End Sub

Private Sub chkLogonStatus_Click()
    rtbStatus.Text = ""
End Sub


Private Sub filKix_Click()
    Dim strLocalName As String
    Dim NetR As NETRESOURCE
    Dim ErrInfo As Long
    Dim MyPass As String, MyUser As String
    
    Screen.MousePointer = vbHourglass
    lblStatus = "Accessing Log file on server"
    On Error Resume Next    ' required to remove file open errors
        rtbKix.FileName = filKix.Path & "\" & filKix.List(filKix.ListIndex)
    On Error GoTo 0
    rtbKix.Refresh
    computername = Left(filKix.List(filKix.ListIndex), Len(filKix.List(filKix.ListIndex)) - 4)
    If chkLogonStatus = vbChecked Then
        lblStatus = "Removing Drive Z:"
        lblStatus.Refresh
        
        If Dir("z:\msdos.sys") <> "" Then
            strLocalName = "Z:"
            ErrInfo = WNetCancelConnection2(strLocalName, CONNECT_UPDATE_PROFILE, False)
            If ErrInfo = NO_ERROR Then
                'MsgBox "Net Disconnection Successful!", vbInformation, "Share Disconnected"
            Else
                MsgBox "ERROR: " & ErrInfo & " - Net Disconnection Failed!", vbExclamation, "Share not Disconnected"
            End If
        End If
    End If
    MyUser = computername & "\administrator"
    MyPass = "W1gGle$"

    NetR.dwScope = RESOURCE_GLOBALNET
    NetR.dwType = RESOURCETYPE_DISK
    NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
    NetR.lpLocalName = "Z:"
    NetR.lpRemoteName = "\\" & computername & "\c$"
    
    If chkLogonStatus = vbChecked Then
        ErrInfo = WNetAddConnection2(NetR, MyPass, MyUser, CONNECT_UPDATE_PROFILE)
        If ErrInfo = NO_ERROR Then
            'MsgBox "Net Connection Successful!", vbInformation, "Share Connected"
            rtbStatus.Visible = True
            lblLogonStatus.Caption = "LogonStatus.txt"
        Else
            MsgBox "ERROR: " & ErrInfo & " - Net Connection Failed!", vbExclamation, "Share not Connected"
            rtbStatus.Visible = False
            lblLogonStatus.Caption = "LogonStatus.txt not available."
        End If
    
        
        lblStatus = "Accessing Z:\temp\Logonstatus.txt (\\" & computername & "\c$\temp\logonstatus.txt)"
        lblStatus.Refresh
        
        er = Dir("z:\temp\logonstatus.txt")
        If er <> "" Then
            FileCopy "z:\temp\logonstatus.txt", "c:\temp\otherstatus.txt"
            rtbStatus.FileName = "c:\temp\otherstatus.txt"
        Else
            MsgBox "This file cannot be retrieved.  Maybe the computer is powered off, or you do not have permission to the c$ share on this computer.", vbCritical + vbOKOnly, "Problem accessing file"
        End If
        ChDir "c:\"
    End If
    Screen.MousePointer = vbDefault
    lblStatus = "Ready"
    
    ' This section graphs the results
    chtLogs.ColumnCount = 2    ' Columns is WAN or LAN
    chtLogs.RowCount = Entries(filKix.ListIndex + 1) - 1
    If chtLogs.RowCount < 2 Then
        chtLogs.Visible = False
    Else
        chtLogs.Visible = True
    End If
    For col = 1 To 2
        chtLogs.Column = col
        If col = 1 Then chtLogs.ColumnLabel = "WAN Access Time"
        If col = 2 Then chtLogs.ColumnLabel = "Processing Time"
        For rw = 1 To chtLogs.RowCount
            chtLogs.Row = rw
            If col = 1 Then chtLogs.Data = WANAccessTime(filKix.ListIndex + 1, rw)
            If col = 2 Then chtLogs.Data = ProcessingTime(filKix.ListIndex + 1, rw)
        Next rw
    Next col
    chtLogs.Repaint = True
    chtLogs.Refresh
End Sub

Private Sub Form_Load()
    LogDir = GetSetting("Message Viewer", "Options", "LogDir", "\\CBDXAAI\Kixlog$")
    Me.Show
    lblStatus = "Retrieving list of computers... Please wait."
    Me.Refresh
    Setup
    lblStatus = "Ready"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    strLocalName = "Z:"
    ErrInfo = WNetCancelConnection2(strLocalName, CONNECT_UPDATE_PROFILE, False)
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuExtract_Click()
    lblStatus = "Extracting data to .csv file"
    eFlag = 0
    Open "c:\temp\PerfLogK.csv" For Output As #1
        For lst = 1 To filKix.ListCount
            For ent = 1 To Entries(lst)
                lne = Left(filKix.List(lst - 1), Len(filKix.List(lst - 1)) - 4) & ","
                lne = lne & Username(lst) & ","
                lne = lne & Sitename(lst) & ","
                lne = lne & Server(lst) & ","
                lne = lne & WANAccessTime(lst, ent) & ","
                lne = lne & ProcessingTime(lst, ent)
                
                If InStr(lne, ":") <> 0 Then
                    eFlag = 1
                Else
                    If Val(WANAccessTime(lst, ent)) <> 0 And Val(WANAccessTime(lst, ent)) > 0 And Val(ProcessingTime(lst, ent)) > 0 And Val(ProcessingTime(lst, ent)) > 0 Then
                        Print #1, lne
                    End If
                End If
            Next ent
        Next lst
    Close 1
    If eFlag <> 0 Then
        MsgBox "Problem found with data.  This dataset may not be consistent."
    End If
    MsgBox "Extract complete and saved as C:\temp\PerfLogK.csv"
    lblStatus = "Ready"
End Sub

Private Sub mnuLogonStatus_Click()
    If chkLogonStatus.Value = vbChecked Then
        chkLogonStatus.Value = vbUnchecked
    Else
        chkLogonStatus.Value = vbChecked
    End If
    mnuLogonStatus.Checked = chkLogonStatus.Value
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuRefresh_Click()
    Setup
End Sub
