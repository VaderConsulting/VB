VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon 3.2 Background Process"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7920
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstUsers 
      Height          =   2985
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      Width           =   7695
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   7695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   10680
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblUserNo 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblUsers 
      Alignment       =   2  'Center
      Caption         =   "Domain Users"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strDomain As String, strGroup As String, strUserDir As String, Paused As Boolean
Public TotalUsers As Integer, ServerNo As Integer, strPDCDir As String

Private Sub Form_Load()
    Dim commandLine As Boolean, msg As String
    lblStatus = "Initialising..."
    Paused = False
    Me.Show
    Me.Refresh
    ' If there is ANY commandline option specified, then the process will run only once
    If Command$ <> "" Then commandLine = True
    msg = "Started."
    'lstResults.AddItem msg
    LogLocal msg
doSAMLoop:
    StartTime = Now
    CheckSAM
    FinishTime = Now
    ElapsedTime = DateDiff("s", StartTime, FinishTime)
    msg = lblUserNo & " of " & TotalUsers & " accounts parsed. " & lblChangeNo & " Changes, " & lblNewNo & " new."
    If ElapsedTime <> 0 Then
        msg2 = "(" & Int((TotalUsers / ElapsedTime) * 10) / 10 & " users per second)"
    Else
        msg2 = ""
    End If
    'lstResults.AddItem msg & " " & msg2
    LogLocal msg & " " & msg2
    
    ' Reset variables
    lstUsers.Clear
    'lstGroups.Clear
    'lstNew.Clear
    'lstChanges.Clear
    lblUserNo = "0"
    'lblNewNo = "0"
    'lblChangeNo = "0"
    msg = ""
    msg2 = ""
    
    'lstResults.ListIndex = lstResults.ListCount - 1
    frmMain.Refresh
    'If commandLine = False Then
    '    GoTo doSAMLoop
    'Else
    '    mnuExit_Click
    'End If
End Sub

Sub LogLocal(msg As String)
    Open App.Path & "\appLog" & Format(Date, "DD-MM-YY") & ".log" For Append As #1
        Print #1, Format(Date, "DD/MM/YYYY") & " " & Format(Time, "Long Time") & " " & msg
    Close 1
End Sub

Sub CheckSAM()
    Dim msg As String
    strUserDir = "C:\TEMP\USERS"
    strPDCDir = "\\CBDXAAA\USERS"
    strDomain = "CBDXAAR"
    Set Domain = GetObject("WinNT://" & strDomain & ",computer")
    Domain.Filter = Array("user")
    If Err.Number <> 0 Then
        msg = "Error contacting " & strDomain
        'lstResults.AddItem msg
        'LogLocal msg
        ' This is a critical error, so quit at this point
        End
    End If
    
    TotalUsers = 0
    For Each member In Domain
        lblStatus = "Retrieving " & member.Name
        lblStatus.Refresh
        ssid = ""
        s = ""
        'If Member.LoginScript <> "" Then
        '    lstGroups.Clear
            UserName = UCase(member.Name)
            'lstUsers.AddItem UserName & " " & member.Guid
            SID = member.Get("objectSID")
            For i = LBound(SID) To UBound(SID)
              s = s & SID(i)
            Next
            ssid = Left(s, Len(s) - 1)
            ssid = "S-" & Left(ssid, 2) & Right(ssid, Len(ssid) - 7)
            ssid = Left(ssid, 3) & "-" & Right(ssid, Len(ssid) - 3)
            ssid = Left(ssid, 5) & "-" & Right(ssid, Len(ssid) - 5)
            ssid = Left(ssid, 8) & "-" & Right(ssid, Len(ssid) - 8)
            ssid = Left(ssid, 19) & "-" & Right(ssid, Len(ssid) - 19)
            ssid = Left(ssid, 30) & "-" & Right(ssid, Len(ssid) - 30)
            ssid = Left(ssid, 41) & "-" & Right(ssid, Len(ssid) - 41)
            lstUsers.AddItem UserName & " " & ssid
            'lblStatus = "Retrieving groups for " & UserName
            'lblStatus.Refresh
            'For Each Group In member.Groups
            '    lstGroups.AddItem Group.Name
            'Next
            'lstUsers.ListIndex = lstUsers.ListCount - 1 ' Initiates the
            lblUserNo = lstUsers.ListCount              ' lstUsers_Click() procedure
            lblUserNo.Refresh
        'End If
        DoEvents
        TotalUsers = TotalUsers + 1
    Next
    'lstResults.ListIndex = lstResults.ListCount - 1
    lblStatus = "Done checking SAM"
    lblStatus.Refresh
End Sub

'Private Sub lstUsers_Click()
'    Dim strUser As String, RenameFile As Boolean, SiteServer As String, msg As String
'    strUser = lstUsers.List(lstUsers.ListIndex)
'    RenameFile = True
'
'    Open strUserDir & "\" & strUser & ".tmp" For Output As #1
'        For lp = 0 To lstGroups.ListCount - 1
'            Print #1, lstGroups.List(lp)
'        Next lp
'    Close 1
'
'    ' In this code, a 'user' is actually a file called PD?????.txt (eg PD71044.txt)
'    ' A new user may actually mean that the file does not exist, not that it is a new domain user
'    '
'    If Dir(strUserDir & "\" & strUser & ".txt") <> "" Then
'        ' Existing user, so check CRC's
'        FileCRC1 = GetCRC(strUserDir & "\" & strUser & ".tmp")
'        FileCRC2 = GetCRC(strPDCDir & "\" & strUser & ".txt")
'
'        If FileCRC1 = FileCRC2 Then
'            ' Dont copy, user has same group membership
'            RenameFile = False
'            ' Remove temp file
'            Kill strUserDir & "\" & strUser & ".tmp"
'        Else
'            ' Files are different.
'            lstChanges.AddItem strUser
'            lstChanges.ListIndex = lstChanges.ListCount - 1
'            lblChangeNo = lstChanges.ListCount
'            msg = "Changed user - " & strUser
'            lstResults.AddItem msg
'            LogLocal msg
'            'Copy file to PDC
'            On Error Resume Next
'                FileCopy strUserDir & "\" & strUser & ".tmp", strPDCDir & "\" & strUser & ".txt"
'                E = Err.Number
'                If E <> 0 Then
'                    msg = "Error " & Err.Number & " (" & Err.Description & ") copying changed user info to " & strPDCDir & " for " & strUser
'                    lstResults.AddItem msg
'                    LogLocal msg
'                    Err.Clear
'                End If
'            On Error GoTo 0
'            ' Rename .tmp to .txt to keep locally
'            If Dir(strUserDir & "\" & strUser & ".txt") <> "" Then
'                Kill strUserDir & "\" & strUser & ".txt"
'            End If
'            Name strUserDir & "\" & strUser & ".tmp" As strUserDir & "\" & strUser & ".txt"
'        End If
'    Else
'        'New user
'        lstNew.AddItem strUser
'        lstNew.ListIndex = lstNew.ListCount - 1
'        lblNewNo = lstNew.ListCount
'        msg = "New user - " & strUser
'        lstResults.AddItem msg
'        LogLocal msg
'        ' Copy .tmp file to PDC as .txt file
'        On Error Resume Next
'            FileCopy strUserDir & "\" & strUser & ".tmp", strPDCDir & "\" & strUser & ".txt"
'            If Err.Number <> 0 Then
'                msg = "Error " & Err.Number & " copying new user info to " & strPDCDir
'                lstResults.AddItem msg
'                LogLocal msg
'                Err.Clear
'            End If
'        On Error GoTo 0
'        ' Rename .tmp to .txt to keep locally
'        Name strUserDir & "\" & strUser & ".tmp" As strUserDir & "\" & strUser & ".txt"
'    End If
'    lstResults.ListIndex = lstResults.ListCount - 1
'End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mnuExit_Click
End Sub

Function GetCRC(filename As String) As Long
    Dim CRC As Long, Data As String, Temp As Integer
    filenum = FreeFile
    Open filename For Binary Access Read As #filenum
        Do Until EOF(filenum)
            Data = InputB(1, #filenum)
            Temp = AscB(Data)
            CRC = CRC + Temp
        Loop
    Close filenum
    GetCRC = CRC
End Function

Private Sub mnuExit_Click()
    Dim msg As String
    msg = "Ended." & vbCrLf
    msg = msg & "--------------------------------------------------------"
    LogLocal msg
    End
End Sub

Private Sub mnuPause_Click()
    If Paused Then
        Paused = False
        mnuPause.Caption = "Pause"
        lblStatus = "Continuing"
    Else
        Paused = True
        mnuPause.Caption = "Continue"
        lblStatus = "Paused"
    End If
    frmMain.Refresh
    If Paused = True Then
        Do Until Paused = False
            DoEvents
        Loop
    End If
End Sub
