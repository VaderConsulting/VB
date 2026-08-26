VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Domain Accounts"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9120
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   9120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkEnum 
      Caption         =   "Computers"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox chkEnum 
      Caption         =   "Groups"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox chkEnum 
      Caption         =   "Users"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtTargetDC 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstBat 
      Height          =   2985
      Left            =   9240
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.ListBox lstMembers 
      Height          =   2595
      Left            =   120
      MouseIcon       =   "frmMain.frx":0442
      TabIndex        =   0
      Top             =   480
      Width           =   8895
   End
   Begin VB.Label lblGroups 
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblGroupNo 
      Caption         =   "No of Groups:"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblComputerNo 
      Caption         =   "No of Computers:"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblComputers 
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblUsers 
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblUserNo 
      Caption         =   "No of Users:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblTargetDC 
      Caption         =   "Domain Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   8895
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
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
Public Paused As Boolean
Public TotalMembers As Integer, ServerNo As Integer, strPDCDir As String

Const ERR_ACCESS_DENIED As Long = 70
Const ERR_SUCCESS As Long = 0
Const ERR_MISMATCH As Long = 13

Private Sub Form_Load()
    Dim commandLine As Boolean, msg As String
    msg = "Application Started."
    LogLocal msg
    Paused = False
    Me.Show
    Me.Refresh
    'CheckSAM
End Sub

Sub LogLocal(msg As String)
    Open App.Path & "\appLog" & Format(Date, "DD-MM-YY") & ".log" For Append As #1
        Print #1, Format(Date, "DD/MM/YYYY") & " " & Format(Time, "Long Time") & " " & msg
    Close 1
End Sub

Sub CheckSAM()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msg As String
    msg = "Application Ended."
    LogLocal msg
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuExport_Click()
    Dim Filename As String
    Filename = "c:\temp\Members of " & txtTargetDC & ".csv"
    Open Filename For Output As #1
        Print #1, "Name,Class"
        For lp = 0 To lstMembers.ListCount - 1
            Print #1, lstMembers.List(lp)
        Next lp
    Close 1
    MsgBox "File saved in " & Filename
End Sub

'Private Sub mnuImport_Click()
'    Open "c:\temp\curtinhouseMembers.csv" For Input As #1
'        Do Until EOF(1)
'            Line Input #1, stuff
'        Loop
'        Close
'    Close 1
'End Sub

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

Private Sub mnuStart_Click()
    Dim msg As String
    Dim TotalUsers As Integer
    Dim TotalGroups As Integer
    Dim TotalComputers As Integer
    Dim Domain As IADsDomain
    Dim Computer As IADsComputer
    Dim DomainError As Boolean
    'Dim WMIService As Object
    If txtTargetDC = "" Then
        lblStatus = "DC name must be specified"
        Exit Sub
    End If
    lstMembers.Clear
    
    mnuPause.Enabled = True
    
    lblStatus = "Initialising..."
    lblStatus.Refresh
    lblUsers = 0
    lblGroups = 0
    lblComputers = 0
    
    strDomain = txtTargetDC.Text
    lblStatus = "Contacting " & strDomain
    lblStatus.Refresh
    'On Error Resume Next
    Set Domain = GetObject("WinNT://" & strDomain)
    Select Case Err.Number
        Case ERR_SUCCESS
            msg = "Retrieving members of " & strDomain
            LogLocal msg
            DomainError = False
        Case ERR_MISMATCH
            msg = "Error contacting " & strDomain & ". Is this a Domain name?"
            lstResults.AddItem msg
            lblStatus = msg
            LogLocal msg
            DomainError = True
        Case ERR_ACCESS_DENIED
            msg = "Access denied whilst contacting " & strDomain
            lstResults.AddItem msg
            lblStatus = msg
            LogLocal msg
            DomainError = True
        Case Else
            msg = "Error " & Err.Number & " (" & Err.Description & ") whilst contacting " & strDomain
            msg = Replace(msg, vbCr, " ")
            msg = Replace(msg, vbLf, " ")
            lstResults.AddItem msg
            lblStatus = msg
            LogLocal msg
            DomainError = True
    End Select
    
    If Not DomainError Then
        'Domain.Filter = Array("user")
        TotalMembers = 0
        For Each member In Domain
            If (chkEnum(0) And member.Class = "User") Or (chkEnum(1) And member.Class = "Group") Or (chkEnum(2) And member.Class = "Computer") Then
            UserName = UCase(member.Name)
            
            lblStatus = "Retrieving info for " & UserName
            lblStatus.Refresh
            lstMembers.AddItem UserName & "," & member.Class
            Select Case member.Class
                Case "User"
                    TotalUsers = TotalUsers + 1
                Case "Group"
                    TotalGroups = TotalGroups + 1
                Case "Computer"
                    TotalComputers = TotalComputers + 1
                    'Set WMIService = GetObject("winmgmts:")
                    'Set runningServices = WMIService.ExecQuery("select Name from Win32_Service where State = 'Running'")
            End Select
            
            lblUsers = TotalUsers
            lblGroups = TotalGroups
            lblComputers = TotalComputers
            frmMain.Refresh
            End If
        Next
        lblStatus = "Idle"
        lblStatus.Refresh
    End If
    If lstMembers.ListCount = 0 Then
        mnuExport.Enabled = False
    Else
        mnuExport.Enabled = True
    End If
    mnuPause.Enabled = False
    msg = "Completed task"
    LogLocal msg
    Set Domain = Nothing
End Sub

