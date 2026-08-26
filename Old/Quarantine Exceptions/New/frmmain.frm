VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quarantine Exemptions Application"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   9390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame fmeSource 
      Caption         =   "Source"
      Height          =   1335
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtSourceFile 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "redalert.log"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtSource 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Text            =   "C:\Program files\trend\smex\log"
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Log file"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblSourceDir 
         Caption         =   "Source Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fmeServers 
      Caption         =   "Servers"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
      Begin VB.CheckBox chkServers 
         Caption         =   "EXHSRV4"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkServers 
         Caption         =   "EXHSRV3"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkServers 
         Caption         =   "EXHSRV2"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkServers 
         Caption         =   "EXHSRV1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.Label lblStatusTitle 
      Alignment       =   2  'Center
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   9135
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   9135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdStart_Click()
    Dim Source As String, sourceDir As String, Spath As String
    Dim rec As Variant
    Dim exempt(5000) As String
    Dim exemptloc As String
    Dim cntr As Long
    Dim record As String
    
    record = ""
    exemptloc = "c:\temp\exempt.csv"
    Screen.MousePointer = vbHourglass
    For lp = 0 To 3
        If chkServers(lp).Value = vbChecked Then
            sourceDir = Replace(txtSource, ":", "$")
            Spath = "\\" & chkServers(lp).Caption & "\" & sourceDir & "\"
            Source = Spath & txtSourceFile
            lblStatus = "Processing " & chkServers(lp).Caption
            parse Source, chkServers(lp).Caption
            
            lblStatus = "Loading Exemptions .."
            cntr = 0
            Open exemptloc For Input As #1
            Do While Not EOF(1)
                Input #1, record
                exempt(cntr) = record
                cntr = cntr + 1
            Loop
            Close 1
            
            lblStatus = "Comparing Entries .."
            Open Spath & "redalert.csv" For Input As #2
            Open Spath & "redalert.prs" For Output As #3
            Do While Not EOF(2)
                Line Input #2, record
                record = Left$(record, Len(record) - 1)
                record = Right$(record, Len(record) - 1)
                rec = Split(record, ",")
                For i = 0 To cntr - 1
                    If (StrComp(rec(1), exempt(i), 1) = 0) Or (StrComp(rec(2), exempt(i), 1) = 0) Then
                        If Len(rec(9)) = 0 And Len(rec(4)) = 0 And Len(rec(6)) Then
                            rec(9) = Date & Time
                            ' Fields are:  0-scanneddatetime,1-sender,2-recipient,3-senddatetime,4-virus,5-attachmentname,6-quarantinedir,7-x,8-subject
                            SendMail chkServers(lp).Caption, rec(6), rec(5), rec(2), rec(1), rec(3), rec(8)
                            record = record & " " & rec(9)
                        End If
                    End If
                Next i
                Print #3, record
                record = ""
            Loop
            Close 2
            Close 3
            Kill (Spath & "redalert.csv")
            Name Spath & "redalert.prs" As Spath & "redalert.csv"
        End If
    Next lp
    Screen.MousePointer = vbDefault
    lblStatus = "Idle"
End Sub

Public Sub Form_Load()
    DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=CBDXAAI"
    'Create the Session Object
    Set objSession = CreateObject("mapi.session")
    
    'Logon using the session object
    'Specify a valid profile name if you want to
    'Avoid the logon dialog box
    objSession.Logon profileName:="MS Exchange Settings"
    'Create the Session Object
    Set objSession = CreateObject("mapi.session")

    'Logon using the session object
    'Specify a valid profile name if you want to
    'Avoid the logon dialog box
    objSession.Logon profileName:="MS Exchange Settings"
    'Create the Session Object
    
    ' ADO stuff
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoConn2 = CreateObject("ADODB.Connection")
    Set adoData = CreateObject("ADODB.Recordset")
    
    Me.Refresh
    ' Any commandline option will start automatically
    If Command$ <> "" Then
        For lp = 0 To 3
            chkServers(lp).Value = vbChecked
        Next lp
        cmdStart_Click
        Unload Me
        End
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    objSession.Logoff
    Set objRecipient = Nothing
    Set objMessage = Nothing
    Set objSession = Nothing
    Set adoConn = Nothing
    Unload Me
    
End Sub

