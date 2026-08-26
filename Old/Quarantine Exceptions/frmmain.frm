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
    Dim Source As String, sourceDir As String
    Screen.MousePointer = vbHourglass
    For lp = 0 To 3
        If chkServers(lp).Value = vbChecked Then
            sourceDir = Replace(txtSource, ":", "$")
            'On Error Resume Next
                Source = "\\" & chkServers(lp).Caption & "\" & sourceDir & "\" & txtSourceFile
                dr = Dir(Source)
                If dr <> "" Then
                    lblStatus = "Processing " & chkServers(lp).Caption
                    Parse Source, chkServers(lp).Caption
                End If
            'On Error GoTo 0
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

