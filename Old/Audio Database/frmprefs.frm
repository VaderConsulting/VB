VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrefs 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkKey 
      Caption         =   "View Key"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000A&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   -240
      X2              =   7800
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   -240
      X2              =   7800
      Y1              =   4455
      Y2              =   4455
   End
   Begin VB.Label lblDatabase 
      Caption         =   "Database"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkKey_Click()
    If chkKey.Value = vbChecked Then
        frmDatabase.fmeKey.Visible = True
    Else
        frmDatabase.fmeKey.Visible = False
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim E As Long
    With cdlFiles
        .Filter = "Microsoft Access Databases (*.mdb)|*.mdb"
        .DialogTitle = "Select music database"
        .CancelError = True
    End With
    On Error Resume Next
    cdlFiles.ShowOpen
    E = Err.Number
    If E <> 32755 And E > 0 Then
        MsgBox Error(E), vbCritical + vbOKOnly, "Unexpected error!"
    End If
    If cdlFiles.Filename <> "" Then txtDatabase = cdlFiles.Filename
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtDatabase <> "" Then
        Db = txtDatabase
        SaveSetting "D. Robinson\" & App.EXEName, "Setup", "Db Path", Db
        DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Db & ";Persist Security Info=False"
    End If
    If chkKey.Value = vbChecked Then
        SaveSetting "D. Robinson\" & App.EXEName, "Setup", "Show Key", "True"
    Else
        SaveSetting "D. Robinson\" & App.EXEName, "Setup", "Show Key", "False"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    txtDatabase = GetSetting("D. Robinson\" & App.EXEName, "Setup", "Db Path", App.Path & "\music.mdb")
    If GetSetting("D. Robinson\" & App.EXEName, "Setup", "Show Key", "True") = "True" Then
        chkKey.Value = vbChecked
    Else
        chkKey.Value = vbUnchecked
    End If
    frmDatabase.Show
End Sub
