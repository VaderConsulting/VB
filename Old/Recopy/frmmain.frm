VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recursive File Copy"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbrCopy 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDestination 
      Caption         =   "..."
      Height          =   285
      Left            =   7200
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "..."
      Height          =   285
      Left            =   7200
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdlDestination 
      Left            =   7680
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlSource 
      Left            =   7680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDestination 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   5775
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label lblDestination 
      Caption         =   "Destination"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblSource 
      Caption         =   "Source"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDestination_Click()
    
    txtDestination = GetBrowseDirectory(Me.hWnd)
    'On Error Resume Next
    'With cdlDestination
    '    .Filter = "Program Files (*.exe;*.com)|*.exe;*.com|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    '    .CancelError = True
    '    .ShowOpen
    'End With
    'On Error GoTo 0
    'If Err = 0 Then
    '    txtDestination.Text = cdlDestination.FileName
    'End If
End Sub

Private Sub cmdSource_Click()
    On Error Resume Next
    With cdlSource
        .Filter = "Program Files (*.exe;*.com)|*.exe;*.com|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .CancelError = True
        .ShowOpen
    End With
    On Error GoTo 0
    If Err = 0 Then
        txtSource.Text = cdlSource.FileName
    End If
End Sub

Public Function GetBrowseDirectory(OwnerHandle As Long) As String
    Const BIF_RETURNONLYFSDIRS = &H1
    
    Dim BInfo As BROWSEINFO
    Dim r As Long
    Dim pidl As Long
    Dim sPath As String
    Dim pos As Integer
    
    BInfo.hOwner = OwnerHandle
    BInfo.pidlRoot = 0&
    BInfo.lpszTitle = "Choose a directory from the list."
    BInfo.ulFlags = BIF_RETURNONLYFSDIRS
    pidl = SHBrowseForFolder(BInfo)
    
    sPath = Space$(512)
    r = SHGetPathFromIDList(ByVal pidl, ByVal sPath)
    
    If r Then
        pos = InStr(sPath, Chr$(0))
        sPath = Left(sPath, pos - 1)
        
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        GetBrowseDirectory = sPath
    Else
        GetBrowseDirectory = ""
    End If
End Function
 
 


