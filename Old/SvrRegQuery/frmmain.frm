VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registry Query Tool for multiple targets"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   5640
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstResults 
      Height          =   2985
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   5415
   End
   Begin VB.ListBox lstKeys 
      Height          =   2595
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.ListBox lstTargets 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   5160
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   5415
   End
   Begin VB.Label lblKeys 
      Alignment       =   2  'Center
      Caption         =   "Keys"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblTargets 
      Alignment       =   2  'Center
      Caption         =   "Targets"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9000
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   9000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Begin VB.Menu mnuTargets 
            Caption         =   "Targets"
         End
         Begin VB.Menu mnuKeys 
            Caption         =   "Keys"
         End
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Begin VB.Menu mnuResults 
            Caption         =   "Results"
         End
      End
      Begin VB.Menu mnuTargets2 
         Caption         =   "Targets"
         Begin VB.Menu mnuAdd 
            Caption         =   "Add/Remove..."
         End
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Const DRIVE_LOCAL As Integer = 2
Const DRIVE_NETWORK As Integer = 3
Const DRIVE_CDROM As Integer = 4

Private Function GetValue(ByVal Hostname As String, ByVal Key As String, ByVal Value As String, ByRef ECode As Boolean) As String
    Dim OpenKeyVal As Long
    Dim OpenHiveVal As Long
    Dim RResult As Long
    Dim InfoTextStr As String
    
    'Init
    ECode = False
    Hostname = Trim(Hostname)
    Key = Trim(Key)
    Value = Trim(Value)
    GetValue = ""
        
    RResult = RegConnectRegistry(Hostname, HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then
      GetValue = "No Data"
      Exit Function
    End If
    
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
    InfoTextStr = RegistryQueryValue(OpenKeyVal, Value, REG_SZ)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    GetValue = InfoTextStr
    ECode = True
End Function

Private Function CreateKey(ByVal Hostname As String, ByVal Key As String, NewKey As String) As Boolean
    Dim OpenKeyVal As Long
    Dim OpenHiveVal As Long
    Dim RResult As Long
    Dim lCreateResult As Long
    
    'Init
    Hostname = Trim(Hostname)
    Key = Trim(Key)
        
    RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then
      CreateKey = False
      Exit Function
    End If
    
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
    lCreateResult = RegistryCreateKey(OpenKeyVal, NewKey)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    CreateKey = True
    
End Function

Private Function SetValueString(Hostname As String, Key As String, Value As String, Data As String, regType As Long) As Long
    Dim OpenKeyVal As Long
    Dim RResult As Long
    Dim OpenHiveVal As Long
    Dim strValue, CMPTRName, keytogo As String, InfoTextStr As String
    Dim x As Integer
    
    'GetIPCConnection (Trim(Hostname))
    
    CMPTRName = Trim(Hostname)
    
    RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
    keytogo = Trim(Key)
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
    RegistryWriteValue Data, OpenKeyVal, Value, regType
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    'DisIPCConnection (Trim(Hostname))
        
End Function

Private Sub mnuAdd_Click()
  frmTargets.Show vbModal
End Sub

Private Sub mnuExit_Click()
  Unload Me
  End
End Sub

Private Sub mnuKeys_Click()
  With cdlFiles
    .DefaultExt = ".txt"
    .DialogTitle = "Select list of keys to import"
    .Filter = "Text files (*.txt)|*.txt|All Files (*.*)|*.*"
    .InitDir = "c:\temp"
    .CancelError = True
  End With
  On Error Resume Next
    Do Until cdlFiles.FileName <> "" Or Err.Number <> 0
      cdlFiles.ShowOpen
    Loop
    If Err.Number = 32755 Then
      cdlFiles.FileName = ""
      Exit Sub
    End If
    If Err.Number <> 0 Then
      MsgBox "Unexpected Error - " & Err.Number & " (" & Err.Description & ")"
      cdlFiles.FileName = ""
      Exit Sub
    End If
  On Error GoTo 0
  lstKeys.Clear
  Open cdlFiles.FileName For Input As #1
    Do Until EOF(1)
      Line Input #1, keyname
      If keyname <> "" Then
        lstKeys.AddItem UCase(keyname)
      End If
    Loop
  Close 1
  cdlFiles.FileName = ""
End Sub

Private Sub mnuResults_Click()
  With cdlFiles
    .DefaultExt = ".txt"
    .DialogTitle = "Export results to file"
    .Filter = "Text files (*.txt)|*.txt|All Files (*.*)|*.*"
    .InitDir = "c:\temp"
    .CancelError = True
  End With
  On Error Resume Next
    Do Until cdlFiles.FileName <> "" Or Err.Number <> 0
      cdlFiles.ShowSave
    Loop
    If Err.Number = 32755 Then
      cdlFiles.FileName = ""
      Exit Sub
    End If
    If Err.Number <> 0 Then
      MsgBox "Unexpected Error - " & Err.Number & " (" & Err.Description & ")"
      cdlFiles.FileName = ""
      Exit Sub
    End If
  On Error GoTo 0
  
  Open cdlFiles.FileName For Output As #1
    For lp1 = 0 To lstResults.ListCount - 1
      Print #1, lstResults.List(lp1)
    Next lp1
  Close 1
  cdlFiles.FileName = ""
End Sub

Private Sub mnuStart_Click()
  Dim RegValue As String, lp1 As Long, lp2 As Long
  Dim KeyToQuery As String, ValueToQuery As String, HostResults As String
  
  lstResults.Clear
  
  For lp1 = 0 To lstTargets.ListCount - 1           '             Get Target names
    Hostname = lstTargets.List(lp1)
    For lp2 = 0 To lstKeys.ListCount - 1            '             Get Key names
      For lp3 = Len(lstKeys.List(lp2)) To 1 Step -1 '             Extract Key and Value
        If Mid(lstKeys.List(lp2), lp3, 1) = "\" Then Exit For
      Next lp3
      If lp3 <> 1 Then
        RegValue = "No Data"
        KeyToQuery = Left(lstKeys.List(lp2), lp3 - 1)
        ValueToQuery = Right(lstKeys.List(lp2), Len(lstKeys.List(lp2)) - lp3)
        RegValue = GetValue(Hostname, KeyToQuery, ValueToQuery, False)
        If Trim(RegValue) <> "" Then
          HostResults = "\\" & Hostname & "\" & "HKLM\" & KeyToQuery & "\" & ValueToQuery & " = " & RegValue
        Else
          HostResults = "\\" & Hostname & "\" & "HKLM\" & KeyToQuery & "\" & ValueToQuery & " NOT FOUND"
        End If
        lstResults.AddItem HostResults
        Debug.Print HostResults
        HostResults = ""
      End If
    Next lp2
  Next lp1

End Sub

Private Sub mnuTargets_Click()
  With cdlFiles
    .DefaultExt = ".txt"
    .DialogTitle = "Select list of servers to import"
    .Filter = "Text files (*.txt)|*.txt|All Files (*.*)|*.*"
    .InitDir = "c:\temp"
    .CancelError = True
  End With
  On Error Resume Next
    Do Until cdlFiles.FileName <> "" Or Err.Number <> 0
      cdlFiles.ShowOpen
    Loop
    If Err.Number = 32755 Then
      cdlFiles.FileName = ""
      Exit Sub
    End If
    If Err.Number <> 0 Then
      MsgBox "Unexpected Error - " & Err.Number & " (" & Err.Description & ")"
      cdlFiles.FileName = ""
      Exit Sub
    End If
  On Error GoTo 0
  lstTargets.Clear
  Open cdlFiles.FileName For Input As #1
    Do Until EOF(1)
      Line Input #1, servername
      If servername <> "" Then
        lstTargets.AddItem UCase(servername)
      End If
    Loop
  Close 1
  cdlFiles.FileName = ""
End Sub
