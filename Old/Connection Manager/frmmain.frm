VERSION 5.00
Object = "{18D91AD0-D0BE-11D1-A6B4-00AA002075DA}#1.0#0"; "flshtray.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection Manager"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cmbConnections 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Select a Name ..."
      Top             =   120
      Width           =   1815
   End
   Begin TrayIconPrj.TrayIcon tryConnections 
      Left            =   120
      Top             =   480
      _ExtentX        =   1905
      _ExtentY        =   953
      Icon            =   "frmMain.frx":5C12
      ToolTipText     =   "Connection Manager"
      Enabled         =   -1  'True
      TrueClick       =   -1  'True
      Visible         =   -1  'True
      FlashSound      =   0
      FlashIcon       =   "frmMain.frx":5F2C
      FlashInterval   =   1000
      FlashEnabled    =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const CONNECT_UPDATE_PROFILE = &H1
Const RESOURCETYPE_DISK = &H1
Const RESOURCETYPE_PRINT = &H2
Const RESOURCETYPE_ANY = &H0
Const RESOURCE_CONNECTED = &H1
Const RESOURCE_REMEMBERED = &H3
Const RESOURCE_GLOBALNET = &H2
Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
Const RESOURCEDISPLAYTYPE_GENERIC = &H0
Const RESOURCEDISPLAYTYPE_SERVER = &H2
Const RESOURCEDISPLAYTYPE_SHARE = &H3
Const RESOURCEUSAGE_CONNECTABLE = &H1
Const RESOURCEUSAGE_CONTAINER = &H2

Const WN_Success = &H0
Const WN_Not_Supported = &H1
Const WN_Net_Error = &H2
Const WN_Bad_Pointer = &H4
Const WN_Bad_NetName = &H32
Const WN_Bad_Password = &H6
Const WN_Bad_Localname = &H33
Const WN_Access_Denied = &H7
Const WN_Out_Of_Memory = &HB
Const WN_Already_Connected = &H34
Const WN_Server_Down = &H35
Const WN_Server_Down_2 = &H52E
Const WN_NOT_FOUND = &H43
Const WN_IN_USE = 85
Const WN_NOT_CONNECTED = 2250

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalname As String
    lpRemotename As String
    lpComment As String
    lpProvider As String
End Type

Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUsername As String, ByVal dwFlags As Long) As Long

Private Function WnetError(Errcode As Long) As String
    ' Network error code handling
    Select Case Errcode
        Case ERROR_NO_ERROR
            WnetError = "Success."
        Case WN_Not_Supported
           WnetError = "Function is not supported."
        Case WN_Out_Of_Memory
           WnetError = "Out of Memory."
        Case WN_Net_Error
           WnetError = "An error occurred on the network."
        Case WN_Bad_Pointer
           WnetError = "The Pointer was Invalid."
        Case WN_Bad_NetName
           WnetError = "Invalid Network Resource Name."
        Case WN_Bad_Password
           WnetError = "The Password was Invalid."
        Case WN_Bad_Localname
           WnetError = "The local device name was invalid."
        Case WN_Access_Denied
           WnetError = "A security violation occurred."
        Case ERROR_ACCESS_DENIED
           WnetError = "A security violation occurred."
        Case WN_Already_Connected
           WnetError = "The local device was connected to a remote resource."
        Case WN_Server_Down
           WnetError = "Cannot resolve server name or the server you wish to connect to is down."
        Case WN_Server_Down_2
           WnetError = "Cannot resolve server name or the server you wish to connect to is down.*"
        Case WN_NOT_FOUND
           WnetError = "The network name cannot be found."
        Case WN_IN_USE
            WnetError = "The Local Device name is already in use."
        Case WN_NOT_CONNECTED
            WnetError = "The Local Device name is not connected."
        Case ERROR_OPEN_FILES
            WnetError = "One or more files are in use on that connection."
        Case ERROR_ALREADY_ASSIGNED
            WnetError = "The Local Device name is already in use."
        Case ERROR_BAD_DEVICE
            WnetError = "Bad Device Error."
        Case ERROR_SESSION_CREDENTIAL_CONFLICT
            WnetError = "The credentials supplied conflict with an existing set of credentials"
        Case Else:
           WnetError = "Unrecognized Error " + Str(Errcode) + "."
      End Select
End Function

Private Sub cmbConnections_Click()
    If isLoading Then Exit Sub
    cmdConnect.Enabled = True
End Sub

Private Sub cmdAdd_Click()
    frmAddConnection.Show vbNormal
End Sub

Private Sub cmdConnect_Click()
    Dim lpNetResource As NETRESOURCE
    Dim lpUsername As String
    Dim lpPassword As String
    Dim lpLocalname As String
    Dim lpRemotename As String
    Dim rc As Long
    
    SQL = ""
    SQL = SQL & "SELECT tblConnectionNames.[Name], tblConnections.[DriveLetter], tblConnections.[Server], tblConnections.[Share], tblConnectionNames.[Domain], tblConnectionNames.[Username], tblConnectionNames.[Password] "
    SQL = SQL & "FROM tblConnections, tblConnectionNames "
    SQL = SQL & "WHERE tblConnectionNames.[Name] = '" & cmbConnections.Text & "' AND tblConnectionNames.[ID]=tblconnections.[connid]"
    
    adoConn.Open DSN
    adoRS.Open SQL, adoConn
    Do Until adoRS.EOF
        lpUsername = ""
        lpPassword = ""
        lpLocalname = ""
        lpRemotename = ""
        
        lpUsername = adoRS("Domain") & "\" & adoRS("Username")
        lpPassword = adoRS("Password") & ""
        lpLocalname = adoRS("DriveLetter") & ""
        
        If lpLocalname <> "" And Len(lpLocalname) = 1 Then lpLocalname = lpLocalname & ":"
        
        lpRemotename = "\\" & adoRS("Server") & "\" & adoRS("Share") & ""
        
        Debug.Print "NET USE " & lpLocalname & " " & lpRemotename & " /USER:" & lpUsername & " " & lpPassword
        
        lpNetResource.dwType = RESOURCETYPE_DISK
        lpNetResource.dwScope = RESOURCE_GLOBALNET
        lpNetResource.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
        lpNetResource.dwUsage = RESOURCEUSAGE_CONNECTABLE
        lpNetResource.lpLocalname = lpLocalname
        lpNetResource.lpRemotename = lpRemotename
        'lpPassword = Chr(0) ' THESE LINES SCREW THE CODE!!! LEAVE WELL ENOUGH ALONE
        'lpUserName = Chr(0) ' VB Requires these to be NULL, but the API doesn't accept
                             ' NULL, so don't assign values at all.
        rc = WNetAddConnection2(lpNetResource, lpPassword, lpUsername, 0)
        If rc <> 0 Then
            MsgBox "Error " & rc & " (" & WnetError(rc) & ") ", vbCritical + vbOKOnly, "Error connecting drive"
        End If
        adoRS.MoveNext
    Loop
    
    adoRS.Close
    adoConn.Close
    
    frmMain.Hide
End Sub

Private Sub Form_Activate()
    Dim SQL As String
    
    isLoading = True
    cmbConnections.Clear
    cmbConnections.Text = "Select a Name ..."
    
    SQL = "SELECT Name FROM tblConnectionNames ORDER BY Name"
    
    adoConn.Open DSN
    adoRS.Open SQL, adoConn
    
    Do Until adoRS.EOF
        cmbConnections.AddItem adoRS("Name")
        adoRS.MoveNext
    Loop
    
    adoRS.Close
    adoConn.Close
    
    isLoading = False
End Sub

Private Sub Form_Deactivate()
    frmMain.Visible = False
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState = vbMinimized Then
        Me.Visible = False
        cmdConnect.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set adoRS = Nothing
    Set adoConn = Nothing
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub tryConnections_LeftButtonClick()
    frmMain.Visible = True
    frmMain.WindowState = vbNormal
    frmMain.Show
    frmMain.Refresh
End Sub

Private Sub tryConnections_LeftButtonDoubleClick()
    frmMain.Visible = True
    frmMain.WindowState = vbNormal
    frmMain.Show
    frmMain.Refresh
End Sub

Private Sub tryConnections_RightButtonClick()
    frmMain.PopupMenu mnuFile
End Sub
