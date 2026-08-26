VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asset Control"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeRecordset 
      Caption         =   "Recordset"
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5655
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   5
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   4
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblfield 
         Caption         =   "Description"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblfield 
         Caption         =   "Type"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblfield 
         Caption         =   "Category"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblfield 
         Caption         =   "Serial"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblfield 
         Caption         =   "Asset"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblfield 
         Caption         =   "Barcode"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdFetch 
      Caption         =   "Fetch"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtBarcode 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblBarcode 
      Caption         =   "Barcode:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenLast 
         Caption         =   "1"
         Index           =   0
      End
      Begin VB.Menu mnuOpenLast 
         Caption         =   "2"
         Index           =   1
      End
      Begin VB.Menu mnuOpenLast 
         Caption         =   "3"
         Index           =   2
      End
      Begin VB.Menu mnuOpenLast 
         Caption         =   "4"
         Index           =   3
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_RETURNFSANCESTORS = 8
Private Const BIF_EDITBOX = 16
Private Const BIF_NEWDIALOGSTYLE = 64
Private Const BIF_BROWSEINCLUDEFILES = 16384

Private Const MAX_PATH = 260

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public oDatabase As ADODB.Recordset
Public oConn As ADODB.Connection
Public DSN As String

Private Sub cmdFetch_Click()
    Dim strSQL As String
    Dim iRecordCount As Integer
    Dim strLabel As String
    
    txtBarcode.Text = Left(txtBarcode.Text, 8)
    
    strSQL = "SELECT * FROM Assets WHERE Barcode like '" & txtBarcode.Text & "' ORDER BY Barcode"
    
    oDatabase.Open strSQL, oConn, adOpenDynamic, adLockOptimistic
    lblInfo = "Opened Database."
    
    iRecordCount = oDatabase.RecordCount
    
    Select Case iRecordCount
        Case 1
            strLabel = " record "
        Case Else
            strLabel = " records "
    End Select
    
    lblInfo = iRecordCount & strLabel & "found."
    If iRecordCount = 0 Then
        txtData(0) = ""
        txtData(1) = ""
        txtData(2) = ""
        txtData(3) = ""
        txtData(4) = ""
        txtData(5) = ""
        
        strSQL = ""
        strSQL = strSQL & "INSERT INTO ASSETS "
        strSQL = strSQL & "(BARCODE,[ASSET SERIAL NO],[LAST STOCKTAKE DATE],[FIRST FOUND],ADDED,[ASSET TYPE],MANUFACTURER,STATUS,NETWORK,[USER GROUP],FLOOR_ROOM_WS,[PURCHASE DATE],[PURCHASE PRICE],[PARENT_LOCATION],CCC) VALUES "
        strSQL = strSQL & "('" & txtBarcode.Text & "','[Please Add]','" & Date & "',-1,'" & Now & "',1801,'UNKNOWN',30,10,31,1,'01 January 1900',0,162963,131)"
        oConn.Execute strSQL
        lblInfo = "Record added."
        txtBarcode.Text = ""
        txtBarcode.SetFocus
    Else
        oDatabase.Close
        
        strSQL = "SELECT Assets.Asset, Assets.Barcode, Assets.[Asset Serial No], as_category.description, asset_types.description, asset_descriptions.description FROM (asset_types RIGHT JOIN (asset_descriptions RIGHT JOIN Assets ON asset_descriptions.adesc_no = Assets.[Asset Type Description]) ON asset_types.atype_no = Assets.[Asset Type]) LEFT JOIN as_category ON asset_types.category_no = as_category.category_no WHERE Assets.Barcode LIKE '" & txtBarcode.Text & "'"
        oDatabase.Open strSQL, oConn, adOpenDynamic, adLockOptimistic
        
        txtData(0) = oDatabase.Fields("Barcode").Value
        txtData(1) = oDatabase.Fields("Asset").Value & ""
        txtData(2) = oDatabase.Fields("Asset Serial No").Value & ""
        txtData(3) = oDatabase.Fields("As_Category.Description").Value & ""
        txtData(4) = oDatabase.Fields("Asset_Types.Description").Value & ""
        txtData(5) = oDatabase.Fields("Asset_Descriptions.Description").Value & ""
        txtBarcode.Text = ""
        txtBarcode.SetFocus
    End If
    oDatabase.Close
End Sub

Private Sub Form_Load()
    Dim LastFile(3) As String
    Dim i As Integer
    
    For i = 0 To 3
        LastFile(i) = GetSetting("Asset Control", "Previous", CStr(i), "")
        If LastFile(i) <> "" Then
            mnuOpenLast(i).Caption = LastFile(i)
            mnuOpenLast(i).Visible = True
        Else
            mnuOpenLast(i).Visible = False
        End If
    Next
    
    Set oDatabase = CreateObject("ADODB.Recordset")
    Set oConn = CreateObject("ADODB.Connection")
    
    mnuOpen.Enabled = True
    mnuClose.Enabled = False
    cmdFetch.Enabled = False
    txtBarcode.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    ExitApp
End Sub

Private Sub mnuAbout_Click()
    MsgBox "For assisting CSC/Saab/RAN personnel in Asset Control" & vbCrLf & vbCrLf & "(c) 2002 D. Robinson", vbOKOnly + vbInformation, "Asset Control"
End Sub

Private Sub mnuClose_Click()
    On Error GoTo Hell
    Select Case oDatabase.State
        Case ADODB.adStateOpen
            lblInfo = "Closing database..."
            frmMain.Refresh
            oDatabase.Close
        Case ADODB.adStateClosed
            ' Do nothing
        Case Else
            Do Until oDatabase.State = ADODB.adStateOpen
                lblInfo = "Waiting to close database..."
                frmMain.Refresh
                DoEvents
            Loop
            oDatabase.Close
    End Select
    lblInfo = "Database closed."
    
    mnuOpen.Enabled = True
    mnuClose.Enabled = False
    
    frmMain.Refresh
    Exit Sub
Hell:
    lblInfo = "Error " & Err.Number & " attempting to close database."
    frmMain.Refresh
End Sub

Private Sub mnuExit_Click()
    ExitApp
End Sub

Sub ExitApp()
    ' Close database objects,
    If oDatabase.State = ADODB.adStateOpen Then
        oDatabase.Close
    End If
    
    If oConn.State = ADODB.adStateOpen Then
        oConn.Close
    End If
    
    ' Remove references, before finally,
    
    Set oDatabase = Nothing
    Set oConn = Nothing
    
    ' Close application
    End
End Sub

Private Sub mnuOpen_Click()
    Dim strDatabaseName As String
    Dim OpenFile As OPENFILENAME, SelectedFile As String
    Dim lReturn As Long
    Dim sFilter As String
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hWndOwner = frmMain.hWnd
    OpenFile.hInstance = App.hInstance
    sFilter = "All Files (*.*)" & Chr(0) & "*.*" & Chr(0) & "Databases (*.mdb)" & Chr(0) & "*.MDB" & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 2
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = App.Path
    OpenFile.lpstrTitle = "Select database"
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        'The User pressed the Cancel Button
    Else
        SelectedFile = Trim(OpenFile.lpstrFile)
        If SelectedFile <> "" Then
            strDatabaseName = SelectedFile
        Else
            strDatabaseName = ""
        End If
        
        ' Remove CHR(0)'s
        strDatabaseName = Replace(strDatabaseName, Chr(0), "")
        
        If strDatabaseName <> "" Then
            OpenDatabase strDatabaseName
        End If
        
    End If

End Sub

Private Sub mnuOpenLast_Click(Index As Integer)
    OpenDatabase mnuOpenLast(Index).Caption
End Sub

Private Sub txtBarcode_Change()
    If txtBarcode.Text <> "" Then
        cmdFetch.Enabled = True
    Else
        cmdFetch.Enabled = False
    End If
End Sub

Private Sub txtBarcode_GotFocus()
    txtBarcode.SelStart = 0
    txtBarcode.SelLength = Len(txtBarcode.Text)
End Sub

Sub OpenDatabase(strFilename As String)
    DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & strFilename & ";Mode=ReadWrite"
    oConn.Open DSN
    oConn.CursorLocation = adUseClient
    
    mnuOpen.Enabled = False
    mnuClose.Enabled = True
    txtBarcode.Enabled = True
    lblInfo = "Opened connection to Database."
    txtBarcode.SetFocus
    SaveSetting "Asset Control", "Previous", 0, strFilename
End Sub


Private Sub txtBarcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmdFetch_Click
    End If
End Sub
