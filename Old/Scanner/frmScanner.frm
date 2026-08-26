VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scanner Interface"
   ClientHeight    =   1872
   ClientLeft      =   144
   ClientTop       =   432
   ClientWidth     =   10176
   Icon            =   "frmScanner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1872
   ScaleWidth      =   10176
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   10
      Top             =   480
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   9
      Top             =   840
      Width           =   3735
   End
   Begin VB.ComboBox cmbFields 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   3735
   End
   Begin VB.ComboBox cmbTables 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   0
      Top             =   1440
      _ExtentX        =   837
      _ExtentY        =   837
      _Version        =   327681
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6360
      TabIndex        =   13
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Field example:"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Table 2:"
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Field 2:"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Field example:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblFieldData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label lblFields 
      Caption         =   "Field:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lbltable 
      Caption         =   "Table:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lbldbTitle 
      Caption         =   "Database:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lbldatabase 
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Begin VB.Menu mnuDatabase 
            Caption         =   "Database"
         End
         Begin VB.Menu mnuTable 
            Caption         =   "Table"
         End
         Begin VB.Menu mnuField 
            Caption         =   "Field"
         End
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "Mode"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
         Begin VB.Menu mnuBarcode 
            Caption         =   "Barcode for an item to existing table"
         End
         Begin VB.Menu mnuData 
            Caption         =   "Item and barcode to existing table"
         End
      End
      Begin VB.Menu mnuQuery 
         Caption         =   "Query"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFields_Click()
    On Error GoTo erra
    lblFieldData = ""
    Dim db As Database          'set up the database to be read
    Dim td As TableDef
    Dim rs As Recordset
    Dim fds As Fields
    Dim fld As Field
    Set db = OpenDatabase(lbldatabase)
    Set rs = db.OpenRecordset(cmbTables.List(cmbTables.ListIndex))
    rs.MoveFirst
    lblFieldData = rs.Fields(cmbFields.ListIndex) & ""  ' the & "" is for null parsing
    MsgBox "You are now ready to scan", vbInformation + vbOKOnly, "Congratulations"
    Exit Sub
erra:
    ' Oh oh, an unexpected error occurred.  Display this error to the user.
    MsgBox "Unexpected error. (" & Error$(Err) & ")" & vbCrLf & vbCrLf & "Try checking your network connection.", vbCritical + vbOKOnly, "Critical Error"
End Sub

Private Sub cmbTables_Change()
    mnuField_Click
End Sub

Private Sub cmbTables_Click()
    mnuField_Click
End Sub

Private Sub Form_Load()
    Left = 0    'leftmost edge of window relative to screen
    Top = 0     'topmost edge of window relative to screen
    frmInterface.Width = 5250
    lblFields.Caption = "Field:"
    mnuDatabase.Enabled = True
    mnuTable.Enabled = False
    mnuField.Enabled = False
    mnuMode.Enabled = False
End Sub

Private Sub mnuBarcode_Click()
    mnuBarcode.Checked = True
    mnuQuery.Checked = False
    mnuData.Checked = False
End Sub

Private Sub mnuData_Click()
    mnuData.Checked = True
    mnuBarcode.Checked = False
    mnuQuery.Checked = False
End Sub

Private Sub mnuQuery_Click()
    mnuQuery.Checked = True
    mnuBarcode.Checked = False
    mnuData.Checked = False
End Sub

Private Sub mnuDatabase_Click()
    On Error GoTo erra                          'start error trapping
    cdlFiles.filename = ""
    cmbTables.Clear
    cmbFields.Clear
    lblFieldData = ""
    While cdlFiles.filename = ""                'do this continuously until a filename is selected
        With cdlFiles
            .CancelError = True                 'enable the user to cancel the open dialog
            .DefaultExt = ".mdb"                'Default extension
            .DialogTitle = "Select database"    'Window title
            .Filter = "*.MDB|*.MDB"             'filename filter
            .InitDir = "J:\Databases"           'initial directory
            .ShowOpen                           'display the open dialog box
        End With
    Wend
    lbldatabase = cdlFiles.filename             'show the name of the database
    mnuTable.Enabled = True                     'allow the selection of a table
    mnuTable_Click                              'manually click the table menu option
    Exit Sub                                    'end this subroutine
erra:                                           'error parsing
    If Err = cdlCancel Then ' the user selected the cancel button on the dialog box
        MsgBox "You selected Cancel.  No database will be loaded", vbExclamation + vbOKOnly, "Open database"
        Exit Sub
    End If
    ' Oh oh, an unexpected error occurred.  Display this error to the user.
    MsgBox "Unexpected error. (" & Error$(Err) & ")" & vbCrLf & vbCrLf & "Try checking your network connection.", vbCritical + vbOKOnly, "Critical Error"
End Sub

Private Sub mnuExit_Click()
    ' ask them if they want to exit
    reply = MsgBox("Do you really want to exit?", vbYesNo + vbQuestion, "Exit")
    If reply = vbYes Then End   ' exit program if they selected Yes
End Sub

Private Sub mnuField_Click()
    cmbFields.Clear
    lblFieldData = ""
    On Error GoTo erra
    Dim db As Database          'set up the database to be read
    Dim td As TableDef
    Dim rs As Recordset
    Dim fds As Fields
    Dim fld As Field
    Set db = OpenDatabase(lbldatabase)
    Set rs = db.OpenRecordset(cmbTables.List(cmbTables.ListIndex))
    Set fds = rs.Fields
    For Each fld In fds         'add each field name to the fields combo box
        cmbFields.AddItem fld.Name
    Next
    cmbFields.Enabled = True    'enable the fields combo box
    mnuMode.Enabled = True      'make the Mode type (Add or Query) selectable
    Exit Sub
erra:                                           'error parsing
    ' Oh oh, an unexpected error occurred.  Display this error to the user.
    MsgBox "Unexpected error. (" & Error$(Err) & ")" & vbCrLf & vbCrLf & "Try checking your network connection.", vbCritical + vbOKOnly, "Critical Error"
End Sub

Private Sub mnuMode_Click()
    ' this subroutine changes the current mode of operation to/from add or query
    If mnuMode.Caption = "Add" Then
        mnuMode.Caption = "Query"
        frmInterface.Width = 5250
        lbltable.Caption = "Table:"
        lblFields.Caption = "Field:"
        Exit Sub
    End If
    If mnuMode.Caption = "Query" Then
        mnuMode.Caption = "Add"
        frmInterface.Width = 10260
        lbltable.Caption = "Table 1:"
        lblFields.Caption = "Field 1:"
        Exit Sub
    End If
End Sub

Private Sub mnuTable_Click()
    cmbTables.Clear
    cmbFields.Clear
    lblFieldData = ""
    Dim db As Database
    Dim td As TableDef
    Set db = OpenDatabase(lbldatabase)
    For Each td In db.TableDefs
        cmbTables.AddItem td.Name
    Next
    cmbTables.Enabled = True
    mnuField.Enabled = True
End Sub
