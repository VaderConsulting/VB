VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Users from 'Same As' user"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7050
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlImport 
      Left            =   6000
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grdUsers 
      Height          =   3135
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtYourUsername 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Caption         =   "NOTE: This application imports the selected file, and inserts the info into tblCreateUsers on CBDXAAI"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   5775
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Authority"
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Call #"
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "'Same As'"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "DOB"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Surname"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "First Name"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Username"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Caption         =   "This application requires the import file follow this format (separated by comma's)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5775
   End
   Begin VB.Label lblYourUsername 
      Caption         =   "Your Username"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoConn As ADODB.Connection
Dim adoRS As ADODB.Recordset
Dim SQL As String
Dim DSN As String
Dim SQLServer As String
Dim Database As String

Private Sub cmdImport_Click()
  Dim Username, FirstName, LastName, DOB, SameAs, CallNo, Auth
  Dim i As Integer
  
  cdlImport.Filter = "Comma Separated Values (*.csv)|*.csv|All Files (*.*)|*.*"
  cdlImport.InitDir = App.Path
  cdlImport.CancelError = True
  On Error Resume Next
  Do Until cdlImport.FileName <> "" Or Err.Number <> 0
    cdlImport.ShowOpen
  Loop
  If Err.Number = 0 Then
    grdUsers.Clear
    Open cdlImport.FileName For Input As #1
      i = 0
      Line Input #1, stuff ' throw away first line
      Do Until EOF(1)
        i = i + 1
        Line Input #1, stuff
      Loop
    Close 1
    grdUsers.Rows = i - 1
    Open cdlImport.FileName For Input As #1
      Line Input #1, stuff ' throw away first line
      For lp = 0 To i - 1
        Input #1, Username, FirstName, LastName, DOB, SameAs, CallNo, Auth
        grdUsers.Row = lp
        grdUsers.Col = 0
        grdUsers.Text = Username
        grdUsers.Col = 1
        grdUsers.Text = FirstName
        grdUsers.Col = 2
        grdUsers.Text = LastName
        grdUsers.Col = 3
        grdUsers.Text = DOB
        grdUsers.Col = 4
        grdUsers.Text = SameAs
        grdUsers.Col = 5
        grdUsers.Text = CallNo
        grdUsers.Col = 6
        grdUsers.Text = Auth
      Next
    Close 1
    ' Go through each sameAs user, retrieving info (ADSI)
    ' Go through each group for each user, retrieving group name (ADSI)
    ' Once you have this info, save to DB
  End If
  On Error GoTo 0
End Sub

Private Sub Form_Load()
  Set adoConn = CreateObject("ADODB.Connection")
  Set adoRS = CreateObject("ADODB.Recordset")
  
  SQLServer = "CBDXAAI"
  Database = "POLICE"
  
  DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Database & ";Data Source=" & SQLServer
  'SQL = "SELECT * FROM tblCreateUsers WHERE complete = 0"
  adoConn.Open DSN
  
  txtYourUsername = Environ("username")
  
  
  adoConn.Close
  'adoRS.Close
  Set adoRS = Nothing
  Set adoConn = Nothing
End Sub
