VERSION 5.00
Begin VB.Form frmAddConnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Connection"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmAddConnection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtShare 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Text            =   "ipc$"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtDrive 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtDomain 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cmbLocations 
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Select a Location..."
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   3840
      Y1              =   2540
      Y2              =   2540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   3840
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblShare 
      Alignment       =   2  'Center
      Caption         =   "Share"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblServer 
      Alignment       =   2  'Center
      Caption         =   "Server"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblDrive 
      Alignment       =   2  'Center
      Caption         =   "Drive"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblDomain 
      Caption         =   "Domain"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblLocations 
      Caption         =   "Location"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmAddConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doCreateLocation As Boolean

Private Sub cmdAdd_Click()
    Dim Location As String
    Location = InputBox("Enter a new Location name:", "New location")
    If Location <> "" Then
        cmbLocations.Text = Location
        cmbLocations.AddItem Location
        txtDomain.SetFocus
        doCreateLocation = True
    End If
End Sub

Private Sub cmdSave_Click()
    Dim ID As Integer ', intLoop As Integer, doCreateLocation As Boolean
    
    If txtDomain <> "" And txtUsername <> "" And txtPassword <> "" And txtServer <> "" And txtShare <> "" Then
        SQL = ""
        
        If doCreateLocation Then
            SQL = SQL & "INSERT INTO tblConnectionNames (Name, [Domain], Username, [Password]) VALUES ("
            SQL = SQL & "'" & cmbLocations.Text & "',"
        Else
            SQL = SQL & "INSERT INTO tblConnectionNames ([Domain], Username, [Password]) VALUES ("
        End If
        
        doCreateLocation = False
        
        SQL = SQL & "'" & txtDomain.Text & "',"
        SQL = SQL & "'" & txtUsername.Text & "',"
        SQL = SQL & "'" & txtPassword.Text & "'"
        SQL = SQL & ")"
        adoConn.Open DSN
        adoConn.Execute SQL
        SQL = "SELECT ID FROM tblConnectionNames WHERE Name = '" & cmbLocations.Text & "';"
        adoRS.Open SQL, adoConn
        ID = adoRS("ID")
        adoRS.Close
        
        SQL = ""
        SQL = SQL & "INSERT INTO tblConnections (Driveletter, Server, Share, ConnID) VALUES ("
        SQL = SQL & "'" & txtDrive.Text & "',"
        SQL = SQL & "'" & txtServer.Text & "',"
        SQL = SQL & "'" & txtShare.Text & "',"
        SQL = SQL & "'" & ID & "'"
        SQL = SQL & ")"
        adoConn.Execute SQL
        adoConn.Close
        frmMain.cmbConnections.AddItem cmbLocations.Text
        frmMain.cmbConnections.Text = cmbLocations.Text
        frmMain.Visible = True
        frmMain.WindowState = 0 ' Normal
        Unload frmAddConnection
    End If
End Sub

Private Sub Form_Load()
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    
    cmbLocations.Clear
    
    DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ConnManager.mdb;Persist Security Info=False"
    
    SQL = "SELECT Name FROM tblConnectionNames ORDER BY Name"
    
    adoConn.Open DSN
    adoRS.Open SQL, adoConn
    
    Do Until adoRS.EOF
        cmbLocations.AddItem adoRS("Name")
        adoRS.MoveNext
    Loop
    
    adoRS.Close
    adoConn.Close
End Sub

Private Sub txtServer_GotFocus()
    If Right(txtDrive, 1) = ":" Then
        txtDrive = Left(txtDrive, 1)
        frmAddConnection.Refresh
    End If
End Sub
