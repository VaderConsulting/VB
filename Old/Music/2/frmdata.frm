VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmDatabase 
   Caption         =   "More Info required..."
   ClientHeight    =   1290
   ClientLeft      =   2160
   ClientTop       =   8295
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1290
   ScaleWidth      =   6720
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdDatabaseBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtDatabasePath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DefaultExt      =   "bmp"
      DialogTitle     =   "Open Graphic"
      FileName        =   "*.bmp"
      Filter          =   "Bitmaps (*.bmp;*.ico)|*.bmp;*.ico"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Music database path:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDatabaseBrowse_Click()
    dlgBrowse.Filter = "Access Databases (*.MDB)|*.mdb"
    dlgBrowse.filename = "MUSIC.MDB"
    dlgBrowse.DialogTitle = "More info required"
    dlgBrowse.ShowOpen
    Do Until dlgBrowse.filename <> "*.MDB"
        DoEvents
    Loop
    txtDatabasePath = dlgBrowse.filename
End Sub

Private Sub cmdOK_Click()
    If txtDatabasePath <> "" Then SaveSetting gINIFile, "Setup", "Database", txtDatabasePath
    gDatabasePath = txtDatabasePath
    frmDatabase.Hide
End Sub

Private Sub Form_Load()
    'Width = gVGAWidth
    'Height = gVGAHeight
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub
