VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmSetup 
   Caption         =   "Setup"
   ClientHeight    =   6915
   ClientLeft      =   2175
   ClientTop       =   1455
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6915
   ScaleWidth      =   9600
   Begin VB.TextBox txtDatabasePath 
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Top             =   4080
      Width           =   6495
   End
   Begin VB.CommandButton cmdDatabaseBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   8400
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   13
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtPath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   3240
      Width           =   6495
   End
   Begin VB.Frame fmeFooter 
      Caption         =   "Footer Format"
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton optFormat 
         Caption         =   "Custom"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optFormat 
         Caption         =   "Informal"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optFormat 
         Caption         =   "Standard"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fmeLogo 
      Caption         =   "Program Logo"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
      Begin VB.OptionButton optLogo 
         Caption         =   "User defined"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optLogo 
         Caption         =   "Default 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optLogo 
         Caption         =   "Default 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtFooter 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   6855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   6480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   9000
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
      Alignment       =   1  'Right Justify
      Caption         =   "Music Database Path"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Line lneSeperator 
      Index           =   1
      X1              =   120
      X2              =   9480
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lblPath 
      Alignment       =   1  'Right Justify
      Caption         =   "Logo Path"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Image imgLogo 
      Height          =   855
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   975
   End
   Begin VB.Line lneSeperator 
      Index           =   0
      X1              =   120
      X2              =   9480
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblExample 
      Alignment       =   2  'Center
      Caption         =   "Example"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
    dlgBrowse.filename = "DEFAULT.BMP"
    dlgBrowse.ShowOpen
    Do Until dlgBrowse.filename <> "DEFAULT.BMP"
        DoEvents
    Loop
    txtPath = dlgBrowse.filename
    imgLogo.Picture = LoadPicture(txtPath)
End Sub

Private Sub cmdCancel_Click()
    frmStart.Show
    frmSetup.Hide
End Sub

Private Sub cmdDatabaseBrowse_Click()
    dlgBrowse.Filter = "Access Databases (*.MDB)|*.mdb"
    dlgBrowse.filename = "MUSIC.MDB"
    dlgBrowse.DialogTitle = "Open Music Database"
    dlgBrowse.ShowOpen
    Do Until dlgBrowse.filename <> "MUSIC.MDB"
        DoEvents
    Loop
    txtDatabasePath = dlgBrowse.filename
    'imgLogo.Picture = LoadPicture(txtPath)
End Sub

Private Sub cmdOK_Click()
    'Result = fctSaveSettings(lstUsers.List(lstUsers.ListIndex), UserPassword, UserLevel, UserLock)
    gFooter = txtFooter
    If gRegistered = 1 Then gLogoPath = txtPath
    For varLoop = 0 To 2
        If optLogo(varLoop).Value <> 0 Then gLogo = varLoop
    Next varLoop
    SaveSetting gINIFile, "Setup", "Logo", gLogo
    If gLogoPath <> "" Then SaveSetting gINIFile, "Setup", "LogoPath", gLogoPath
    If txtDatabasePath <> "" Then SaveSetting gINIFile, "Setup", "Database", txtDatabasePath
    frmStart.Show
    frmSetup.Hide
End Sub

Private Sub Form_Activate()
    If gCompany = "" Then gCompany = gBusiness
    If gPlace = "" Then gPlace = gHome
    If gName = "" Then gName = gAuthorChristian
    If gSurname = "" Then gSurname = gAuthorSurname
    txtFooter = gCompany & ", " & gPlace & " Printed " & Str$(Date)
    gFooter = txtFooter
    optLogo(gLogo).Value = 1
    If gRegistered = 1 Then
        optLogo(2).Enabled = 1
        txtPath.Enabled = 1
        cmdBrowse.Enabled = 1
    End If
    Logo = GetSetting(gINIFile, "Setup", "Logo", 0)
    If Logo = 0 Then imgLogo.Picture = frmLogin.imgCBB.Picture
    If Logo = 1 Then imgLogo.Picture = frmLogin.imgKeys.Picture
    If Logo = 2 And txtPath <> "" Then imgLogo.Picture = LoadPicture(txtPath)
    If Logo = 2 And txtPath = "" Then
        txtPath = GetSetting(gINIFile, "Setup", "LogoPath")
        txtPath.Enabled = 1
        txtPath.SetFocus
    End If
    txtDatabasePath = GetSetting(gINIFile, "Setup", "Database")
End Sub

Private Sub Form_Load()
    Width = gVGAWidth
    Height = gVGAHeight
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub


Private Sub optFormat_Click(Index As Integer)
    If Index = 0 Then txtFooter = gCompany & ", " & gPlace & ".  Printed " & Str$(Date)
    If Index = 1 And gPassword = 1 Then txtFooter = "Printed by " & gUserName & " of " & gCompany & " on " & Str$(Date)
    If Index = 1 And gPassword = 0 Then txtFooter = "Printed by " & gCompany & " on " & Str$(Date)
    If Index = 2 Then
    
        txtFooter = "Printed by me today"
        txtFooter.SelLength = Len(txtFooter.Text)
        txtFooter.SetFocus
    End If
End Sub

Private Sub optLogo_Click(Index As Integer)
    If Index = 0 Then imgLogo.Picture = frmLogin.imgCBB.Picture
    If Index = 1 Then imgLogo.Picture = frmLogin.imgKeys.Picture
    If Index = 2 And txtPath <> "" Then imgLogo.Picture = LoadPicture(txtPath)
    If Index = 2 And txtPath = "" Then
        txtPath = GetSetting(gINIFile, "Setup", "LogoPath")
        txtPath.Enabled = 1
        txtPath.SetFocus
    End If
End Sub
