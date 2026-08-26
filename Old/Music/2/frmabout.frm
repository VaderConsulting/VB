VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3945
   ClientLeft      =   1410
   ClientTop       =   3015
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3945
   ScaleWidth      =   5790
   Begin VB.CommandButton cmdAuthor 
      Caption         =   "Author info"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdThanks 
      Cancel          =   -1  'True
      Caption         =   "Thanks"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Image imgLogo 
      Height          =   855
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblMemory 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   5550
   End
   Begin VB.Label lblPlace 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Label lblRegName 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label lblRegistered 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAuthor_Click()
    subAuthor
End Sub

Private Sub cmdRegister_Click()
    frmAbout.Hide
    frmStart.Hide
    frmRegister.Show
End Sub

Private Sub cmdThanks_Click()
    frmStart.Show
    frmAbout.Hide
End Sub


Private Sub Form_Activate()
    Dim Resource(2) As Integer
    subGetRegoInfo
    If gRegistered = 0 Then
        cmdRegister.Visible = 1
        lblRegistered = "Please register this great program."
        lblRegName = "Unregistered and owned by " & gBusiness
        lblPlace = gHome
    End If
    If gRegistered = 1 Then
        cmdRegister.Visible = 0
        lblRegistered = "Thank you for buying this great program."
        lblRegName = "Registered to " & gName & " " & gSurname
        lblPlace = gPlace
    End If
    Rem Resource(0) = GetFreeSystemResources%(0)
    Rem Resource(1) = GetFreeSystemResources%(1)
    Rem lblMemory = "Free Memory: User " & Resource(0) & "% GDI " & Resource(1) & "%"
    If gLogo = 0 Then imgLogo.Picture = frmLogin.imgCBB.Picture
    If gLogo = 1 Then imgLogo.Picture = frmLogin.imgKeys.Picture
    If gLogo = 2 And gLogoPath <> "" Then imgLogo.Picture = LoadPicture(txtPath)
End Sub

Private Sub Form_Load()
    Width = 5910
    Height = 4350
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
    lblName = gProgramName
    lblAuthor = gAuthor
    lblDate = "Copyright " & gDate
    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub



