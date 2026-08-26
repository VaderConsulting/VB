VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Package Configuration"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   554
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PakConfig.xptopbuttons xptopbuttons3 
      Height          =   315
      Left            =   7200
      Top             =   45
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      Value           =   2
   End
   Begin PakConfig.xptopbuttons xptopbuttons2 
      Height          =   315
      Left            =   7560
      Top             =   45
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      Value           =   1
   End
   Begin PakConfig.xptopbuttons xptopbuttons1 
      Height          =   315
      Left            =   7920
      Top             =   45
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin VB.PictureBox pictxt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10920
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   10
      Top             =   240
      Width           =   975
      Begin VB.TextBox txt1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Text            =   "xptextbox"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Computer"
      Height          =   1215
      Left            =   4080
      TabIndex        =   4
      Top             =   1800
      Width           =   4095
      Begin PakConfig.xpcmdbutton cmdBrowseComputer 
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtComputername 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
      Begin PakConfig.xpcmdbutton cmdAddComputer 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "Add"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "User"
      Height          =   1215
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   4095
      Begin PakConfig.xpcmdbutton cmdBrowseUser 
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
      Begin PakConfig.xpcmdbutton cmdAddUser 
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "Add"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Packages"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      Begin VB.ListBox lstPackages 
         Height          =   5910
         ItemData        =   "frmMain.frx":030A
         Left            =   240
         List            =   "frmMain.frx":030C
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
   Begin PakConfig.xp_canvas xp_canvas1 
      Height          =   7095
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   12515
      Caption         =   "Package Configuration"
      Icon            =   "frmMain.frx":030E
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oElastic As New clsResize
Dim bDoneInitialSize As Boolean

Private Sub cmdBrowseComputer_Click()
    frmBrowse.lblItemType = "Computer"
    frmBrowse.Show vbModal
End Sub

Private Sub cmdBrowseUser_Click()
    frmBrowse.lblItemType = "User"
    frmBrowse.Show vbModal
End Sub

Private Sub Form_Activate()
    XPStyle True, True
    xptxt frmMain.txtComputername, pictxt, RGB(240, 232, 224), Normal
    xptxt frmMain.txtUsername, pictxt, RGB(240, 232, 224), Normal
    bDoneInitialSize = True
End Sub

Private Sub Form_Load()
    
    oElastic.Init Me, 0
End Sub

Private Sub Form_Resize()
    If bDoneInitialSize Then
        oElastic.FormResize Me
    End If
End Sub


Private Sub xptopbuttons1_Click()
    End
End Sub

Private Sub xptopbuttons2_Click()
    Me.WindowState = vbMaximized
End Sub

Private Sub xptopbuttons3_Click()
    Me.WindowState = vbMinimized
End Sub
