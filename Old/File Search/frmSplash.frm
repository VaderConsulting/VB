VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4240
   ClientLeft      =   256
   ClientTop       =   1408
   ClientWidth     =   7376
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4240
   ScaleWidth      =   7376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   5000
      Left            =   0
      Top             =   3584
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright '98"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "D. Robinson"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning - Use of this program is at the users risk.  No guarantees are implied or given for the correctness of any data presented."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   448
         Left            =   144
         TabIndex        =   2
         Top             =   3536
         Width           =   6848
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.28
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   5
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "for Windows NT, '95 and '98"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.55
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   352
         Left            =   2784
         TabIndex        =   6
         Top             =   2336
         Width           =   4080
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   31.91
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2528
         TabIndex        =   8
         Top             =   1136
         Width           =   4256
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Licensed to HMAS ARUNTA - Royal Australian Navy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Chips Bits and Bytes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   416
         Left            =   2352
         TabIndex        =   7
         Top             =   704
         Width           =   3408
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    lblProductName.Caption = "File Search"
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub tmrSplash_Timer()
    Unload Me
End Sub
