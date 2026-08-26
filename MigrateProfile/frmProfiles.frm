VERSION 5.00
Begin VB.Form frmProfiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile selection"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK2 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   2760
      Width           =   495
   End
   Begin VB.DirListBox dirProfiles 
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   4455
   End
End
Attribute VB_Name = "frmProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strNTDomainProfileDir As String
Private strADDomainProfileDir As String

Private Sub cmdOK1_Click()
    strADDomainProfileDir = dirProfiles.List(dirProfiles.ListIndex)
    
    cmdOK1.Visible = False
    cmdOK2.Visible = True
    
    lblInfo.Caption = "Select the directory for this user without an extension."
    dirProfiles.Path = "C:\Documents and Settings"
End Sub

Private Sub cmdOK2_Click()
    Dim Temp As String
    
    strNTDomainProfileDir = dirProfiles.List(dirProfiles.ListIndex)
    
    ' Rename APAC profile dir to .old
    strTemp = Left(strADDomainProfileDir, InStr(1, strADDomainProfileDir, ".") - 1)
    Name strADDomainProfileDir As strTemp & ".old"
    
    ' Rename BHP_AU for APAC use
    Name strNTDomainProfileDir As strADDomainProfileDir
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strData As String
    Dim strUsername As String               ' Username of the user being migrated, NOT the current user name
    
    lblInfo.Caption = "Select the directory for this user with ." & strDomainName & " as the extension"
    dirProfiles.Path = "C:\Documents and Settings"
    ' Get previous profile name
    cmdOK2.Visible = False
    
    
End Sub
