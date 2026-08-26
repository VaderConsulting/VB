VERSION 5.00
Object = "{34052989-CA79-44A1-8E31-31A6F14B21F6}#1.0#0"; "SetACL.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SetRegPerms"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPermission 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "Full"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "Users"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "hklm\Software\Microsoft\Shared Tools\Graphics Filters"
      Top             =   120
      Width           =   4095
   End
   Begin SETACLLib.SetACL oACL 
      Left            =   480
      Top             =   1680
      _Version        =   65536
      _ExtentX        =   926
      _ExtentY        =   926
      _StockProps     =   0
   End
   Begin VB.Label lblPermission 
      Caption         =   "Permission"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblPath 
      Caption         =   "Registry Path:"
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
' http://setacl.sourceforge.net/html/doc-reference.html

Private Sub cmdSet_Click()
    Dim strError As String
    
    On Error Resume Next
    With oACL
        .SetObject Me.txtPath.Text, SE_REGISTRY_KEY
        .SetAction ACTN_ADDACE
        .AddACE txtUsername.Text, False, txtPermission.Text, INHPARNOCHANGE, False, GRANT_ACCESS, ACL_DACL
        .Run
        strError = .GetLastAPIErrorMessage
    End With
    
    MsgBox strError, vbInformation, "Complete"

    Me.cmdSet.Enabled = False
End Sub
