VERSION 5.00
Begin VB.Form frmRegDemo 
   Caption         =   "Registry Demo"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete All Values"
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Frame fraByte 
      Caption         =   "Byte Data"
      Height          =   1815
      Left            =   1920
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtByte 
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtByte 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtByte 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveByte 
         Caption         =   "&Save"
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdGetByte 
         Caption         =   "&Retrieve"
         Height          =   615
         Left            =   2400
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Frame fraLong 
      Caption         =   "Long Data"
      Height          =   1815
      Left            =   1920
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtLong 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Text            =   "0"
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton cmdSaveLong 
         Caption         =   "&Save"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdGetLong 
         Caption         =   "&Retrieve"
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.OptionButton optByte 
      Caption         =   "&Byte (Binary)"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton optLong 
      Caption         =   "&Long (DWord)"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton optString 
      Caption         =   "S&tring"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame fraString 
      Caption         =   "String Data"
      Height          =   1815
      Left            =   1920
      TabIndex        =   15
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdGetString 
         Caption         =   "&Retrieve"
         Height          =   615
         Left            =   2400
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdSaveString 
         Caption         =   "&Save"
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtString 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Label lblinfo1 
      AutoSize        =   -1  'True
      Caption         =   "For more demonstration Visual Basic Projects, please visit:"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   4080
   End
   Begin VB.Label lblurl 
      AutoSize        =   -1  'True
      Caption         =   "http://www.vb-world.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lblemail 
      AutoSize        =   -1  'True
      Caption         =   "john@vb-world.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblinfo2 
      AutoSize        =   -1  'True
      Caption         =   "To contact us, please send email to:"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   2565
   End
End
Attribute VB_Name = "frmRegDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdDelete_Click()

DeleteKey HKEY_CURRENT_USER, "Software\VBW\RegistryDemo"

End Sub

Private Sub cmdGetByte_Click()
Dim byTemp As Variant
Dim byDefault(2) As Byte

byDefault(0) = 0
byDefault(1) = 0
byDefault(2) = 0

byTemp = GetSettingByte(HKEY_CURRENT_USER, "Software\VBW\RegistryDemo", "Byte", byDefault)

txtByte(0) = CStr(byTemp(0))
txtByte(1) = CStr(byTemp(1))
txtByte(2) = CStr(byTemp(2))
End Sub

Private Sub cmdGetLong_Click()

txtLong = CStr(GetSettingLong(HKEY_CURRENT_USER, "Software\VBW\RegistryDemo", "Long", 0))

End Sub

Private Sub cmdGetString_Click()

txtString = GetSettingString(HKEY_CURRENT_USER, "Software\VBW\RegistryDemo", "String", "")

End Sub

Private Sub cmdSaveByte_Click()
Dim byTemp(2) As Byte

byTemp(0) = CByte(txtByte(0))
byTemp(1) = CByte(txtByte(1))
byTemp(2) = CByte(txtByte(2))

SaveSettingByte HKEY_CURRENT_USER, "Software\VBW\RegistryDemo", "Byte", byTemp

End Sub

Private Sub cmdSaveLong_Click()

SaveSettingLong HKEY_CURRENT_USER, "Software\VBW\RegistryDemo", "Long", CLng(txtLong)

End Sub

Private Sub cmdSaveString_Click()

SaveSettingString HKEY_CURRENT_USER, "Software\VBW\RegistryDemo", "String", txtString

End Sub

Private Sub Form_Load()
lblemail = email
lblurl = URL
End Sub

Private Sub lblemail_Click()
sendemail
End Sub

Private Sub lblurl_Click()
gotoweb
End Sub

Private Sub optByte_Click()
fraByte.Visible = True
fraLong.Visible = False
fraString.Visible = False
End Sub

Private Sub optLong_Click()
fraByte.Visible = False
fraLong.Visible = True
fraString.Visible = False
End Sub

Private Sub optString_Click()
fraByte.Visible = False
fraLong.Visible = False
fraString.Visible = True
End Sub
