VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShellExecute Demo Project"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open this File with the default program"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblinfo1 
      AutoSize        =   -1  'True
      Caption         =   "For more demonstration Visual Basic Projects, please visit:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1080
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
      TabIndex        =   5
      Top             =   1320
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
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblinfo2 
      AutoSize        =   -1  'True
      Caption         =   "To contact us, please send email to:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2565
   End
   Begin VB.Label lblFile 
      Caption         =   "&File to Open"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Note:  This code is great for opening
'web document into the default
'browser.  To open http://www.vb-world.net
'into the default browser the following
'code would be used:
'
'Call ShellExecute(hwnd,"Open","http://www.vb-world.net","",app.path,1)

'For contacting information see module

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOpen_Click()

'check that the file in the text box exists
If Dir(txtFile) = "" Then
    Call MsgBox("The file in the text box does not exist.", vbExclamation)
    Exit Sub
End If

'open the file with the default program
Call ShellExecute(hwnd, "Open", txtFile, "", App.Path, 1)

'Note:  This is the equivalent of
'right clicking on a file in Windows
'and selecting "Open"
'
'If you would like to do something
'else to the file rather than opening
'it, right click on a file and see
'what options are in the menu.  Then
'change the "Open" in the code above to
'read what the menu item says.

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

