VERSION 5.00
Begin VB.Form frmClick 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hello World"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click me"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Sub cmdClick_Click()
    Dim retvalue As Long
    retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim retvalue As Long
    retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)
    Set frmClick = Nothing
End Sub
