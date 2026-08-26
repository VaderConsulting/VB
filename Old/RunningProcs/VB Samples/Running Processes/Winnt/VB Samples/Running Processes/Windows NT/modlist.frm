VERSION 5.00
Begin VB.Form frmModuleList 
   Caption         =   "Process/Module List"
   ClientHeight    =   3915
   ClientLeft      =   5685
   ClientTop       =   3000
   ClientWidth     =   4560
   Icon            =   "modlist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      Height          =   330
      Left            =   3570
      TabIndex        =   2
      Top             =   105
      Width           =   750
   End
   Begin VB.ListBox List1 
      Height          =   3195
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   525
      Width           =   4215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   330
      Left            =   2730
      TabIndex        =   0
      Top             =   105
      Width           =   750
   End
   Begin VB.Label lblOS 
      Caption         =   "Operating System:"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   105
      Width           =   2115
   End
End
Attribute VB_Name = "frmModuleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
'dss delete this button
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
Screen.MousePointer = vbHourglass
'
' See which version of windows is running
' and call the appropriate function.
'
Select Case GetVersion()
    Case VER_PLATFORM_WIN32_WINDOWS
        lblOS.Caption = "Operating System: Win 9x"
        Call FillProcessList9x(List1)
    Case VER_PLATFORM_WIN32_NT
        lblOS.Caption = "Operating System: Win NT"
        Call FillProcessListNT(List1)
    Case Else
        MsgBox "Operating system not recognized!", vbCritical
End Select
    
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'
' Let the operating system process other
' events so the form can refresh.
'
DoEvents
'
' Fill the list for the first time.
'
cmdRefresh_Click
End Sub

Private Sub Form_Resize()
'
' Resize the Listbox.
'
On Error Resume Next
List1.Height = Me.ScaleHeight - (List1.Top)
List1.Width = Me.ScaleWidth - (List1.Left * 2)
End Sub
