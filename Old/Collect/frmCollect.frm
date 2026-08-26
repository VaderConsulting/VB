VERSION 5.00
Begin VB.Form frmCollect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collect"
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmCollect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sending your computers asset information back to server..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmCollect.Visible = True
    Label1.Visible = True
    Open "c:\temp\collect.bat" For Output As #1
        Print #1, "copy %logonserver%\netlogon\setup\collect.exe c:\temp\collect.exe"
        Print #1, "c:\temp\kix32.exe %logonserver%\netlogon\collect.kix"
    Close 1
    retval = Shell("c:\temp\collect.bat")
    frmCollect.Visible = False
    MsgBox ("Thank you for your Co-Operation")
    Unload Me
End Sub
