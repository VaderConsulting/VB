VERSION 5.00
Begin VB.Form FrmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menus"
   ClientHeight    =   1005
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      Begin VB.Menu MenuShowMain 
         Caption         =   "Show Main Form"
      End
      Begin VB.Menu MenuDesk 
         Caption         =   "Desk Top Form"
         Begin VB.Menu MenuShowDesk 
            Caption         =   "Show Desk Top Form"
         End
         Begin VB.Menu Menuhidedesktop 
            Caption         =   "Hide Desk Top Form"
         End
      End
      Begin VB.Menu menuline3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSettings 
         Caption         =   "Settings..."
      End
      Begin VB.Menu menuline2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu menuline1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MenuAbout_Click()
frmAbout.Show
End Sub

Private Sub MenuExit_Click()
If MsgBox("This will close the program, Continue?", vbYesNo) = vbYes Then
Call FrmMain.CleanUpSystray
DoEvents
End
End If
End Sub

Private Sub Menuhidedesktop_Click()
Unload FrmDesktop
DoEvents
End Sub

Private Sub MenuSettings_Click()
FrmSettings.Show
End Sub

Private Sub MenuShowDesk_Click()
FrmDesktop.Show
DoEvents
End Sub

Private Sub MenuShowMain_Click()
FrmMain.Show
End Sub
