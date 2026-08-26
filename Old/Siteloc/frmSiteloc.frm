VERSION 5.00
Begin VB.Form frmSiteloc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Siteloc detect"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2610
   Icon            =   "frmSiteloc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   2610
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSiteloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    siteloc = Environ("siteloc")
    If siteloc = "" Then
        MsgBox "There is a problem with the SITELOC variable on this computer.  Please contact the helpdesk for assistance on 1800 651 327", vbCritical + vbOKOnly, "Logon"
    End If
    End
End Sub
