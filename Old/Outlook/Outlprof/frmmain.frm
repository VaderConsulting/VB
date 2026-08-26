VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Outlook Profile Switcher"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    HomeShare = Environ$("HOMESHARE")
    SlashPos = InStr(3, HomeShare, "\")
    ServerName = UCase(Left(HomeShare, SlashPos - 1))
    SiteServer = UCase(Environ$("SiteServer"))
    If Dir("c:\temp\outlook.prf") <> "" Then
        Kill "c:\temp\outlook.prf"
    End If
    If ServerName = SiteServer Then
        FileCopy "c:\temp\home.prf", "c:\temp\outlook.prf"
    Else
        FileCopy "c:\temp\travel.prf", "c:\temp\outlook.prf"
    End If
    End
End Sub
