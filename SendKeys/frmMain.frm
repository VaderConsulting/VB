VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Citrix"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim strAppName As String
    
    strAppName = "c:\Program Files\Citrix\ICA Client\pn.exe"
    
    If strAppName <> "" Then
        Shell strAppName, vbNormalFocus
        'Pause 10
        'VbSendKeys "wopur2"
        'VbSendKeys "{TAB}"
        'VbSendKeys "woodside7"
        'VbSendKeys "{TAB}"
        'VbSendKeys "woodside"
        'VbSendKeys "{ENTER}"
        ' Pause 5 seconds
        Pause 5
        VbSendKeys "{ALT}f"
        VbSendKeys "a"
        VbSendKeys "{SHIFT}{TAB}"
        VbSendKeys "{RIGHT}{RIGHT}"
        VbSendKeys "{TAB}"
        VbSendKeys "{UP}{UP}{UP}{UP}"
        VbSendKeys "{TAB}"
        VbSendKeys "{ENTER}"
        Pause 2
        VbSendKeys "y"
        Pause 6
        VbSendKeys "wopur2"
        VbSendKeys "{TAB}"
        VbSendKeys "woodside7"
        VbSendKeys "{TAB}"
        VbSendKeys "woodside"
        VbSendKeys "{ENTER}"
        Pause 5
        VbSendKeys "{ALT}f"
        VbSendKeys "c"
    End If
End Sub

Public Sub Pause(iSeconds As Integer)
    Dim d As Date
    
    d = Now
    
    Do
        DoEvents
    Loop Until Now > DateAdd("s", CDbl(iSeconds), d)
End Sub
