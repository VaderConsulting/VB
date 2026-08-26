VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regini"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Cmd As String
    Cmd = Command
    
    If Trim(Cmd) = "" Then
        MsgBox "Usage:" & vbCrLf & "Runs regini.exe with the supplied command-line.", vbInformation, "Help"
    Else
        Cmd = Replace(Cmd, "HKLM", "\Registry\Machine")
        Cmd = Replace(Cmd, "HKEY_LOCAL_MACHINE", "\Registry\Machine")
        Cmd = Replace(Cmd, "HKU", "\Registry\Users")
        Cmd = Replace(Cmd, "HK_USERS", "\Registry\Users")
        Cmd = Replace(Cmd, "HKEY_USERS", "\Registry\Users")
        Shell "Regini.exe " & Cmd
    End If
    'Debug.Print "Regini " & Cmd
    End
End Sub
