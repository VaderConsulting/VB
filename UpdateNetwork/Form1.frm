VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Network"
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2250
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   2250
   StartUpPosition =   1  'CenterOwner
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim lngReturnValue As Long
    SendMessageTimeout HWND_BROADCAST, WM_SETTINGCHANGE, 0&, "Environment", SMTO_ABORTIFHUNG, 5000&, lngReturnValue
    End
End Sub
