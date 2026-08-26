VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents sh As cSysHook
Attribute sh.VB_VarHelpID = -1

'constants
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_SYSTEMKEYDOWN = &H104
Private Const WM_SYSTEMKEYUP = &H105

Private Sub Form_Load()
    Set sh = New cSysHook
    sh.SetHook
End Sub
    
Private Sub Form_Unload(Cancel As Integer)
    sh.RemoveHook
    Set sh = Nothing
End Sub
    
Private Sub sh_msgout(hwnd As Long, imsg As Long, LP As Long, HP As Long, itime As Long)
    Select Case imsg
        Case WM_RBUTTONUP
            Me.Caption = "rightmouse pressed on hwnd " & CStr(hwnd)
    End Select
End Sub


