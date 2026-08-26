VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Linker"
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1260
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   1260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    Dim strFilename As String
    Dim strWorkingDir As String
    
    If Command$ = "" Then
        End
    Else
        strWorkingDir = Environ$("ALLUSERSPROFILE")
        strFilename = Command$
        strFilename = Replace(strFilename, Chr(34), "")
        If Dir(strFilename, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then
            On Error Resume Next
                Call ShellExecute(frmMain.hwnd, "Open", Chr(34) & strFilename & Chr(34), "", strWorkingDir, 1)
            On Error GoTo 0
        End If
        End
    End If
End Sub
