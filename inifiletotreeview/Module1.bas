Attribute VB_Name = "Module1"
Option Explicit

Const MAX_PATH = 255

Private Declare Function GetWindowsDirectory Lib "kernel32" _
Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
ByVal nSize As Long) As Long

Public Function GetWindowsDir() As String

    Dim sRet As String, lngRet As Long
    sRet = String$(MAX_PATH, 0)
    lngRet = GetWindowsDirectory(sRet, MAX_PATH)
    GetWindowsDir = Left(sRet, lngRet)

End Function

Public Function FileExists(Filename As String) As Boolean

    On Error Resume Next
    FileExists = (Dir$(Filename, vbSystem + vbHidden) <> "")

End Function

Public Sub AutoSelectText(EditControl As Object)

On Error GoTo ErrorHandler
Dim lTextLength As Long

    lTextLength = Len(EditControl.Text)
    If lTextLength Then
        EditControl.SelStart = 0
        EditControl.SelLength = lTextLength
    End If

ErrorHandler:

    Exit Sub

End Sub
