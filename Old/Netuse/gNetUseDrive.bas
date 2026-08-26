Attribute VB_Name = "gNetUseDrive"
Option Explicit

Sub Main()

Dim NetUse As New ClassNetUse

Call NetUse.Connect("K:", "\\server01\abc")
If NetUse.rc <> 0 Then
    MsgBox NetUse.ErrorMsg
End If

Call NetUse.DisConnect("K:")

If NetUse.rc <> 0 Then
    MsgBox NetUse.ErrorMsg
End If

Set NetUse = Nothing

End Sub
