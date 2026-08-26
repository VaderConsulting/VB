Attribute VB_Name = "Module1"
Option Explicit

Function TFileCopy(FileFrom, FileTo)
    Dim StrBuffer As String
    Dim i As Integer
    
    Open FileFrom For Binary As #1
    Open FileTo For Binary As #2
    StrBuffer = Space(1024)
    
    For i = 1 To LOF(1) \ 1024
    
        Get #1, , StrBuffer
        Put #2, , StrBuffer
        
        'If ProgressBar1.Value = 100 Then ProgressBar1.Value = 0
        'ProgressBar1.Value = ProgressBar1.Value + 1
    
    Next i
    
    If LOF(1) Mod 1024 > 0 Then
        StrBuffer = Space(LOF(1) Mod 1024)
        Get #1, , StrBuffer
        Put #2, , StrBuffer
        
    End If
    Close #1, #2: i = 0
    'ProgressBar1.Value = 0: MsgBox "FileCopy Done"
    
End Function



