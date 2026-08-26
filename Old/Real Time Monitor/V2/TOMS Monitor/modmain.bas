Attribute VB_Name = "Module1"

Sub Main()
    ' TWO ways to do this:
    
    ' 1.  Lame commandline way
    'Dim cmdLine As String, d As String, stuph As String
    'cmdLine = ""
    'cmdLine = cmdLine & "echo ok to delete me >running.dat" & vbCrLf
    'cmdLine = cmdLine & "TNSPING.exe mojsc1tomsprod >results.dat" & vbCrLf
    'cmdLine = cmdLine & "del running.dat"
    'Open App.Path & "\tomscheck.bat" For Output As #1
    '    Print #1, cmdLine
    'Close 1
    'Shell "cmd.exe /c " & App.Path & "\tomscheck.bat"
    'd = Dir(App.Path & "\running.dat")
    'Do Until d = ""
    '    DoEvents
    '    d = Dir
    'Loop
    'Open App.Path & "\results.dat" For Input As #1
    '    Do Until EOF(1)
    '        Line Input #1, stuph
    '        MsgBox stuph
    '    Loop
    'Close 1
    
    ' 2.  Cool Sockets way
    frmMain.Show
End Sub
