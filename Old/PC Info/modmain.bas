Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
    frmMain.Show
End Sub

Public Function isUP(Computername As String) As Boolean
    Dim objWMIService As Object, colPingedComputers As Object
    Dim Computer As Object
    
    On Error GoTo Hell
    
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colPingedComputers = objWMIService.ExecQuery("Select * from Win32_PingStatus Where Address = '" & Computername & "'")
    
    For Each Computer In colPingedComputers
        If Computer.StatusCode = 0 Then
            isUP = True
        Else
            isUP = False
       End If
    Next
    
    Set objWMIService = Nothing
    Set colPingedComputers = Nothing
    Exit Function
Hell:
    frmMain.LogMessage "Error " & Err.Number & " (" & Err.Description & ")", 1
    Err.Clear
    Resume Next
End Function
