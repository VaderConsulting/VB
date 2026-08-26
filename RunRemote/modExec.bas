Attribute VB_Name = "modExec"

Option Explicit

Private Const CONST_EXECUTE As Integer = 4
Private Const CONST_KILL As Integer = 5

'Declare variables
Private intProcessID As Integer
Private strCommand As String

Public Sub ExecuteCmd(strServer As String, strCommand As String)
    Dim objFileSystem As Object, objService As Object, objInstance As Object
    Dim strQuery As String, strMessage As String
    Dim intProcessID As Integer, intStatus As Integer
    
    On Error Resume Next
    
        Call HTMLHeaders

        'Establish a connection with the server.
        If blnConnect("root\cimv2", strUserName, strPassword, strServer, objService) Then
            'Call document.write("")
            'Call document.write("Please check the server name, " & "credentials and WBEM Core.")
            Exit Sub '>---> Bottom
        End If

        strMessage = ""
        intProcessID = 0

        Set objInstance = objService.Get("Win32_Process")
        If blnErrorOccurred(" occurred getting a " & " Win32_Process class object.") Then Exit Sub ':( Expand Structure or consider reversing Condition

        If objInstance Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition

        intStatus = objInstance.Create(strCommand, Null, Null, intProcessID)
        If blnErrorOccurred(" occurred in creating process " & strCommand & ".") Then Exit Sub ':( Expand Structure or consider reversing Condition

        If intStatus = 0 Then
            If intProcessID < 0 Then
                '4294967296 is 0x100000000.
                intProcessID = intProcessID + 4294967296#
            End If
            strMessage = "Succeeded in executing " & strCommand & "." & vbCrLf
            strMessage = strMessage & "The process id is " & intProcessID & "."
        Else 'NOT INTSTATUS...
            strMessage = "Failed to execute " & strCommand & "." & vbCrLf
            strMessage = strMessage & "Status = " & intStatus
        End If
        WriteLine strMessage
    On Error GoTo 0
End Sub

Public Sub Kill(strServer As String, intProcessID As Integer)
    Dim objFileSystem As Object, objService As Object, objInstance As Object
    Dim strWBEMClass As String, strMessage As String
    Dim intStatus As Integer
    
    On Error Resume Next
    
        'Establish a connection with the server.
        If blnConnect("root\cimv2", strUserName, strPassword, strServer, objService) Then
            'Call document.write("")
            'Call document.write("Please check the server name, " & "credentials and WBEM Core.")
            Exit Sub '>---> Bottom
        End If

        'Now executes the method.
        If strServer = "" Then
            strWBEMClass = "Win32_Process.Handle=" & intProcessID
        Else 'NOT STRSERVER...
            strWBEMClass = "\\" & strServer & "\root\cimv2:Win32_Process.Handle=" & intProcessID
        End If

        Set objInstance = objService.Get(strWBEMClass)
        If blnErrorOccurred(" occurred in getting process " & strWBEMClass & ".") Then Exit Sub ':( Expand Structure or consider reversing Condition

        intStatus = objInstance.Terminate

        If intStatus = 0 Then
            strMessage = "Process " & intProcessID & " has been killed."
        Else 'NOT INTSTATUS...
            strMessage = "Failed to kill process " & intProcessID & "."
        End If

        WriteLine strMessage
    On Error GoTo 0
End Sub
