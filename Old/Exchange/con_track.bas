Attribute VB_Name = "modTrack"
' The original procedure to do this was provided by Brian Feige.
' VB is used to provide a constant, efficient and standard way of extracting the logs
' but in a compiled manner
' End Blurb
' :-)


Sub Main()
    Dim Source As String
    Dim Destination As String
    Dim Mode As String
    cmdline = Command$
    If cmdline = "" Then
        frmMain.Show
    Else
        pos1 = InStr(1, cmdline, "/s:") + 3
        pos2 = InStr(1, cmdline, "/d:") + 3
        pos3 = InStr(1, cmdline, "/m:") + 3
        If pos1 > 3 And pos2 > 3 Then
            Source = Mid(cmdline, pos1, pos2 - 8)
            
            ' set mode switch as appropriate
            If pos3 > 3 Then
                Destination = Trim(Mid(cmdline, pos2, pos3 - pos2 - 3))
                Mode = Mid(cmdline, pos3, 200)
            Else
                Destination = Trim(Mid(cmdline, pos2, pos2 - 3))
                Mode = "all"
            End If
            
            ' Call Conversion using Source log file, Destination CSV file, mode
            ' If mode = "" then all is assumed
            ' If mode = "all" then all data is extracted
            ' If mode = "sub" then only that data that is pertinent to tracking
            '     messages from users is extracted
            
            Call Extract(Source, Destination, Mode)
            'eg Convert_tracking_data(\\exhsrv1\c$\exchsrvr\tracking.log\20000708.log, c:\temp\exhsrv1.csv, all)
        End If
    End If
End Sub

Sub DR_Extract(Source As String, Destination As String, Mode As String)
    record = ""
    recipient = ""
    reccnt = 0
    
End Sub

Public Sub Extract(Source As String, Destination As String, Mode As String)
    On Error GoTo ErrorRoutine
    record = ""
    recipient = ""
    reccnt = 0
    Open Source For Input As #1
        Open Destination For Append As #2
        flag = 0
        If Mode = "all" Then
            Print #2, "Message-ID,?,Date/Time,?,Connector,?,Originator,?,Size,?,?,?,?,Recipient"
        Else
            Print #2, "Message-ID,Date/Time,Connector,Originator,Size,Recipient"
        End If
        Do While Not EOF(1)
            DoEvents
            inputdata = Input(1, #1)
            If Asc(inputdata) = 10 Then
                If flag = 1 Then
                     Print #2, record; ; ","; recipient
                     recipient = ""
                Else
                    flag = 1
                End If
                inputdata = Input(1, #1)
                If Asc(inputdata) = 10 Then   ' CR character
                    record = ""
                    recipient = ""
                    flag = 0
                End If
            End If
            If inputdata = Chr(9) Then        ' TAB character
                inputdata = ","
            End If
            If flag = 0 Then
                If inputdata <> Chr(10) Then  ' CR character
                    record = record + inputdata
                End If
            Else
                recipient = recipient + inputdata
            End If
        Loop
        Close #1
    Close #2
    Exit Sub
ErrorRoutine:
    errorNumber = Err.Number
    msg = "Error during extract process:" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & errorNumber & " (" & Error(errorNumber) & ")" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & " Source file: " & Source & vbCrLf
    msg = msg & " Destination file: " & Destination & vbCrLf
    msg = msg & vbCrLf
    msg = msg & " Debug info follows:" & vbCrLf
    msg = msg & " Mode                  = " & Mode & vbCrLf
    msg = msg & " Record Count (reccnt) = " & reccnt & vbCrLf
    msg = msg & " Variables:" & vbCrLf
    msg = msg & "          InputData    = " & inputdata & vbCrLf
    msg = msg & "          Record       = " & record & vbCrLf
    msg = msg & "          Recipient    = " & recipient & vbCrLf
    msg = msg & "          Flag         = " & flag & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "The extract process will continue from the next file in the list." & vbCrLf
    msg = msg & "Click the OK button to continue this process" & vbCrLf
    MsgBox msg, vbCritical + vbOKOnly, "Critical Error"
    Close 2
    Close 1
End Sub
