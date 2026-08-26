Attribute VB_Name = "modMain"
    Public objSession As Object
    Public objMessage As Object
    Public objRecipient As Object
    Public adoConn As ADODB.Connection
    Public adoConn2 As ADODB.Connection
    Public adoData As ADODB.Recordset

Public Sub parse(SourceFilename As String, Server As String)
    Dim tmpfilename
    Dim LogFields(9) As String, SQL As String
    Dim lngRetVal As Long
    Dim processFilename As String
    Dim p As Long

    If SourceFilename = "" Then Exit Sub
    
    lngRetVal = Len(Dir$(SourceFilename))
    If Err Or lngRetVal = 0 Then Exit Sub
    
    p = InStr(p + 1, SourceFilename, ".", vbBinaryCompare)
    tmpfilename = Left$(SourceFilename, p - 1) & "." & "tmp"
    processFilename = Left$(SourceFilename, p - 1) & "." & "csv"
   
    Name SourceFilename As tmpfilename

    record = ""
    recField = 0
    Open tmpfilename For Input As #1
    Open processFilename For Append As #2
        inputdata = Input(1, #1)
        Do While Not EOF(1)
            DoEvents
            frmMain.lblStatus.Refresh
            If inputdata <> "," Then
                fieldData = ""
                If inputdata = "/" Then
                    inputdata = Input(1, #1)
                    Select Case inputdata
                        Case "A", "F", "O", "R", "S", "T", "V", "X"
                            LogFields(recField) = record
                            recField = recField + 1
                            record = ""
                        Case "U"
                            LogFields(recField) = record
                            record = ""
                            recField = 0
                            For lp = 0 To 8
                                record = record & LogFields(lp) & ","
                                LogFields(lp) = ""
                            Next lp
                            frmMain.lblStatus = "Writing .. " & record
                            Write #2, record
                            record = ""
                        Case Else
                            record = record + "/"
                            record = record + inputdata
                    End Select
                ElseIf inputdata <> Chr(13) And inputdata <> Chr(10) Then
                        record = record + inputdata
                End If
            End If
            inputdata = Input(1, #1)
            If inputdata = Chr(13) Or inputdata = Chr(10) Then
                inputdata = Input(1, #1)
            End If
        Loop
    Close 1
    Close 2
End Sub

Public Sub SendMail(SourceServer As String, QuarantineFilename, OriginalFilename, Recipient, Sender, SendDateTime, OriginalSubject)
    DoEvents
    
    'Add a new message object to the OutBox
    Set objMessage = objSession.Outbox.Messages.Add

    ' Get original file and copy to c:\temp as new filename
    ReplaceString = Replace(QuarantineFilename, ":", "$")
    QuarantineFilename = SourceServer & "\" & ReplaceString
    FileCopy "\\" & QuarantineFilename, "c:\temp" & "\" & OriginalFilename & ".qto"
    
    'Set the properties of the message object
    With objMessage ' message object
        .Subject = "Blocked attachment - " & OriginalFilename & " (" & OriginalSubject & ")"
        .Text = "-----Original Message-----" & vbCrLf
        .Text = .Text & "From:             " & Sender & vbCrLf
        .Text = .Text & "Sent:             " & CDate(SendDateTime) & vbCrLf
        .Text = .Text & "To:               " & Recipient & vbCrLf
        .Text = .Text & "Subject:          " & OriginalSubject & vbCrLf
        .Text = .Text & vbCrLf & vbCrLf
        .Text = .Text & "Please save the included attachment and rename it to " & OriginalFilename
        .Text = .Text & " " & vbCrLf ' add placeholder for attachment
        .Text = .Text & "This message was automatically sent by the CSC Quarantine Exceptions application." & vbCrLf
        Set objAttach = .Attachments.Add ' add the attachment
        With objAttach
            .Type = CdoFileData
            .Position = 0 ' render at first character of message
            .Name = OriginalFilename & ".qto"
            .ReadFromFile "\\" & QuarantineFilename
         End With
         objAttach.Name = OriginalFilename & ".qto"
         .Update ' update message to save attachment in MAPI system
    End With
    'Add a recipient object to the objMessage.Recipients collection
    Set objRecipient = objMessage.Recipients.Add

    'Set the properties of the recipient object
    objRecipient.Name = Recipient  '<---Replace this with a valid
                                       'display name or e-mail alias
    objRecipient.Type = mapiTo
    objRecipient.Resolve

    'Send the message
    objMessage.Send showDialog:=False
    
    Kill "c:\temp" & "\" & OriginalFilename & ".qto"
End Sub

