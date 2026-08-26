Attribute VB_Name = "Module1"
    Public objSession As Object
    Public objMessage As Object
    Public objRecipient As Object
    Public adoConn As ADODB.Connection
    Public adoConn2 As ADODB.Connection
    Public adoData As ADODB.Recordset
    Public DSN

Public Sub Parse(SourceFilename As String, Server As String)
    Dim LogFields(9) As String, SQL As String
    Record = ""
    adoConn.Open DSN
    Open SourceFilename For Input As #1
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
                            recField = recField + 1
                            LogFields(recField) = Record
                            Record = ""
                        Case "U"
                            fieldData = Record
                            LogFields(9) = Record
                            ' Check if sender or recipient is in exemptions list
                            frmMain.lblStatus = "Checking '" & LogFields(9) & "' for exempt sender or recipient. (" & LogFields(3) & " or " & LogFields(2) & ")"
                            frmMain.lblStatus.Refresh
                            If Exempt(LogFields(3), LogFields(2), LogFields(3), LogFields(9)) Or Exempt(LogFields(2), LogFields(2), LogFields(3), LogFields(9)) Then
                                If LogFields(6) <> "" And LogFields(7) <> "" Then
                                    ' Fields are:  scanneddatetime,sender,recipient,senddatetime,virus,attachmentname,quarantinedir,x,subject
                                    SendMail Server, LogFields(7), LogFields(6), LogFields(3), LogFields(2), LogFields(4), LogFields(9)
                                    SQL = "INSERT INTO tblBlockingResends (Sender, Recipient, Subject) VALUES (" & LogFields(2) & "," & LogFields(3) & "," & LogFields(9) & ")"
                                    adoData.Open SQL, adoConn
                                    'adoConn.Close
                                End If
                            End If
                            Record = ""
                            recField = 0
                            For lp = 1 To 9
                                LogFields(lp) = ""
                            Next lp
                        Case Else
                            Record = Record + "/"
                            Record = Record + inputdata
                    End Select
                ElseIf inputdata <> Chr(13) And inputdata <> Chr(10) Then
                        Record = Record + inputdata
                End If
            End If
            inputdata = Input(1, #1)
            If inputdata = Chr(13) Or inputdata = Chr(10) Then
                inputdata = Input(1, #1)
            End If
        Loop
    Close 1
End Sub

Public Sub SendMail(SourceServer As String, QuarantineFilename As String, OriginalFilename As String, Recipient As String, Sender As String, SendDateTime, OriginalSubject As String)
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

Function Exempt(ExchangeDisplayName As String, Sender As String, Recipient As String, Subject As String) As Boolean
    adoConn2.Open DSN
    SearchName = Replace(ExchangeDisplayName, "'", "''")
    SearchName = Replace(SearchName, "[", "_")
    SearchName = Replace(SearchName, "]", "_")
    SQL = "SELECT * FROM tblBlockingExemptions WHERE DisplayName like '" & SearchName & "'"
    adoData.Open SQL, adoConn
    Exempt = False
    If Not adoData.EOF Then
        If adoData.Fields(0) <> "" Then Exempt = True
        'adoConn.Close
        SQL = "SELECT * FROM tblBlockingResends WHERE Sender like '" & Sender & "' AND Recipient like '" & Recipient & "' AND Subject like '" & Subject & "'"
        adoData.Open SQL, adoConn
        If Not adoData.EOF Then
            If adoData.Fields(0) <> "" Then Exempt = False
        End If
    End If
    adoConn.Close
End Function
