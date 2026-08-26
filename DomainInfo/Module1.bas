Attribute VB_Name = "Module1"
Public Sub GetDomainObjects(sDomainName As String, Optional Users As Boolean = False, Optional Groups As Boolean = False, Optional Computers As Boolean = False)
    Dim oDomain As IADsContainer, sOutput As String
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Er
    
    Set oDomain = GetObject("WinNT://" & sDomainName & ",domain")
    If Users Then oDomain.Filter = Array("User")
    If Groups Then oDomain.Filter = Array("Group")
    If Computers Then oDomain.Filter = Array("Computer")
    
    Form1.lstObjects.Clear
    
    For Each ODomainObject In oDomain
        sOutput = ODomainObject.Name
        Form1.lstObjects.AddItem sOutput
        Form1.lblCount.Caption = Form1.lstObjects.ListCount & " listed"
    Next
    
    Set oDomain = Nothing
    Form1.lblInfo.Caption = "Complete"
    Screen.MousePointer = vbDefault
    Exit Sub
Er:
    ' There was an unexpected error!
    Screen.MousePointer = vbDefault
    Form1.lblInfo.Caption = "Error " & Err.Number & " (" & Err.Description & ")."
    Err.Clear
End Sub

Public Sub GetGroupMembership(Objectname As String)
    Dim oObject As IADsUser
    Dim i As Integer, isListed As Boolean
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Er
    
    Set oObject = GetObject("WinNT://" & Form1.txtDomain.Text & "/" & Objectname)
    
    For Each ogroup In oObject.Groups
        isListed = False
        For i = 0 To Form1.lstGroups.ListCount - 1
            If ogroup.Name = Form1.lstGroups.List(i) Then
                isListed = True
                Exit For
            End If
        Next i
        
        If isListed = False Then
            Form1.lstGroups.AddItem ogroup.Name
        End If
    Next
    
    If Form1.lstGroups.ListCount > 0 Then
        Form1.cmdExportGroups.Enabled = True
    Else
        Form1.cmdExportGroups.Enabled = False
    End If
    
    Set oObject = Nothing
    Form1.lblInfo.Caption = "Complete"
    Screen.MousePointer = vbDefault
    Exit Sub
Er:
    ' There was an unexpected error!
    Screen.MousePointer = vbDefault
    Form1.lblInfo.Caption = "Error " & Err.Number & " (" & Err.Description & ")."
    Err.Clear
    Resume Next
End Sub
