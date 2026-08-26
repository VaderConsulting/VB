Attribute VB_Name = "modTCP"
Option Explicit

Public isAuthorised As Boolean
Private DataStr As String
Public ConnState As Integer
' csNotAuthorised
' csAuthorised
' csConfigured
' csReady

Private EncKey As String


Private Sub ProcessPacket(ByVal sData As String)
  Dim Cmd As String
  Dim Parm As String
  Dim PathOrID As String
  Dim p As Integer
  Dim PID As Long, Speech As String
  Dim i As Integer, j As Integer, SQL As String
  Dim AlertText As String, SeverityText As String, SystemText As String, GroupText As String, HintText As String
  Dim intTemp As Integer
  ' *******************************
  sData = Decrypt(sData, EncKey)
  ' *******************************
  If (Len(sData) < 5) Then Exit Sub
  Cmd = Left(sData, 5)
  Parm = Right(sData, Len(sData) - 5)
  p = InStr(1, Parm, ",")
  If p > 0 Then
    PathOrID = Right(Parm, Len(Parm) - p)
    Parm = Left(Parm, p - 1)
    'Debug.Print PathOrID
  End If
  
  'AddToHistory "Processing :" & Cmd + Parm
  
  ' In any state
  Select Case Cmd
    Case "S10D:" ' Request client time
      SendString "R10D:" & Format(Time, "HH:MM:SS") & vbCrLf
    Case "S10E:" ' Change client time
      If Not IsDate(Parm) Then
        SendString "R10E:ERR" & vbCrLf
        Exit Sub
      End If
      Time = CDate(Parm)
      SendString "R10E:OK" & vbCrLf
    Case "S1FD:"
      If Parm = "OK" Then
        ' Receipt of client broadcast
      Else
        AddToHistory "Master " & MasterHost & "Shutting down"
        WriteToLog 0, "Master " & MasterHost & "Shutting down"
        ConnState = csNotAuthorised
        frmMain.tcpMain.Close
        CloseObjects
        End   ' <--------- Modify  !!!!!!!!!!!
      End If
    Case "S1FE:" ' Repeat last string
      Debug.Print "Resending last String (" & LastSentString & ")"
      SendString LastSentString
      Exit Sub
    Case "S1FF:" ' Disconnect command from Master
      AddToHistory "Received disconnect from " & MasterHost & " : Reason - " & Parm
      WriteToLog 0, "Received disconnect from " & MasterHost & " : " & Parm
      'StopAllProcesses
      ConnState = csNotAuthorised
      frmMain.tcpMain.Close
      CloseObjects
      End
    Case "A001:" ' Authorisation Response
      If (ConnState <> csNotAuthorised) Then Exit Sub
      If (Parm = "OK") Then ConnState = csAuthorised
      Exit Sub
    Case "S101:"                                             ' Configuration Response
      If (ConnState <> csAuthorised) Then Exit Sub
      If Parm = "NAK" Then Exit Sub                          ' Failed Configuration
      If Left(Parm, 2) = "OK" Then                           ' Successful Configuration
        ConnState = csConfigured
        'AgentID = CInt(PathOrID)                             ' Extract Agent ID
        Exit Sub
      End If
      Exit Sub
    Case "S102:"
      If Parm = "OK" Then
        ConnState = csReady
      Else
        ' Do not understand this response (yet)
      End If
      Exit Sub
    Case "S104:"
      If ConnState <> csConfigured Then Exit Sub    ' <---- more graceful ??
      StartDistroTasks
      Exit Sub
    Case "S105:"
      If ConnState <> csReady Then Exit Sub         ' <---- more graceful ??
      StopDistroTasks
      Exit Sub
    Case "S1FC:"
      Dim Alert() As String
      frmMain.lblStatus = "Receiving event(s) from Master"
      frmMain.lblStatus.Refresh
      Debug.Print "Receiving event(s) from Master"
      Alert() = Split(Parm, Chr(7))
      AlertText = Chr(34) & Alert(4) & Chr(34)
      SeverityText = Alert(2)
      SystemText = Chr(34) & Alert(1) & Chr(34)
      ' Default to NTSS group if the Group string isn't numeric
      If IsNumeric(Alert(3)) Then
        GroupText = Alert(3)
      Else
        GroupText = 1
      End If
      If Alert(5) <> "" Then
        HintText = Chr(34) & Alert(5) & Chr(34)
      Else
        HintText = Chr(34) & "No hint available" & Chr(34)
      End If
      frmMain.lstHistory.AddItem Alert(0) & "," & Alert(4) & "," & Alert(1) & "," & Alert(3) & "," & Alert(2) & "," & Alert(5), 0
      If frmMain.lstHistory.ListCount > 50 Then
        frmMain.lstHistory.RemoveItem 50
      End If
      
      SQL = "INSERT INTO tblAlerts ([DateTime],Alert,System,[Group],Severity,Hint) VALUES (" & Chr(34) & Alert(0) & Chr(34) & "," & AlertText & "," & SystemText & "," & GroupText & "," & SeverityText & "," & HintText & ")"
      Set adoConn = New ADODB.Connection
      adoConn.Open DSN
      adoConn.Execute SQL
      adoConn.Close
      Set adoConn = Nothing
      SQL = ""
      If frmMain.Visible = False And boolAgent = True Then
        If frmMain.CharLoaded Then
          frmMain.Character.Show
        End If
      End If
      If boolSounds Then
        Select Case SeverityText
          Case "1"
            Speak "", App.Path & "\" & strInformationSound, CInt(SeverityText), 1
          Case "2"
            Speak "", App.Path & "\" & strWarningSound, CInt(SeverityText), 2
          Case "3"
            Speak "", App.Path & "\" & strCriticalSound, CInt(SeverityText), 3
        End Select
      End If
      If boolOutputSummaryOnly Then
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
        adoConn.Open DSN
        SQL = "SELECT Count (*) as Total FROM tblAlerts WHERE Complete=false AND Severity=3"
        adoRS.Open SQL, DSN
        If adoRS("Total") > 0 Then
          Speech = adoRS("Total") & " Pending Critical Alerts"
          Speak Speech, "", 15, 3
        End If
        adoRS.Close
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
        adoConn.Open DSN
        SQL = "SELECT Count (*) as total FROM tblAlerts WHERE Complete=false AND Severity=2"
        adoRS.Open SQL, DSN
        If adoRS("Total") > 0 Then
          Speech = adoRS("Total") & " Pending Warning Alerts"
          Speak Speech, "", 15, 3
        End If
        adoRS.Close
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
        adoConn.Open DSN
        SQL = "SELECT Count (*) as total FROM tblAlerts WHERE Complete=false AND Severity=1"
        adoRS.Open SQL, DSN
        If adoRS("Total") > 0 Then
          Speech = adoRS("Total") & " Pending Information Alerts"
          Speak Speech, "", 15, 3
        End If
        adoRS.Close
      Else
        Select Case Alert(2)
          Case 1
            Speech = "Information event "
          Case 2
            Speech = "Warning event "
          Case 3
            Speech = "Critical alert "
        End Select
        If boolVerbose Then Speech = Speech & "for "
        If Not IsNumeric(Alert(3)) Then
          Select Case Alert(3)
            Case "NTSS"
              Speech = Speech & "NTSS"
            Case "WAN"
              Speech = Speech & "WAN"
            Case "EUC"
              Speech = Speech & "EUC"
            Case "Helpdesk"
              Speech = Speech & "Help desk"
          End Select
        Else
          If (Alert(3) And 1) <> 0 Then Speech = Speech & "NTSS "
          If (Alert(3) And 2) <> 0 Then Speech = Speech & "Help desk "
          If (Alert(3) And 4) <> 0 Then Speech = Speech & "EUC "
          If (Alert(3) And 8) <> 0 Then Speech = Speech & "WAN "
        End If
        Speak Speech, "", CInt(Alert(3)), CInt(SeverityText)
        If boolVerbose Then
          Speech = "Description follows: "
          Speak Speech, "", CInt(Alert(3)), CInt(SeverityText)
        End If
        Speech = ""
        Speech = Speech & Alert(4)
        Speak Speech, "", CInt(Alert(3)), CInt(SeverityText)
        Speech = ""
        If boolVerbose Then
          Speech = "Suggested action to correct this follows: "
          Speak Speech, "", CInt(Alert(3)), CInt(SeverityText)
        End If
        Speech = ""
        Speech = Speech & Alert(5)
        Speak Speech, "", CInt(Alert(3)), CInt(SeverityText)
        frmMain.lblStatus = "Idle"
        If frmMain.lstDetails.ListCount = 0 Then
          frmMain.lstDetails.AddItem "[Select event type]"
        End If
      End If
      If frmMain.Visible = False And boolAgent = True Then
        If frmMain.CharLoaded Then
          frmMain.Character.Hide
        End If
      End If
  End Select
End Sub

Public Sub TCPDataArrival(ByVal BytesTotal As Integer)
  Dim s As String
  Dim i As Integer
  Dim j As Integer
  Dim t As String
  
  frmMain.tcpMain.GetData s, vbString
  DataStr = DataStr + s
  
  'Check for CRLF Chars
  Do
    i = MinNZ(0, InStr(DataStr, Chr(13)))
    i = MinNZ(i, InStr(DataStr, Chr(10)))
    'i = InStr(5, DataStr, vbCrLf)
    If (i <> 0) Then
      t = Left(DataStr, i - 1)
      If DataStr <> "" Then
        DataStr = Right(DataStr, Len(DataStr) - i)
      End If
      ProcessPacket t
    End If
  Loop Until (i = 0)
  DoEvents
End Sub

Public Sub SendString(ByVal Data As String)
  On Error GoTo SendStringError
  LastSentString = Data
  
  If Right(Data, 2) = vbCrLf Then Data = Left(Data, Len(Data) - 2)
  'AddToHistory "Sending " & Data & " to " & MasterHost
  
  Data = Encrypt(Data, EncKey) + vbCrLf
  
   frmMain.tcpMain.SendData Data
  DoEvents
  Exit Sub
SendStringError:
  Err.Clear
  Exit Sub
End Sub

