VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xID As Integer
Private xIP As String
Private isAuthorised As Boolean
Private DataStr As String
Private ConnState As Integer

Private Sub ProcessPacket(ByVal sData As String)
  Dim Cmd As String, d As String
  Dim Parm As String, i As Integer, j As Integer, res As Boolean
  Dim FD As String, FT As String, FZ As String, sGeneral As String
  If (Len(sData) < 5) Then
    If sData = "?" Then
      SendString "****************" + vbCrLf
      SendString "Valid commands:" + vbCrLf
      SendString "A001: Authorisation [password required]" + vbCrLf
      SendString "R002: Request tasklist" + vbCrLf
      SendString "R003: TCP Connection status" + vbCrLf
      SendString "R004: Create Process" + vbCrLf
      SendString "R005: Kill Process [PID Required]" + vbCrLf
      SendString "R006: Kill ALL Client Processes" + vbCrLf
      SendString "R007: Get Time" + vbCrLf
      SendString "R008: Get Dir [Path required]" + vbCrLf
      SendString "R009: Get Free Disk Space (Mb) [Driveletter Required]" + vbCrLf
      SendString "R00A: Get Total Disk Space (Mb) [Driveletter Required]" + vbCrLf
      SendString "R00B: Get % Free Disk Space (Mb) [Driveletter Required]" + vbCrLf
      SendString "R00C: Get % Used Disk Space (Mb) [Driveletter Required]" + vbCrLf
      SendString "R00D: Get Registry Value [Path required]" + vbCrLf
      SendString "R00E: Terminate Connection [ID Required]" + vbCrLf
      SendString "R999: Disconnect" + vbCrLf
      SendString "****************" + vbCrLf
    Else
      SendString "?" + vbCrLf
    End If
    Exit Sub
  End If
  Cmd = Left(sData, 5)
  Parm = Right(sData, Len(sData) - 5)
  If isAuthorised Or Cmd = "R999:" Then
    Select Case Cmd
      Case "R002:" ' Requests a list of all running processes  R002:<TaskID>,<PID>,<Displayname>
        i = Processes.Count
        SendString "Processes" & vbCrLf
        For j = 1 To i
          SendString Processes.Item(j).TaskID & "," & Processes.Item(j).dwProcessID & "," & Processes.Item(j).DisplayName & vbCrLf
        Next j
        SendString "S002:OK" & vbCrLf
      Case "R003:" '  TCP Connections Status
        SendString "TCP Connections" & vbCrLf
        For i = 1 To TCPConns.Count
          SendString i & " " & frmMain.tcpMain.Item(TCPConns.Item(i).ID).RemoteHostIP & ":" & frmMain.tcpMain.Item(TCPConns.Item(i).ID).RemotePort & vbCrLf
        Next i
        SendString "S003:OK" & vbCrLf
      Case "R004:" '  Create Process
        If Parm <> "" Then
          res = ExecuteCmd(Parm, xIP)
          If res Then
            SendString "R004:OK" & vbCrLf
          Else
            SendString "R004:ERROR" & vbCrLf
          End If
        Else
          SendString "? Executable required" + vbCrLf
        End If
      Case "R005:" '  Kill Process
        If Parm <> "" Then
          KillProcess (Parm)
          SendString "S005:OK" & vbCrLf
        Else
          SendString "? PID required" + vbCrLf
        End If
        SendString "S005:OK" & vbCrLf
      Case "R006:" '  Kill All Processes
        i = Processes.Count
        For j = 1 To i
          KillProcess (Processes.Item(j).dwProcessID)
        Next j
        SendString "S006:OK" & vbCrLf
      Case "R007:" '  Get Time
        SendString "S007:" & Format(Time, "HH:MM:SS") & vbCrLf
        SendString "S007:OK" & vbCrLf
      Case "R008:" '  Get Dir
        If Parm <> "" Then
          'On Error Resume Next
          If Right(Parm, 1) <> "\" Then Parm = Parm & "\"
          d = Dir(Parm & "*.*", vbArchive + vbDirectory + vbHidden + vbNormal + vbReadOnly + vbSystem)
          SendString "Directory of " & Parm & vbCrLf
          Do Until Trim(d) = ""
            If d <> "." And d <> ".." And d <> "?" Then
              FT = Format(CDate(FileDateTime(Parm & Trim(d))), "hh:mm:ss")
              FD = Format(FileDateTime(Parm & Trim(d)), "dd-mm-yyyy")
              If (GetAttr(Parm & d) And vbDirectory) = vbDirectory Then
                FZ = "<DIR>       "
              Else
                FZ = Format(FileLen(Parm & Trim(d)), "#,###,###,##0")
              End If
              If Len(FZ) < 13 Then
                For i = Len(FZ) To 13
                  FZ = " " & FZ
                Next i
              End If
              SendString FD & " " & FT & " " & FZ & " " & d & vbCrLf
            End If
            d = Dir
          Loop
        Else
          SendString "? Path required" & vbCrLf
        End If
        SendString "S008:OK" & vbCrLf
        'On Error GoTo 0
      Case "R009:" '  Get Free Disk Space
        If Parm <> "" Then
          SendString "Free Space on " & Parm & " (Mb)" & vbCrLf
          If Right(Parm, 1) <> ":" Then Parm = Parm & ":"
          If Right(Parm, 1) = "\" Then Parm = Left(Parm, Len(Parm) - 1)
          SendString ShowFreeSpace(Parm, 1) & vbCrLf
        Else
          SendString "? Drive letter required" & vbCrLf
        End If
        SendString "R009:OK" & vbCrLf
      Case "R00A:" '  Get Total Disk Space
        If Parm <> "" Then
          SendString "Total Space on " & Parm & " (Mb)" & vbCrLf
          If Right(Parm, 1) <> ":" Then Parm = Parm & ":"
          If Right(Parm, 1) = "\" Then Parm = Left(Parm, Len(Parm) - 1)
          SendString ShowFreeSpace(Parm, 2) & vbCrLf
        Else
          SendString "? Drive letter required" & vbCrLf
        End If
        SendString "R00A:OK" & vbCrLf
      Case "R00B:" '  Get % free Disk Space
        If Parm <> "" Then
          SendString "Percentage Free Space on " & Parm & " (Mb)" & vbCrLf
          If Right(Parm, 1) <> ":" Then Parm = Parm & ":"
          If Right(Parm, 1) = "\" Then Parm = Left(Parm, Len(Parm) - 1)
          SendString ShowFreeSpace(Parm, 3) & vbCrLf
        Else
          SendString "? Drive letter required" & vbCrLf
        End If
        SendString "R00B:OK" & vbCrLf
      Case "R00C:" '  Get % used Disk Space
        If Parm <> "" Then
          SendString "Percentage Used Space on " & Parm & " (Mb)" & vbCrLf
          If Right(Parm, 1) <> ":" Then Parm = Parm & ":"
          If Right(Parm, 1) = "\" Then Parm = Left(Parm, Len(Parm) - 1)
          SendString ShowFreeSpace(Parm, 4) & vbCrLf
        Else
          SendString "? Drive letter required" & vbCrLf
        End If
        SendString "R00C:OK" & vbCrLf
      Case "R00D:" '  Get Registry
        If Parm <> "" Then
          SendString "Registry Value: " & Parm & vbCrLf
          i = Len(Parm)
          j = 0
          Do
            If Mid(Parm, i, 1) = "\" Then j = i
            i = i - 1
          Loop Until j > 0 Or i = 1
          If j > 0 Then
            sGeneral = Trim(GetValue("", Left(Parm, i), Right(Parm, Len(Parm) - (j)), False))
            If sGeneral <> "" Then
              SendString sGeneral & vbCrLf
              SendString "R00D:OK" & vbCrLf
            Else
              SendString "R00D:ERR" & vbCrLf
            End If
          Else
            SendString "R00D:NAK" & vbCrLf
          End If
        Else
          SendString "? Registry Path required" & vbCrLf
        End If
      Case "R00E:" ' Terminate Connection request
        If Parm <> "" Then
          SendString "Disconnecting " & Parm & vbCrLf
          frmMain.tcpMain_Close CInt(Parm)
          SendString "R00E:OK" & vbCrLf
        Else
          SendString "? Connection ID required" & vbCrLf
        End If
      Case "R999:" ' Disconnect request
        SendString "Disconnecting" + vbCrLf
        frmMain.tcpMain_Close xID
      Case Else
        SendString "?" + vbCrLf
    'End If
    End Select
  Else
    Select Case Cmd
      Case "A001:" ' Requesting Authorisation
      If Authorise(Parm) Then
        SendString "S001:OK" + vbCrLf
        Exit Sub
      Else
        SendString "S999:NAK" + vbCrLf
        Exit Sub
      End If
    Case Else
      SendString "Not authorised" & vbCrLf
    End Select
  End If
End Sub

Public Sub DataArrival(ByVal BytesTotal As Integer) ' ************** MODIFIED
  Dim s As String
  Dim i As Integer
  Dim j As Integer
  
  frmMain.tcpMain(xID).GetData s, vbString
  DataStr = DataStr + s
  
  'Check for CRLF Chars
  Do
    i = MinNZ(0, InStr(DataStr, vbCrLf))
    If (i <> 0) Then
      Debug.Print "Received string '" & Left(DataStr, i - 1) & "' from remote host"
      ProcessPacket Left(DataStr, i - 1)
      DataStr = Right(DataStr, Len(DataStr) - i - 1)
      End If
  Loop Until (i = 0)
  DoEvents
End Sub

Private Function Authorise(ByVal Pwd As String)
  Authorise = False
  If isAuthorised Then
    Authorise = True
    Exit Function
  End If
  isAuthorised = (Pwd = "HELLO") 'Change this to a shifting password
  Authorise = isAuthorised
  If Authorise Then
    SendString "Welcome " & xIP & vbCrLf
  End If
End Function

Property Get ID() As Integer
  ID = xID
End Property

Public Sub SendBroadcast(ByVal Data As String)
  If isAuthorised Then SendString Data
End Sub

Private Sub SendString(ByVal Data As String)
  On Error GoTo SendStringError
  frmMain.tcpMain(xID).SendData Data
  Debug.Print "Sending data to client " & xID & ": " & Left(Data, Len(Data) - 2)
  Exit Sub
SendStringError:
  Err.Clear
  Exit Sub
End Sub

Public Sub SetupTCPInfo(ByVal tpcMainID As Integer, RemoteIP As String)
  xID = tpcMainID
  xIP = RemoteIP
End Sub

Private Sub Class_Initialize()
  isAuthorised = False
  xID = 0
  xIP = ""
  DataStr = ""
  ConnState = 0
End Sub
