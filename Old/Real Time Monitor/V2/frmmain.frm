VERSION 5.00
Object = "{B186D399-9B91-41CB-9242-E12B96CA99E1}#1.0#0"; "DRPing.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tusk Real Time Monitoring"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTOMSFailure 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4440
      Top             =   1920
   End
   Begin MSWinsockLib.Winsock tcpMain 
      Left            =   3960
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrDb 
      Interval        =   10000
      Left            =   3960
      Top             =   1440
   End
   Begin VB.ListBox lstServers 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Timer tmrPing 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3480
      Top             =   1440
   End
   Begin Ping.drPing drPing 
      Left            =   3480
      Top             =   1920
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Label lblQueryingTOMS 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblTOMSStatus 
      Alignment       =   2  'Center
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblTOMS 
      Alignment       =   2  'Center
      Caption         =   "TOMS Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DSN As String
Dim adoConn As ADODB.Connection, adoRS As ADODB.Recordset, adoRS2 As ADODB.Recordset
Dim SQLServerName As String, SQLUserName As String, DefaultTimeout As Integer

Private Sub Form_Load()
    Dim TraceResult As String, SQL As String
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    Set adoRS2 = CreateObject("ADODB.Recordset")
    
    ' Read application defaults
    SQLServerName = ReadIniFile(App.Path & "\" & "Monitor.ini", "Setup", "SQLServerName")
    SQLUserName = ReadIniFile(App.Path & "\" & "Monitor.ini", "Setup", "SQLUserName")
    DefaultTimeout = CInt(ReadIniFile(App.Path & "\" & "Monitor.ini", "Setup", "DefaultTimeout"))
    
    ' Set application minimums (overriding defaults retrieved from .ini if necessary)
    If SQLServerName = "" Then SQLServerName = "(Local)"
    If SQLUserName = "" Then SQLUserName = "sa"
    If DefaultTimeout = 0 Then DefaultTimeout = 3000
    
    Me.Show
    Me.Refresh
    
    ' For SQL/MSDE:
    DSN = "Provider=SQLOLEDB.1;Initial Catalog=TOMSMonitor;Data Source=" & SQLServerName & ";User Id=" & SQLUserName
        
    SQL = "SELECT TOP 10 * FROM tblHosts ORDER BY Name"
    
    adoConn.Open DSN
    adoRS.Open SQL, adoConn
    
    lblStatus = "Retrieving hostnames and current state"
    lblStatus.Refresh
    
    Do Until adoRS.EOF
        If adoRS("isUp") Then
            lstServers.AddItem adoRS("Name") & ": UP"
        Else
            lstServers.AddItem adoRS("Name") & ": DOWN"
        End If
        adoRS.MoveNext
    Loop
    
    lblStatus = "Idle"
    lblStatus.Refresh
    
    adoRS.Close
    
    ' Startup timer to initiate monitoring
    tmrPing.Enabled = True
End Sub

Sub UpdateRoute(strIPAddress As String, HostID As Long)
    Dim TraceResult As String, SQL As String
    TraceResult = TraceRT(strIPAddress)
    If Left(TraceResult, 1) = "," Then TraceResult = Right(TraceResult, Len(TraceResult) - 1)
            
    ' Insert route into db
    If Trim(TraceResult) <> "" Then
        SQL = "INSERT INTO tblRoute (HostID,Path) VALUES (" & HostID & ",'" & TraceResult & "')"
        adoConn.Execute SQL
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set adoRS2 = Nothing
    Set adoRS = Nothing
    Set adoConn = Nothing
End Sub

Private Sub lblTOMSStatus_Change()
    Dim EmailName As String
    If lblTOMSStatus = "OK" Then
        lblTOMSStatus.ForeColor = &HFF00&
    ElseIf lblTOMSStatus = "DOWN" Then
        lblTOMSStatus.ForeColor = &HFF&
    Else
        lblTOMSStatus.ForeColor = &H80000012
    End If
    If adoConn.State = 0 Then
        adoConn.Open DSN
    End If
    ' Get mail recipient name
    EmailName = ReadIniFile(App.Path & "\" & "Monitor.ini", "Setup", "TOMSEmail")
    
    adoConn.Execute "INSERT INTO tblTOMS ([Time],Event) VALUES ('" & Format(Now, "dd/mmm/yyyy hh:nn:ss") & "','" & LCase(lblTOMSStatus) & "')"
    If EmailName <> "" Then
        SendMail EmailName, "TOMS is " & LCase(lblTOMSStatus) & ".", "TOMS is " & LCase(lblTOMSStatus) & "."
    End If
    adoConn.Close
End Sub

Private Sub tmrDb_Timer()
    Dim d As Date
    lblQueryingTOMS = "Querying..."
    d = Now
    DoOracleCommand "(CONNECT_DATA=(COMMAND=ping))"
    tmrTOMSFailure.Enabled = True
End Sub

Private Sub tmrPing_Timer()
    Dim pingResult As String, isUp As Boolean, WasUp As Boolean
    Dim TraceResult As String, EventID As Long
    Dim HostID As Long, Hostname As String, PingTimeout As Long, HostIP As String, Pingtime As Date
    Dim ENumber As Long, EDescription As String
    Dim Path As String, Hosts() As String, lp As Integer
    Dim RecipientID As Integer, RecipientList As String, Recipients() As String
    Dim SQL As String, SQL2 As String
        
    On Error GoTo er
    
    tmrPing.Enabled = False
    lstServers.Clear
    
    SQL = "SELECT TOP 10 * FROM tblHosts ORDER BY Name"
    
    If adoConn.State = 0 Then adoConn.Open DSN
    ' The above line ensures that the Connection object is open after a run-time error.
    
    adoRS.Open SQL, adoConn
    Debug.Print "Opened adoRS"
    ' For each specified host, check its state via a ping
    Do Until adoRS.EOF
        Hostname = adoRS("Name")
        HostIP = adoRS("IP")
        PingTimeout = adoRS("Timeout")
        HostID = adoRS("ID")
        RecipientID = adoRS("EmailID")
        
        Debug.Print "Inside adoRS.eof Loop for " & Hostname
        ' Confirm Route exists
        SQL2 = "SELECT path from tblRoute WHERE HostID = " & HostID
        adoRS2.Open SQL2, adoConn
        If adoRS2.EOF Then ' No route!
            UpdateRoute HostIP, HostID
        End If
        adoRS2.Close
        
        lblStatus = "Pinging " & Hostname
        lblStatus.Refresh
        pingResult = drPing.Ping(HostIP, PingTimeout)
        
        ' Get time+date so we know when host was <un>available
        Pingtime = Now
        
        ' On the off chance the string returned is not in the form xxxxxx where x is any digit, make it a numeric.
        If Not IsNumeric(pingResult) Then pingResult = PingTimeout + 1
        
        ' Is the host up?  If the time returned was greater than the timeout specified, then we consider it as down
        If CLng(pingResult) > PingTimeout Then
            isUp = False ' Nope
            lstServers.AddItem Hostname & ": DOWN"
            
            ' Update database
            adoConn.Execute "UPDATE tblHosts SET isUp = 0, [Time] = '" & Format(Pingtime, "dd/mmm/yyyy hh:nn:ss") & "' WHERE Name = '" & Hostname & "'"
            
            Debug.Print "Wish to initiate tracert for " & Hostname
            ' Update user
            lblStatus = "Initiating pseudo tracert for " & Hostname
            lblStatus.Refresh
            
            ' Retrieve path to destination from db
            SQL2 = "SELECT path FROM tblRoute WHERE HostID = " & HostID
            Debug.Print "SQL Statement: " & SQL2
            adoRS2.Open SQL2, adoConn
            Path = adoRS2("Path")
            adoRS2.Close
            
            ' Split out the known route into individual IP destinations
            Hosts() = Split(Path, ",")
            
            For lp = LBound(Hosts()) To UBound(Hosts())
                ' Ping the IP destination and see if it is up - default timeout is 'DefaultTimeout' mS
                Debug.Print "Pinging route: " & Hosts(lp)
                lblStatus = "Pinging route to " & Hostname & ": " & Hosts(lp)
                frmMain.Refresh
                pingResult = drPing.Ping(Hosts(lp), DefaultTimeout)
                If pingResult > DefaultTimeout Then ' IP Destination not found
                    TraceResult = TraceResult & "," & Hosts(lp) & ":0"
                    lblStatus = lblStatus & " DOWN"
                Else
                    TraceResult = TraceResult & "," & Hosts(lp) & ":1"
                    lblStatus = lblStatus & " UP"
                End If
                If Left(TraceResult, 1) = "," Then TraceResult = Right(TraceResult, Len(TraceResult) - 1)
                Debug.Print TraceResult
            Next lp
        Else
            isUp = True  ' Certainly is
            lstServers.AddItem Hostname & ": UP"
            
            ' Update database
            adoConn.Execute "UPDATE tblHosts SET isUp = 1, [Time] = '" & Format(Pingtime, "dd/mmm/yyyy hh:nn:ss") & "' WHERE Name = '" & Hostname & "'"
        End If
        
        WasUp = adoRS("isUp") ' wassup?  :)
        
        ' Now check if the current state is different to what it was
        If WasUp <> isUp Then
            lblStatus = "Change of state: " & Hostname
            lblStatus.Refresh
            
            ' Get recipient(s) of state change event
            SQL2 = "SELECT Address FROM tblEmail WHERE ID = " & RecipientID
            adoRS2.Open SQL2, adoConn
            If Not adoRS2.EOF Then
                RecipientList = adoRS2("Address")
                Recipients() = Split(RecipientList, ";")
            End If
            adoRS2.Close
            
            If isUp Then
                adoConn.Execute "INSERT INTO tblEvents ([Time],HostID,Event) VALUES ('" & Format(Pingtime, "dd/mmm/yyyy hh:nn:ss") & "'," & HostID & ",'up')"
                For lp = LBound(Recipients()) To UBound(Recipients())
                    SendMail Recipients(lp), Hostname & " is now up." & vbCrLf & "This mail was automatically generated by TOMSMonitor.", Hostname & " up."
                Next lp
            Else
                adoConn.Execute "INSERT INTO tblEvents ([Time],HostID,Event) VALUES ('" & Format(Pingtime, "dd/mmm/yyyy hh:nn:ss") & "'," & HostID & ",'down')"
                
                For lp = LBound(Recipients()) To UBound(Recipients())
                    SendMail Recipients(lp), Hostname & " is down." & vbCrLf & "This mail was automatically generated by TOMSMonitor.", Hostname & " down."
                Next lp
                
                ' Get ID of most recent event
                SQL2 = "SELECT TOP 1 id FROM tblEvents ORDER BY id DESC"
                adoRS2.Open SQL2, adoConn
                EventID = adoRS2("ID")
                adoRS2.Close
                
                ' Insert tracert results into db
                adoConn.Execute "INSERT INTO tblTrace (EventID,Result) VALUES ('" & EventID & "','" & TraceResult & "')"
                Debug.Print "Trace results:" & TraceResult
                
            End If
        End If
        adoRS.MoveNext
        
        'Reset variables
        isUp = False
        pingResult = ""
        Hostname = ""
        HostIP = ""
        PingTimeout = 0
        EventID = 0
        HostID = 0
        Pingtime = CDate("01 Jan 1970 12:00")
        TraceResult = ""
        RecipientID = 0
        RecipientList = ""
        Erase Recipients
        DoEvents
    Loop
    
    lblStatus = "Monitoring cycle complete."
    lblStatus.Refresh
    
    adoRS.Close
    Debug.Print "Closed adoRS"
    'adoConn.Close
    
    tmrPing.Enabled = True
    Exit Sub
er:
    'On Error Resume Next
    ENumber = Err.Number
    EDescription = Err.Description
    If ENumber = 424 Then ' 'Object required' - happens if you try and use an object when it is unloaded
        Resume Next
    Else
        ' Insert error string into db
        Debug.Print "Error!: " & EDescription
        If ENumber <> 0 Then
            EDescription = Replace(EDescription, "'", "''")
            SQL = "INSERT INTO tblErrors (Source,[Number],Description) VALUES ('tmrPing.Timer'," & ENumber & ",'" & EDescription & "')"
            Set adoConn = CreateObject("ADODB.Connection")
            Set adoRS = CreateObject("ADODB.Recordset")
            If adoConn.State <> 0 Then adoConn.Close
            adoConn.Open DSN
            adoConn.Execute SQL
        
            ' Shut down gracefully
            adoConn.Close
        End If
    End If
    tmrPing.Enabled = True
End Sub

Private Sub tcpMain_Connect()
    Dim bytData() As Byte
    Dim intIdx As Integer
    
    ReDim bytData(intPacketLength - 1)
    For intIdx = 0 To intPacketLength - 1
        bytData(intIdx) = bytPacket(intIdx)
    Next
    
    tcpMain.SendData bytData
End Sub

Private Sub DisplayResponse(ByRef bytData() As Byte)
    Dim intIdx As Integer
    Dim strMisc As String
    Dim blnStartDisplay As Boolean
    Dim intIndent As Integer
    
    intIndent = 0
    strMisc = ""
    For intIdx = LBound(bytData) To UBound(bytData)
        blnStartDisplay = (blnStartDisplay Or (bytData(intIdx) = Asc("(")))
        If blnStartDisplay Then
            strMisc = strMisc + Chr(bytData(intIdx))
            If (bytData(intIdx) = Asc("(")) Then
                intIndent = intIndent + 1
                strMisc = strMisc + vbCrLf + String(intIndent * 4, " ")
            End If
            If (bytData(intIdx) = Asc(")")) Then
                intIndent = intIndent - 1
                strMisc = strMisc + vbCrLf + String(intIndent * 4, " ")
            End If
        End If
    Next
    If InStr(1, strMisc, "ERR=0") > 0 Then
        lblTOMSStatus = "OK"
    Else
        lblTOMSStatus = "ERROR"
    End If
End Sub

Private Sub tcpMain_DataArrival(ByVal bytesTotal As Long)
  Dim bytInbound() As Byte
  
  ReDim bytInbound(bytesTotal)
  
  ' Disable reporting of TOMS failure
  tmrTOMSFailure.Enabled = False
  lblQueryingTOMS = ""
  
  tcpMain.GetData bytInbound, vbByte
  
  DisplayResponse bytInbound
  
End Sub

Private Sub tmrTOMSFailure_Timer()
    tmrTOMSFailure.Enabled = False
    lblTOMSStatus = "DOWN"
    lblQueryingTOMS = ""
End Sub

Sub SendMail(strRecipient As String, strMessage As String, strSubject As String)
    Dim oMail As CDONTS.NewMail
    Set oMail = CreateObject("CDONTS.NewMail")
    oMail.To = strRecipient
    oMail.From = "TOMS_Monitor@tusk.com.au"
    oMail.Subject = strSubject
    
    strMessage = strMessage & vbCrLf & "This mail was automatically generated by TOMSMonitor hosted on " & Environ$("computername") & "."
    
    oMail.Body = strMessage
    oMail.Send
    Debug.Print "Mail sent to " & strRecipient
    Set oMail = Nothing
End Sub
