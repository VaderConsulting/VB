VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Monitor"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5250
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstGroups 
      Height          =   3180
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox lstUsers 
      Height          =   3180
      Left            =   2760
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoConn As ADODB.Connection
Public adoConn2 As ADODB.Connection
Public adoData As ADODB.Recordset
Public DSN As String

' This app uses action type 4 in tblEvents

Private Sub Form_Load()
  Me.Show
  Me.Refresh
  If Dir("c:\temp\All Groups.txt") <> "" Then
    Kill "c:\temp\All Groups.txt"
  End If
  GetGroups
  For lp = 0 To lstGroups.ListCount - 1
    lstGroups.ListIndex = lp
  Next lp
  MsgBox "Output files have been saved in c:\temp"
  Unload Me
  End
End Sub

Sub GetGroups()
  Dim Group As IADsGroup, Member As Object
  lblStatus = "Retrieving groups"
  lblStatus.Refresh
  DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=CBDXAAI"
  ' ADO stuff
  Set adoConn = CreateObject("ADODB.Connection")
  Set adoConn2 = CreateObject("ADODB.Connection")
  Set adoData = CreateObject("ADODB.Recordset")
  
  ' tblMonitor is the list of monitored tasks
  ' Task type 1 = Groupname to monitor
  sql = "SELECT * from tblMonitor WHERE Type = 1"
  adoConn.Open DSN
  adoData.Open sql, adoConn
  
  Do Until adoData.EOF
    Groupname = adoData("Task")
    lstGroups.AddItem adoData("Task")
    adoData.MoveNext
  Loop
  adoData.Close
  Set adoData = Nothing
  Set adoConn2 = Nothing
  Set adoConn = Nothing
  lblStatus = "Idle"
  lblStatus.Refresh
End Sub

Sub GetUsers(Groupname As String)
  Dim Group As IADsGroup, Member As Object
  lblStatus = "Retrieving Users"
  lblStatus.Refresh
  On Error Resume Next
  Open "c:\temp\" & Groupname & ".txt" For Output As #1
  Open "c:\temp\All Groups.txt" For Append As #2
    Print #1, "Members of " & Groupname & ":"
    Print #2, "Members of " & Groupname & ":"
    ' Do NOT take any notice of SMS users.
    Set Group = GetObject("WinNT://POLICE/" + Groupname)
    For Each Member In Group.Members
      If UCase(Left(Member.Name, 3)) <> "SMS" Then
        If Member.Class = "Group" Then
          lstUsers.AddItem Member.Name ' & " (Group)"
          Print #1, Member.Name & " (Group)"
          Print #2, Member.Name & " (Group)"
        Else
          lstUsers.AddItem Member.Name ' & " (" & Member.FullName & ")"
          Print #1, Member.Name & " (" & Member.FullName & ")"
          Print #2, Member.Name & " (" & Member.FullName & ")"
        End If
        lstUsers.Refresh
      End If
    Next
    Print #2, vbCrLf
  Close 2
  Close 1
  lblStatus = "Idle"
  lblStatus.Refresh
End Sub

Private Sub lstGroups_Click()
  lstUsers.Clear
  'lstUsers2.Clear
  Groupname = lstGroups.List(lstGroups.ListIndex)
  GetUsers (Groupname)
  'CheckUsers (Groupname)
End Sub

'Function CheckUsers(Groupname As String) As Boolean
'  Dim FoundUser As Boolean, Username1 As String, Username2 As String
'  If Dir(App.Path & "\" & Groupname & ".txt") <> "" Then
'    lblStatus = "Opening " & Groupname & ".txt for comparison"
'    lblStatus.Refresh
'    Open App.Path & "\" & Groupname & ".txt" For Input As #1
'      Do Until EOF(1)
'        Line Input #1, Group
'        lstUsers2.AddItem Group
'      Loop
'    Close 1
'    ' Check current membership
'    For lp = 0 To lstUsers.ListCount - 1
'      Username1 = lstUsers.List(lp)
'      FoundUser = False
'      For lp2 = 0 To lstUsers2.ListCount - 1
'        Username2 = lstUsers2.List(lp2)
'        If Username1 = Username2 Then FoundUser = True
'      Next lp2
'      If FoundUser = False Then
'        lblStatus = Username1 & " has been added to " & Groupname
'        sql = lblStatus
'        lblStatus.Refresh
'        lstStatus.AddItem lblStatus
'        InserttoDb lblStatus.Caption
'        DoAlerter Groupname, Username1, 1
'      End If
'    Next lp
'    ' Check previous membership
'    For lp2 = 0 To lstUsers2.ListCount - 1
'      Username2 = lstUsers2.List(lp2)
'      FoundUser = False
'      For lp = 0 To lstUsers.ListCount - 1
'        Username1 = lstUsers.List(lp)
'        If Username2 = Username1 Then FoundUser = True
'      Next lp
'      If FoundUser = False Then
'        lblStatus = Username2 & " has been removed from " & Groupname
'        lblStatus.Refresh
'        lstStatus.AddItem lblStatus
'        InserttoDb lblStatus.Caption
'        DoAlerter Groupname, Username2, 2
'      End If
'    Next lp2
'  End If
'
'  lblStatus = "Updating " & Groupname & ".txt"
'  lblStatus.Refresh
'
'  ' Create/replace current user list
'  Open App.Path & "\" & Groupname & ".txt" For Output As #1
'    For lp = 0 To lstUsers.ListCount - 1
'      Group = lstUsers.List(lp)
'      Print #1, Group
'    Next lp
'  Close 1
'
'  lblStatus = "Idle"
'  lblStatus.Refresh
'End Function

'Sub InserttoDb(Data As String)
'  Set adoConn = CreateObject("ADODB.Connection")
'  sql = "INSERT INTO tblEvents (Hostname, Action, Initiated_By, Description, Complete, StartDateTime, ActionType) VALUES "
'  sql = sql & "('CBDXAAA','Domain Group Change','Automatic Monitoring (" & Environ$("computername") & ")','" & Data & "',0,GETDATE(),4)"
'  adoConn.Open DSN
'  adoConn.Execute sql ' Use execute, as we don't want to return a recordset
'End Sub

'Function cSQLDate(Dte) As String
'  If Not IsDate(Dte) Then
'    cSQLDate = ""
'    Exit Function
'  End If
'  s = CStr(Month(Dte))
'  If (Len(s) < 2) Then s = "0" + s
'  If (Day(Dte) < 10) Then s = s + "0"
'  s = s + CStr(Day(Dte))
'  cSQLDate = CStr(Year(Dte)) + s
'End Function
'
'Sub DoAlerter(Groupname As String, Username As String, ChangeType As Integer)
'  On Error Resume Next
'  Set User = GetObject("WinNT://POLICE/" + Username)
'  s = "# Alerter Package" & vbCrLf
'  s = s & "Date=" & Date & vbCrLf
'  s = s & "Time=" & Time & vbCrLf
'  s = s & "Severity=4" & vbCrLf
'  s = s & "System=Domain Group Monitoring" & vbCrLf
'  s = s & "Group=22" & vbCrLf
'  s = s & "Visual=" & "Domain Group Membership Change" & vbCrLf
'  Select Case ChangeType
'    Case 1
'      s = s & "Audio=" & User.FullName & " was added to " & Groupname & vbCrLf
'      s = s & "Mail=" & User.FullName & " was added to " & Groupname & vbCrLf
'      's = s & "Recipients=" & vbCrLf
'    Case 2
'      s = s & "Audio=" & User.FullName & " was removed from " & Groupname & vbCrLf
'      s = s & "Mail=" & User.FullName & " was been removed from " & Groupname & vbCrLf
'      's = s & "Recipients=" & vbCrLf
'  End Select
'    s = s & "Hint=Check Domain group Membership" & vbCrLf
'  Open "\\Perthxsam\packages\" & Username & "-" & Groupname & ".txt" For Output As #1
'    Print #1, s
'  Close 1
'  Open "\\Perthxsam\jobs\" & Username & "-" & Groupname & ".txt" For Output As #1
'    Print #1, s
'  Close 1
'End Sub
'
'
