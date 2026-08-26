VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Hosts"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtHostname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar pbrStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label lblHostname 
      Caption         =   "Hostname"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoRS As ADODB.Recordset
Dim adoConn As ADODB.Connection
Dim SQL As String
Dim DSN As String
Dim strDB As String

Private Sub Form_Load()
    Dim intRecords As Integer, Hostname As String
    Dim i As Integer
    Dim Ping As DSPINGLib.Ping
    Dim Retval As Integer
    Dim SQLDateTime As String
    Dim NextDate As Date
    
    Me.Show
    Me.Refresh
    
    Set adoRS = CreateObject("ADODB.Recordset")
    Set adoConn = CreateObject("ADODB.Connection")
    Set Ping = CreateObject("DSPing.Ping")
    
Start:
    strDB = App.Path & "\MOJ_MASTER.mdb"
    DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB & ";Persist Security Info=False"
    adoConn.Open DSN
    
    SQL = "SELECT Name from qHosts WHERE (Datetime <= Date()) OR (isNull(DateTime)) ORDER BY Name"
    
    adoRS.Open SQL, adoConn
    ' Get number of records
    i = 0
    Do Until adoRS.EOF
        i = i + 1
        adoRS.MoveNext
        DoEvents
    Loop
    intRecords = i
    
    adoRS.Close
    pbrStatus.Max = intRecords
    
    pbrStatus.Value = 1
    adoRS.Open SQL, adoConn
    
    i = 0
    Do Until adoRS.EOF
        i = i + 1
        Hostname = adoRS("Name")
        txtHostname = Hostname
        lblStatus = "Pinging"
        frmMain.Refresh
        Retval = Ping.DoPing(Hostname, 2, 3000, 32) ' 2 pings, 3000ms Timeout, 32 bytes sent
        If Retval = 0 Then ' Success!
            lblStatus = "Found host"
            frmMain.Refresh
            SQL = "UPDATE tblHosts SET State = 'Up' WHERE Name = '" & Hostname & "'"
            adoConn.Execute SQL
            lblStatus = "Database updated"
            frmMain.Refresh
        Else
            lblStatus = "Host not found"
            frmMain.Refresh
            NextDate = DateAdd("n", 5, Now)
            SQLDateTime = Format(NextDate, "HH:NN AM/PM") & " " & Format(NextDate, "DD MMMM YYYY")
            SQL = "UPDATE tblHosts SET [DateTime] = #" & SQLDateTime & "# WHERE Name = '" & Hostname & "'"
            adoConn.Execute SQL
            lblStatus = "Database updated"
            frmMain.Refresh
        End If
        pbrStatus.Value = i
        adoRS.MoveNext
        DoEvents
    Loop
    
    adoConn.Close
    
    GoTo Start
    Set Ping = Nothing
    Set adoConn = Nothing
    Set adoRS = Nothing
End Sub
