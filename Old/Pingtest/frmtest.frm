VERSION 5.00
Object = "{FFE08E23-32D7-11D4-9665-00508B63E5C7}#2.0#0"; "ctlPing.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ping Test"
   ClientHeight    =   3270
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5040
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoop 
      Caption         =   "Loop"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   735
   End
   Begin VB.ListBox lstServers 
      Height          =   2595
      Left            =   2640
      TabIndex        =   20
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      TabIndex        =   17
      Top             =   2880
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc adoServers 
      Height          =   330
      Left            =   120
      Top             =   3480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=NTSS;Data Source=PERTHXSAC"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=NTSS;Data Source=PERTHXSAC"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT IP, Name FROM Servers WHERE (supported LIKE 'Y' AND IP > '' ) AND IP <> '%NOT%'"
      Caption         =   "adoServers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtPing 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "16"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "42"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "187"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "165"
      Top             =   4080
      Width           =   615
   End
   Begin ctlPing.uPing uPing1 
      Left            =   3000
      Top             =   4080
      _ExtentX        =   2196
      _ExtentY        =   820
      IPAddress       =   ""
      Octet1          =   "10"
      Octet2          =   "1"
      Octet3          =   "1"
      Octet4          =   "21"
      Octet1          =   "10"
      Octet2          =   "1"
      Octet3          =   "1"
      Octet4          =   "21"
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   3
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   4
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   5
      Left            =   1320
      TabIndex        =   9
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   6
      Left            =   1560
      TabIndex        =   10
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   7
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   8
      Left            =   2040
      TabIndex        =   12
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pbrPing 
      Height          =   1815
      Index           =   9
      Left            =   2280
      TabIndex        =   13
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   300
      Orientation     =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   5160
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5160
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblTimeoutValue 
      Caption         =   "3000 mS"
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblTimeout 
      Caption         =   "Timeout:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuTimeout 
         Caption         =   "Timeout"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Results"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ping As Boolean

Private Sub cmdLoop_Click()
    lstServers.Enabled = False
    Command2.Enabled = False
    cmdLoop.Enabled = False
    Ping = True
    Open "c:\temp\ping " & Format(Date, "dd-mm-yy") & ".csv" For Append As #1
        Print #1, "Name,Response"
    Close 1
    For lp = 0 To lstServers.ListCount - 1
        lstServers.ListIndex = lp
        DoEvents
        uPing1.IPAddress = txtIP
        uPing1.Ping
        res = uPing1.RoundTripTime
        txtPing.Text = pbrPing(lp).Value
        txtPing.Refresh
        pbrPing(lp).Refresh
        If mnuSave.Checked = True Then
            Open "c:\temp\ping " & Format(Date, "dd-mm-yy") & ".csv" For Append As #1
                Print #1, lstServers.Text & "," & uPing1.RoundTripTime
            Close 1
        End If
        If Ping = False Then Exit For
    Next lp
    lstServers.Enabled = True
    Command2.Enabled = True
    cmdLoop.Enabled = True
    Exit Sub
End Sub

Private Sub Command2_Click()
    frmTest.Caption = "Ping Test - " & lstServers.Text
    frmTest.Refresh
    Ping = True
    For lp = 0 To 9
        pbrPing(lp).Max = uPing1.PING_TIMEOUT
        If Ping = False Then Exit Sub
    Next lp
    GetPingResults
End Sub
Sub GetPingResults()
    Do Until Ping = False
        For lp = 0 To 9
            uPing1.IPAddress = txtIP
            uPing1.Ping
            res = uPing1.RoundTripTime
            If res <> "Timeout" And res <> "Error" And res < pbrPing(lp).Max Then
                pbrPing(lp).Value = res
            Else
                pbrPing(lp).Value = pbrPing(lp).Max
            End If
            txtPing.Text = pbrPing(lp).Value
            txtPing.Refresh
            pbrPing(lp).Refresh
            DoEvents
            If Ping = False Then Exit Sub
        Next lp
    Loop
End Sub

Private Sub Command3_Click()
    Ping = False
    frmTest.Caption = "Ping Test"
    frmTest.Refresh
End Sub

Private Sub dtlServers_Click()
    uPing1.IPAddress = dtlServers.BoundText
    txtIP = dtlServers.BoundText
    'uPing1.Ping
End Sub

Private Sub Form_Load()
    Ping = False
    adoServers.Refresh
    Do Until adoServers.Recordset.EOF
        lstServers.AddItem adoServers.Recordset.Fields("Name")
        adoServers.Recordset.MoveNext
    Loop
    If Dir("c:\temp\ping " & Format(Date, "dd-mm-yy") & ".csv") <> "" Then
        retval = MsgBox("Ping results for today already exist.  Do you wish to delete these?", vbDefaultButton1 + vbYesNo, "Confirmation")
        If retval = vbYes Then
            Kill "c:\temp\ping " & Format(Date, "dd-mm-yy") & ".csv"
        End If
    End If
End Sub

Private Sub lstServers_Click()
    If lstServers.Text = "" Then Exit Sub
    adoServers.Recordset.MoveFirst
    
    Do Until adoServers.Recordset.Fields("Name") = lstServers.Text
        adoServers.Recordset.MoveNext
    Loop
    txtIP = adoServers.Recordset.Fields("IP")
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuSave_Click()
    If mnuSave.Checked = True Then
        mnuSave.Checked = False
    Else
        mnuSave.Checked = True
    End If
End Sub

Private Sub mnuTimeout_Click()
    timeout = InputBox("Enter new value for Ping Timeout (in mS)", "New Value", 3000)
    If Val(timeout) <> 0 Then
        uPing1.PING_TIMEOUT = Val(timeout)
        lblTimeoutValue = uPing1.PING_TIMEOUT & " mS"
    End If
End Sub
