VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBounce 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Bounce Log"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   Icon            =   "frmBounce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   8730
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAuto 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6240
      TabIndex        =   26
      Text            =   "3"
      Top             =   4080
      Width           =   375
   End
   Begin VB.OptionButton optUnplanned 
      Caption         =   "Unplanned"
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      ToolTipText     =   "NTSS was unaware of outage, OR site was unaware of outage prior"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton optPlanned 
      Caption         =   "Planned"
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      ToolTipText     =   "NTSS knew about outage, and site was informed prior"
      Top             =   1920
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "AutoComplete"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      ToolTipText     =   "Enter a 'Bounce' entry"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "History..."
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      ToolTipText     =   "View history of selected server"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      ToolTipText     =   "Save to database"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Frame fmeHistory 
      Caption         =   "History"
      Height          =   3255
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   8535
      Begin MSDataGridLib.DataGrid dgrHistory 
         Bindings        =   "frmBounce.frx":030A
         Height          =   2895
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3081
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3081
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtReason 
      Height          =   735
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Detailed description of outage"
      Top             =   2760
      Width           =   5895
   End
   Begin VB.TextBox txtSOMS 
      Height          =   285
      Left            =   7320
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame fmeTimes 
      Caption         =   "Times"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   12
      Top             =   840
      Width           =   6015
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Date of outage"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtDowntime 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4560
         TabIndex        =   4
         ToolTipText     =   "Total time server was down"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtUp 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         ToolTipText     =   "Time services returned"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtDown 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         ToolTipText     =   "Time services were lost "
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDowntime 
         Alignment       =   2  'Center
         Caption         =   "Downtime"
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblUp 
         Alignment       =   2  'Center
         Caption         =   "Up"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDown 
         Alignment       =   2  'Center
         Caption         =   "Down"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   8
      ToolTipText     =   "End program"
      Top             =   3600
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc adoBounce 
      Height          =   330
      Left            =   7440
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   3
      CommandTimeout  =   3
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=PERTHXSAC"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=PERTHXSAC"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblevent order by date"
      Caption         =   "Adodc1"
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
   Begin MSDataListLib.DataList dtlServers 
      Bindings        =   "frmBounce.frx":0322
      CausesValidation=   0   'False
      DataField       =   "name"
      DataSource      =   "adoServers"
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "List of servers - from SQL Database"
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5265
      _Version        =   393216
      ListField       =   "name"
      BoundColumn     =   ""
   End
   Begin MSAdodcLib.Adodc adoServers 
      Height          =   330
      Left            =   2640
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   3
      CommandTimeout  =   3
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=PERTHXSAC"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=PERTHXSAC"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select hostname from tblserver where supported like '1' order by hostname"
      Caption         =   "Adodc1"
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
   Begin VB.Label lblMinutes 
      Alignment       =   2  'Center
      Caption         =   "minutes"
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   4380
      Width           =   615
   End
   Begin VB.Label lblServerName 
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
      TabIndex        =   25
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblServer 
      Caption         =   "Server:"
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
      Left            =   2640
      TabIndex        =   24
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblAuto 
      Alignment       =   2  'Center
      Caption         =   $"frmBounce.frx":033B
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label lblReason 
      Caption         =   "Reason:"
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
      Left            =   2760
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblSOMS 
      Caption         =   "SOMS Call number (optional)"
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
      Left            =   4680
      TabIndex        =   16
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label lblServers 
      Alignment       =   2  'Center
      Caption         =   "Select a server from the list below, and fill in the details of the outage."
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
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmBounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' variables required for SQL connection
Dim Servername As String
Dim SQLSource As String
' These are defined in the application's Form_Load procedure
'

Private Sub Form_Load()
    txtDate = Date
    Height = 4935                              ' Set form height so that user can't see history grid
                                               ' set the following line to SQL Server name
    On Error GoTo dbError
    Servername = "PERTHXSAC"
    SQLSource = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=POLICE;Data Source=" & Servername
    Debug.Print SQLSource
    adoServers.ConnectionString = SQLSource
    adoBounce.ConnectionString = SQLSource
    adoServers.Refresh
    Exit Sub
dbError:
    dbErrorD = Err.Description
    dbErrorN = Err.Number
    MsgBox "Error:" & vbCrLf & dbErrorN & vbCrLf & dbErrorD & vbCrLf & vbCrLf & "Check the connection to " & Servername, vbCritical, "Critical error"
    Unload Me
    End
End Sub

Private Sub cmdAuto_Click()
    ' txtAuto is length of time to bounce in the automatic mode
    If Val(txtAuto) > 1 And Val(txtAuto) < 30 Then ' minimum of 1 minute,
                                                   ' maximum of 30 minutes
        txtDown = Format(Time, "hh:mm")
        txtUp = Format(DateAdd("n", Val(txtAuto), txtDown), "hh:mm")
        calcDowntime
        txtReason.SetFocus
    Else
        MsgBox "The autocomplete value must be between 1 and 30 minutes", vbOKOnly + vbCritical
        txtAuto.SetFocus
        txtAuto.SelStart = 0
        txtAuto.SelLength = 5
    End If
End Sub

Private Sub cmdExit_Click(Index As Integer)
    Unload Me
    End                            ' Finish program
End Sub

Private Sub cmdHistory_Click()
    On Error Resume Next           ' Disregard errors when no entries exist for the selected server
    adoBounce.Recordset.MoveFirst
    frmBounce.Height = 8295        ' Change form height to see history of outages
    cmdHistory.Enabled = False     ' Stop user from selecting button when already selected
    dtlServers.Refresh             ' Refresh data
End Sub

Private Sub cmdSave_Click()
    adoBounce.Recordset.AddNew
    dgrHistory.Columns(0).Text = txtDate
    dgrHistory.Columns(1).Text = lblServerName
                                   ' Only two options - planned or unplanned outage
    If optPlanned.Value = True Then
        dgrHistory.Columns(2).Text = "Planned"
    Else
        dgrHistory.Columns(2).Text = "Unplanned"
    End If
                                   ' Update database
    dgrHistory.Columns(3).Text = txtDown
    dgrHistory.Columns(4).Text = txtUp
    dgrHistory.Columns(5).Text = txtDowntime
    dgrHistory.Columns(6).Text = txtSOMS
    dgrHistory.Columns(7).Text = txtReason
    adoBounce.Recordset.Update
    adoBounce.Refresh
                                   ' clear variables used in this 'session'
    txtUp = ""
    txtDown = ""
    txtDowntime = ""
    txtSOMS = ""
    txtReason = ""
    optPlanned.Value = True
    lblServerName = ""
    txtDown.SetFocus
End Sub

Private Sub dtlServers_Click()
    adoBounce.RecordSource = "select * from tblbounce where server like " & Chr(39) & dtlServers.Text & Chr(39) & "order by date"
    adoBounce.Refresh
    dgrHistory.Columns(2).Caption = "Occasion" ' Originally = Planned/Planned, so change to 'Occasion'
                                               ' Change column widths to suit data
    dgrHistory.Columns(0).Width = 1000
    dgrHistory.Columns(1).Width = 1500
    dgrHistory.Columns(2).Width = 1000
    dgrHistory.Columns(3).Width = 800
    dgrHistory.Columns(4).Width = 800
    dgrHistory.Columns(5).Width = 500
    dgrHistory.Columns(6).Width = 900
    dgrHistory.Columns(7).Width = 15000
    lblServerName = dtlServers.Text
    EnableSave                                 ' Enable save button
End Sub

Private Sub txtDown_LostFocus()
    If txtDown <> "" And txtUp <> "" Then
        calcDowntime                           ' Calculate down time
    End If
End Sub

Private Sub txtReason_Change()
    EnableSave                                 ' Enable save button
End Sub

Sub EnableSave()                               ' Enable or disable save button accordingly
    If txtReason <> "" And lblServerName <> "" And txtDown <> "" And txtUp <> "" And txtDowntime <> "" Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
End Sub

Private Sub txtUp_LostFocus()
    If txtDown <> "" And txtUp <> "" Then
        calcDowntime                           ' Calculate down time
    End If
End Sub

Sub calcDowntime()                             ' Calculate down time
        down1 = DateDiff("n", CVDate(txtDown), CVDate(txtUp))
        down2 = Int(down1 / 60)
        down3 = down1 - (down2 * 60)
        downtime = Format(down2, "00") & ":" & Format(down3, "00")
        txtDowntime = downtime
End Sub
