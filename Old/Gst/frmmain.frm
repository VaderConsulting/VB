VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GST Tool"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar pbrData 
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSDataGridLib.DataGrid dgrData 
      Bindings        =   "frmMain.frx":0000
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin MSAdodcLib.Adodc adoData1 
      Height          =   330
      Left            =   11160
      Top             =   5040
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
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
      Connect         =   $"frmMain.frx":0017
      OLEDBString     =   $"frmMain.frx":00F1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from items"
      Caption         =   "adoData1"
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   13920
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblSearch 
      Caption         =   "Search criteria:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblCount 
      Caption         =   "Records found"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblRecords 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   5040
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuEntries 
         Caption         =   "Entries"
         Begin VB.Menu mnuRandomize 
            Caption         =   "Randomize"
         End
         Begin VB.Menu mnuCalculate 
            Caption         =   "Calculate"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    txtSearch.SetFocus
End Sub

Private Sub mnuCalculate_Click()
    calculate
End Sub

Private Sub mnuExit_Click()
    adoData1.RecordSource = ""
    Unload Me
End Sub

Private Sub mnuRandomize_Click()
    random
End Sub

Private Sub Form_Load()
    lblRecords.Caption = adoData1.Recordset.RecordCount
    'txtSearch.SetFocus
End Sub

Sub calculate()
    pbrData.Visible = True
    Screen.MousePointer = vbHourglass
    adoData1.Recordset.MoveFirst
    pbrData.Max = adoData1.Recordset.RecordCount
    pbrData.Min = 1
    Do Until adoData1.Recordset.EOF
        pbrData.Value = adoData1.Recordset.AbsolutePosition
        Description = adoData1.Recordset.Fields("Description")
        buy = adoData1.Recordset.Fields("Current Tax rate")
        markuprate = adoData1.Recordset.Fields("Mark up rate")
        gst = adoData1.Recordset.Fields("GST Tax rate")
        buyprice = adoData1.Recordset.Fields("Buy Price")
        inctax = buyprice + (buyprice * (buy / 100))
        incgst = buyprice + (buyprice * (gst / 100))
        currentprice = inctax + (inctax * (markuprate / 100))
        gstprice = incgst + (incgst * (markuprate / 100))
        adoData1.Recordset.Fields("Current Price") = currentprice
        adoData1.Recordset.Fields("GST Price") = gstprice
        adoData1.Recordset.Update
        adoData1.Recordset.MoveNext
    Loop
    pbrData.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Sub random()
    pbrData.Visible = True
    Screen.MousePointer = vbHourglass
    Randomize Timer
    adoData1.Recordset.MoveLast
    pbrData.Max = 10000
    pbrData.Min = 1
    For a = 1 To 10000
        pbrData.Value = a
        adoData1.Recordset.AddNew
        adoData1.Recordset.Fields("Description") = "Random " & adoData1.Recordset.AbsolutePosition
        buyprice = Int(Rnd(1) * 1000) + 1
        tax = Int(Rnd(1) * 8) + 1
        If tax = 1 Then buy = 22
        If tax = 2 Then buy = 10
        If tax = 3 Then buy = 19.5
        If tax = 4 Then buy = 30
        If tax = 5 Then buy = 0.5
        If tax > 5 Then buy = 0
        markup = Int(Rnd(1) * 3) + 1
        If markup = 1 Then markuprate = 10
        If markup = 2 Then markuprate = 30
        If markup = 3 Then markuprate = 100
        gst = 10
        If Int(Rnd(1) * 50) + 1 = 1 Then gst = 0
        
        inctax = buyprice + (buyprice * (buy / 100))
        incgst = buyprice + (buyprice * (gst / 100))
        currentprice = inctax + (inctax * (markuprate / 100))
        gstprice = incgst + (incgst * (markuprate / 100))
        barcode = ""
        For b = 1 To 12
            barcode = barcode & Chr(Int(Rnd(1) * 10) + 48) ' random numerals
        Next b
        
        adoData1.Recordset.Fields("Current Tax Rate") = buy
        adoData1.Recordset.Fields("Mark up Rate") = markuprate
        adoData1.Recordset.Fields("GST Tax Rate") = gst
        adoData1.Recordset.Fields("Buy Price") = buyprice
        adoData1.Recordset.Fields("Current Price") = currentprice
        adoData1.Recordset.Fields("GST Price") = gstprice
        adoData1.Recordset.Fields("Barcode") = barcode
    Next a
    adoData1.Recordset.Update
    adoData1.Refresh
    lblRecords = adoData1.Recordset.RecordCount
    pbrData.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtSearch_Change()
    txtSQL = "select * from items where [barcode] like " & Chr(39) & txtSearch & "%" & Chr(39)
    adoData1.RecordSource = txtSQL
    adoData1.Refresh
    lblRecords.Caption = adoData1.Recordset.RecordCount
End Sub
