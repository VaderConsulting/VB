VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   12345
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Right"
      Height          =   375
      Left            =   10920
      TabIndex        =   15
      Top             =   9480
      Width           =   855
   End
   Begin VB.OptionButton optMap 
      Caption         =   "Colour"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   9720
      Width           =   1575
   End
   Begin VB.OptionButton optMap 
      Caption         =   "Black and White"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   12
      Top             =   9360
      Value           =   -1  'True
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc adoMap 
      Height          =   330
      Left            =   5160
      Top             =   9360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\MyData\MAP.mdb;Mode=Read;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=H:\MyData\MAP.mdb;Mode=Read;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Mapinfo"
      Caption         =   "adoMap"
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
   Begin VB.CommandButton cmdDown 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "Command1"
      Height          =   2895
      Left            =   11880
      TabIndex        =   2
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "Command1"
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblPos 
      Height          =   255
      Left            =   7200
      TabIndex        =   14
      Top             =   9360
      Width           =   495
   End
   Begin VB.Label lblCoOrd 
      Height          =   255
      Index           =   7
      Left            =   10440
      TabIndex        =   11
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Label lblCoOrd 
      Height          =   255
      Index           =   6
      Left            =   8880
      TabIndex        =   10
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Label lblCoOrd 
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   9
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Label lblCoOrd 
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   8
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Label lblCoOrd 
      Height          =   255
      Index           =   3
      Left            =   10440
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblCoOrd 
      Height          =   255
      Index           =   2
      Left            =   8880
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblCoOrd 
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblCoOrd 
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image imgMap 
      Height          =   8175
      Left            =   600
      Stretch         =   -1  'True
      Top             =   600
      Width           =   11175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gifnw_x(200) As Long
Dim gifnw_y(200) As Long
Dim gifne_x(200) As Long
Dim gifne_y(200) As Long
Dim gifsw_x(200) As Long
Dim gifsw_y(200) As Long
Dim gifse_x(200) As Long
Dim gifse_y(200) As Long
Dim giffilename(200) As String

Dim bmpnw_x(200) As Long
Dim bmpnw_y(200) As Long
Dim bmpne_x(200) As Long
Dim bmpne_y(200) As Long
Dim bmpsw_x(200) As Long
Dim bmpsw_y(200) As Long
Dim bmpse_x(200) As Long
Dim bmpse_y(200) As Long
Dim bmpfilename(200) As String


Private Sub cmdRight_Click()
    x1Now = lblCoOrd(2)
    x2Now = lblCoOrd(6)
    For lp = 1 To 200
        If gifnw_x(lp) = Val(x1Now) Then
            If gifsw_x(lp) = x2Now Then
                lblPos = lp
                MapRefresh
            End If
        End If
    Next lp
End Sub

Private Sub Command1_Click()
    lblPos = Val(lblPos) + 1
    MapRefresh
End Sub

Sub MapRefresh()
    If optMap(0).Value = True Then
        imgMap.Picture = LoadPicture("c:\images\" & giffilename(lblPos))
    Else
        imgMap.Picture = LoadPicture("c:\images\" & bmpfilename(lblPos))
    End If
    imgMap.Refresh
End Sub

Private Sub Form_Load()
    adoMap.Refresh
    adoMap.Caption = adoMap.Recordset.AbsolutePosition & "/" & adoMap.Recordset.RecordCount
    adoMap.RecordSource = "mapinfo"
    adoMap.Refresh
    adoMap.Recordset.MoveFirst
    For lp = Val(lblPos) To adoMap.Recordset.RecordCount - 1
        gifnw_x(lp) = adoMap.Recordset.Fields("nw_x")
        gifnw_y(lp) = adoMap.Recordset.Fields("nw_y")
        gifne_x(lp) = adoMap.Recordset.Fields("ne_x")
        gifne_y(lp) = adoMap.Recordset.Fields("ne_y")
        gifsw_x(lp) = adoMap.Recordset.Fields("sw_x")
        gifsw_y(lp) = adoMap.Recordset.Fields("sw_y")
        gifse_x(lp) = adoMap.Recordset.Fields("se_x")
        gifse_y(lp) = adoMap.Recordset.Fields("se_y")
        giffilename(lp) = adoMap.Recordset.Fields("filename")
        adoMap.Recordset.MoveNext
    Next lp

    adoMap.RecordSource = "gifinfo"
    adoMap.Refresh
    adoMap.Recordset.MoveFirst
    For lp = 1 To adoMap.Recordset.RecordCount
        bmpnw_x(lp) = adoMap.Recordset.Fields("nw_x")
        bmpnw_y(lp) = adoMap.Recordset.Fields("nw_y")
        bmpne_x(lp) = adoMap.Recordset.Fields("ne_x")
        bmpne_y(lp) = adoMap.Recordset.Fields("ne_y")
        bmpsw_x(lp) = adoMap.Recordset.Fields("sw_x")
        bmpsw_y(lp) = adoMap.Recordset.Fields("sw_y")
        bmpse_x(lp) = adoMap.Recordset.Fields("se_x")
        bmpse_y(lp) = adoMap.Recordset.Fields("se_y")
        bmpfilename(lp) = adoMap.Recordset.Fields("filename")
        adoMap.Recordset.MoveNext
    Next lp
    lblPos = 1
    MapRefresh
End Sub

Private Sub optMap_Click(Index As Integer)
    MapRefresh
End Sub
