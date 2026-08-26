VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Data 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Entry"
   ClientHeight    =   7320
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "Data.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeMagazine 
      Caption         =   "Magazine"
      Height          =   735
      Left            =   120
      TabIndex        =   52
      Top             =   4560
      Width           =   5535
      Begin VB.TextBox txtFields 
         DataField       =   "Magazine Issue"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   3480
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Magazine Name"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   840
         TabIndex        =   53
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Issue:"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   56
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Name:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fmeSongs 
      Caption         =   "Songs"
      Height          =   2175
      Left            =   120
      TabIndex        =   48
      Top             =   840
      Width           =   5535
      Begin VB.CommandButton cmdNewSong 
         Caption         =   "Add"
         Height          =   315
         Left            =   4800
         TabIndex        =   51
         ToolTipText     =   "Create a new Album"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   9
         Left            =   720
         TabIndex        =   50
         Top             =   240
         Width           =   3975
      End
      Begin MSDataListLib.DataList dblSongs 
         Bindings        =   "Data.frx":0442
         DataSource      =   "datPrimaryRS"
         Height          =   1425
         Left            =   720
         TabIndex        =   49
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2514
         _Version        =   393216
         ListField       =   "Song Title"
         BoundColumn     =   "Song Title"
      End
   End
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   7800
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   120
      TabIndex        =   45
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Prev"
      Height          =   315
      Left            =   1320
      TabIndex        =   44
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   315
      Left            =   2400
      TabIndex        =   43
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   4680
      TabIndex        =   42
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   3480
      TabIndex        =   41
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame fmeVisual 
      Caption         =   "Visual"
      Height          =   855
      Left            =   4200
      TabIndex        =   38
      Top             =   3120
      Width           =   1455
      Begin VB.CheckBox chkFields 
         DataField       =   "Video"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   15
         Left            =   960
         TabIndex        =   47
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "Video"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   13
         Left            =   960
         TabIndex        =   39
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Picture"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Video:"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fmeWritten 
      Caption         =   "Written"
      Height          =   855
      Left            =   2280
      TabIndex        =   31
      Top             =   3120
      Width           =   1335
      Begin VB.CheckBox chkFields 
         DataField       =   "Lyrics"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   35
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "Sheet Music"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   960
         TabIndex        =   34
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblLabels 
         Caption         =   "Lyrics:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Sheet:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fmeAudio 
      Caption         =   "Audio"
      Height          =   1335
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   1455
      Begin VB.CheckBox chkFields 
         DataField       =   "Vinyl"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   14
         Left            =   960
         TabIndex        =   32
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "Electronic"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   29
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "CD"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkFields 
         DataField       =   "Tape"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   960
         TabIndex        =   25
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Record:"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Electronic:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "CD:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tape:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "..."
      Height          =   285
      Left            =   5280
      TabIndex        =   23
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton cmdAudio 
      Caption         =   "..."
      Height          =   285
      Left            =   5280
      TabIndex        =   22
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "..."
      Height          =   285
      Left            =   5280
      TabIndex        =   21
      Top             =   5880
      Width           =   375
   End
   Begin MSDataListLib.DataCombo dacEras 
      Bindings        =   "Data.frx":045D
      Height          =   315
      Left            =   720
      TabIndex        =   20
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Era Name"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc adoEras 
      Height          =   330
      Left            =   6480
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   15
      CursorType      =   2
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
      Connect         =   $"Data.frx":0473
      OLEDBString     =   $"Data.frx":04FC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from eras"
      Caption         =   "adoEras"
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
   Begin MSDataListLib.DataCombo dacTempos 
      Bindings        =   "Data.frx":0585
      Height          =   315
      Left            =   4080
      TabIndex        =   19
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Tempo Name"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc adoTempos 
      Height          =   330
      Left            =   6480
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   15
      CursorType      =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tempos"
      Caption         =   "adoTempos"
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
   Begin VB.CommandButton cmdNewArtist 
      Caption         =   "New"
      Height          =   315
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "Create a new Artist"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdNewAlbum 
      Caption         =   "New"
      Height          =   315
      Left            =   4920
      TabIndex        =   17
      ToolTipText     =   "Create a new Album"
      Top             =   480
      Width           =   615
   End
   Begin MSDataListLib.DataCombo dacArtists 
      Bindings        =   "Data.frx":059D
      Height          =   315
      Left            =   840
      TabIndex        =   16
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Name"
      BoundColumn     =   "Name"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dacAlbums 
      Height          =   315
      Left            =   840
      TabIndex        =   15
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Name"
      BoundColumn     =   "Name"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc adoAlbums 
      Height          =   330
      Left            =   6480
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   15
      CursorType      =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from albums"
      Caption         =   "adoAlbums"
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
   Begin MSAdodcLib.Adodc adoArtists 
      Height          =   330
      Left            =   6480
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   15
      CursorType      =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\data\access\music2000.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from artists"
      Caption         =   "adoArtists"
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
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   6360
      TabIndex        =   10
      Top             =   7560
      Visible         =   0   'False
      Width           =   6360
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Text"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   1560
      TabIndex        =   4
      Top             =   6600
      Width           =   3615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Tempo"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   5760
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Picture Filename"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   2
      Top             =   5880
      Width           =   3615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Era"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Audio Filename"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   6240
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   6480
      Top             =   120
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Connect         =   $"Data.frx":05B6
      OLEDBString     =   $"Data.frx":063F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Songs"
      Caption         =   " datPrimaryRS"
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
   Begin VB.Label lblAlbum 
      Caption         =   "Album:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblArtist 
      Caption         =   "Artist:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Text Filename:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   9
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tempo:"
      Height          =   255
      Index           =   11
      Left            =   3240
      TabIndex        =   8
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Picture Filename:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Era:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "Audio Filename:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
End
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
    txtFields(9).Text = ""
End Sub

Private Sub cmdNewAlbum_Click()
    frmAlbums.Show vbModal
End Sub

Private Sub cmdNewArtist_Click()
    frmArtists.Show vbModal
End Sub

Private Sub cmdNext_Click()
    If Not (datPrimaryRS.Recordset.EOF) Then
        datPrimaryRS.Recordset.MoveNext
    Else
    End If
End Sub

Private Sub cmdPrev_Click()
    If Not (datPrimaryRS.Recordset.BOF) Then
        datPrimaryRS.Recordset.MovePrevious
    Else
    End If
End Sub

Private Sub dacAlbums_Change()
    txtSQL = "SELECT * FROM Info WHERE Artists_Name like " & Chr(39) & dacArtists.Text & Chr(39) & " and Albums_Name like " & Chr(39) & dacAlbums.Text & Chr(39)
    datPrimaryRS.RecordSource = txtSQL
    datPrimaryRS.Refresh
    txtFields(9).Refresh
End Sub

Private Sub dacAlbums_Click(Area As Integer)
    dacAlbums_Change
End Sub

Private Sub dacArtists_Change()
    If dacArtists.Text <> "" Then
        txtSQL = "SELECT * FROM Info WHERE Artists_Name like " & Chr(39) & dacArtists.Text & Chr(39)
    Else
    End If
    adoAlbums.RecordSource = txtSQL
    adoAlbums.Refresh
    If adoAlbums.Recordset.RecordCount > 0 Then
        dacAlbums.Text = adoAlbums.Recordset.Fields("Albums_Name")
    Else
        dacAlbums.Text = ""
    End If
    dacAlbums.Refresh
End Sub

Private Sub dacEras_Change()
    txtFields(3) = dacEras.SelectedItem
End Sub

Private Sub dacTempos_Change()
    txtFields(11) = dacTempos.SelectedItem
End Sub

Private Sub dalSongs_Click()
    datPrimaryRS.Recordset.Move dalSongs.SelectedItem - 1, 1
End Sub

Private Sub Form_Activate()
    adoAlbums.Refresh
    dacAlbums.Refresh
    adoArtists.Refresh
    dacArtists.Refresh
    
End Sub

Private Sub Form_Load()
    dacArtists.Text = adoArtists.Recordset.Fields("Name").Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
