VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Search"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "Search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdScale 
      Caption         =   "Scale..."
      Height          =   495
      Left            =   11400
      TabIndex        =   24
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox chkLegend 
      Caption         =   "Show Legend"
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   14280
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "Graph"
      Height          =   495
      Left            =   13320
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12360
      TabIndex        =   22
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "List"
      Height          =   495
      Left            =   11400
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame fmeGraph 
      Caption         =   "Graph Type"
      Height          =   675
      Left            =   5760
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   5130
      Begin VB.OptionButton optChart 
         Caption         =   "3D Bar"
         Height          =   272
         Index           =   0
         Left            =   128
         TabIndex        =   21
         Top             =   240
         Width           =   912
      End
      Begin VB.OptionButton optChart 
         Caption         =   "2D Bar"
         Height          =   272
         Index           =   1
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   912
      End
      Begin VB.OptionButton optChart 
         Caption         =   "3D Area"
         Height          =   272
         Index           =   4
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Width           =   1040
      End
      Begin VB.OptionButton optChart 
         Caption         =   "3D Step"
         Height          =   272
         Index           =   6
         Left            =   4080
         TabIndex        =   18
         Top             =   240
         Width           =   915
      End
      Begin VB.OptionButton optChart 
         Caption         =   "2D Pie"
         Height          =   272
         Index           =   14
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   912
      End
   End
   Begin ComctlLib.ProgressBar pbrResults 
      Height          =   270
      Left            =   1800
      TabIndex        =   15
      Top             =   10800
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   476
      _Version        =   327682
      Appearance      =   1
      Max             =   10
   End
   Begin MSChartLib.MSChart chtResults 
      Height          =   9735
      Left            =   1920
      OleObjectBlob   =   "Search.frx":0442
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   8865
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Height          =   10215
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Text            =   "*.*"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame fmeResults 
      Caption         =   "Results"
      Height          =   10065
      Left            =   1800
      TabIndex        =   4
      Top             =   975
      Width           =   9105
      Begin VB.ListBox List1 
         Height          =   9030
         Left            =   128
         TabIndex        =   7
         Top             =   960
         Width           =   8850
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   6784
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6784
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Types"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   1965
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   8895
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
      Option Explicit
 
Function FindFilesAPI(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
      Dim FileName As String   ' Walking filename variable...
      Dim DirName As String    ' SubDirectory Name
      Dim dirNames() As String ' Buffer for directory name entries
      Dim nDir As Integer   ' Number of directories in this path
      Dim i As Integer      ' For-loop counter...
      Dim hSearch As Long   ' Search Handle
      Dim WFD As WIN32_FIND_DATA
      Dim Cont As Integer
 
      If Right(path, 1) <> "\" Then path = path & "\"
      ' Search for subdirectories.
      nDir = 0
      ReDim dirNames(nDir)
      Cont = True
      hSearch = FindFirstFile(path & "*", WFD)
      If hSearch <> INVALID_HANDLE_VALUE Then
         Do While Cont
            DirName = StripNulls(WFD.cFileName)
            ' Ignore the current and encompassing directories.
            If (DirName <> ".") And (DirName <> "..") Then
               ' Check for directory with bitwise comparison.
               If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                  dirNames(nDir) = DirName
                  DirCount = DirCount + 1
                  nDir = nDir + 1
                  ReDim Preserve dirNames(nDir)
               End If
            End If
            Cont = FindNextFile(hSearch, WFD)  ' Get next subdirectory.
         Loop
         Cont = FindClose(hSearch)
      End If
 
      ' Walk through this directory and sum file sizes.
      hSearch = FindFirstFile(path & SearchStr, WFD)
      Cont = True
      If hSearch <> INVALID_HANDLE_VALUE Then
         While Cont
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
               FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
               FileCount = FileCount + 1
               List1.AddItem path & FileName
            End If
            Cont = FindNextFile(hSearch, WFD)  ' Get next file
         Wend
         Cont = FindClose(hSearch)
         DoEvents
      End If
 
      ' If there are sub-directories...
      If nDir > 0 Then
         ' Recursively walk into them...
         For i = 0 To nDir - 1
           FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
         Next i
      End If
End Function

Private Sub chkLegend_Click()
    chtResults.ShowLegend = chkLegend.Value
End Sub

Private Sub chtResults_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
    lblSelect = UCase$(chtResults.DataGrid.ColumnLabel(Series, 1)) & " (" & Format(FileSize(Series - 1), "#,###,###,##0") & " Megabytes)"
End Sub

Private Sub cmdGraph_Click()
    Dim t As Integer
    Dim SearchPath As String, FindStr As String
    Dim DirPart As String
    Dim NameSize As String
    Dim NamePart As String
    gStop = 0
    fmeResults.Visible = False
    chtResults.Visible = True
    lblSelect.Visible = True
    pbrResults.Visible = True
    Command1.Visible = False
    lblSelect.Caption = ""
    fmeGraph.Visible = True
    chkLegend.Visible = True
    chtResults.ColumnCount = Dir1.ListCount
    For t = 0 To Dir1.ListCount - 1 ' clear present chart
        chtResults.Column = t + 1
        chtResults.Data = 0
    Next t
    chtResults.RowCount = 1
    chtResults.DataGrid.ColumnLabelCount = Dir1.ListCount
    If Dir1.ListCount > 1 Then pbrResults.Max = Dir1.ListCount - 1
    For t = 0 To Dir1.ListCount - 1
        Text3.Text = ""
        Text4.Text = ""
        DirPart = Dir1.List(-1)
        NameSize = Len(DirPart)
        Screen.MousePointer = vbHourglass
        List1.Clear
        SearchPath = Dir1.List(t)
        FindStr = Text2.Text
        FileSize(t) = (FindFilesAPI(SearchPath, FindStr, numFiles(t), NumDirs(t))) / 1048576
        
        chtResults.Column = t + 1
        chtResults.Data = FileSize(t)
        NamePart = Right$(SearchPath, Len(SearchPath) - NameSize)
        
        chtResults.DataGrid.ColumnLabel(t + 1, 1) = Left$(NamePart, Len(NamePart) - 1)
        pbrResults.Value = t
    Next t
    Screen.MousePointer = vbDefault
    pbrResults.Visible = False
    chtResults.Legend.Location.LocationType = VtChLocationTypeRight
    'chtResults.Legend.Location.Rect.Max.X = 500
    cmdPrint.Enabled = True
End Sub

Private Sub cmdPrint_Click()
    Printer.Orientation = vbPRORLandscape
    frmFiles.PrintForm
    Printer.Orientation = vbPRORPortrait
End Sub

Private Sub cmdStop_Click()
    gStop = 1
End Sub

Private Sub cmdScale_Click()
    
End Sub

Private Sub Command1_Click()
    Text3.Text = ""
    Text4.Text = ""
    Dim SearchPath As String, FindStr As String
    Dim FileSize As Long
    Dim numFiles As Integer, NumDirs As Integer
 
    Screen.MousePointer = vbHourglass
    List1.Clear
    SearchPath = Text1.Text
    FindStr = Text2.Text
    FileSize = FindFilesAPI(SearchPath, FindStr, numFiles, NumDirs)
    Text3.Text = numFiles & " Files found in " & NumDirs + 1 & " Directories"
    Text4.Text = SearchPath & " = " & Format(FileSize, "#,###,###,##0") & " Bytes"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = vbArrow
End Sub

Private Sub Dir1_Click()
    Text1 = Dir1.List(Dir1.ListIndex)
    
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
    Text1.Text = Dir1.path
End Sub

Private Sub Form_Load()
    frmSplash.Show vbModal
    Drive1.Drive = "C"
    Dir1.path = "C:\"
    chtResults.ColumnCount = 10
    chtResults.RowCount = 1
End Sub

Private Sub optChart_Click(Index As Integer)
    chtResults.chartType = Index
End Sub

