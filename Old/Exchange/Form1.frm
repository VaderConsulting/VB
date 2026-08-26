VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tracking Log Extract tool"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDelete 
      Caption         =   "Delete destination file before commencing extract"
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   4080
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.FileListBox filLogs 
      Height          =   3210
      Left            =   7200
      Pattern         =   "*.log"
      TabIndex        =   22
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Set Source Dir"
      Height          =   735
      Left            =   120
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   615
      Left            =   6360
      Picture         =   "Form1.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   720
      Width           =   615
   End
   Begin VB.ListBox lstSource 
      Height          =   2985
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   6135
   End
   Begin VB.Frame fmeDates 
      Caption         =   "Dates"
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   4935
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add to source"
         Enabled         =   0   'False
         Height          =   855
         Left            =   2760
         Picture         =   "Form1.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ListBox lstLogs 
         Height          =   840
         Left            =   2760
         MultiSelect     =   2  'Extended
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   21
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Enabled         =   0   'False
         Height          =   855
         Left            =   840
         Picture         =   "Form1.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtTo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtFrom 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblFormat 
         Caption         =   "Date format:  YYYYMMDD"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblTo 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblDate1 
         Alignment       =   2  'Center
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fmeExtract 
      Caption         =   "Data to extract"
      Height          =   855
      Left            =   5160
      TabIndex        =   6
      Top             =   4320
      Width           =   1815
      Begin VB.OptionButton optMode 
         Caption         =   "All Data"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Subset of data"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   5
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   6360
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   6480
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txtDestination 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "c:\temp\exchangelogs.csv"
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   6600
      Width           =   6855
   End
   Begin VB.Label lblDestination 
      Caption         =   "Destination"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblSource 
      Caption         =   "Source file(s)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sourcepath As String


Private Sub cmdAdd_Click()
    For lp = 0 To lstLogs.ListCount - 1
        ' If selected, then add to lstsource
        If lstLogs.Selected(lp) = True Then
            lstSource.AddItem filLogs.Path & "\" & lstLogs.List(lp)
        End If
        
    Next lp
    lstLogs.Clear
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    With cdlOpen
        .DefaultExt = ".csv"
        .Filter = "*.log"
        .InitDir = Sourcepath
    End With
    If Index = 0 Then
        cdlOpen.ShowOpen
        lstSource.AddItem cdlOpen.FileName
    End If
    If Index = 1 Then
        cdlOpen.ShowSave
    End If
    
End Sub

Private Sub cmdDate_Click(Index As Integer)
    frmDate.lblFrom = Index
    frmDate.calDates.Value = Date
    frmDate.Show vbModal
End Sub

Private Sub cmdDelete_Click()
    If lstSource.ListIndex = -1 Then Exit Sub
    lstSource.RemoveItem lstSource.ListIndex
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Public Sub cmdExtract_Click()
    ' Delete file if user has selected option to do so
    On Error Resume Next
        If chkDelete.Value = vbChecked Then
            Kill txtDestination
        End If
    On Error GoTo 0
    Screen.MousePointer = vbHourglass
    Starttime = Time
    For lp = 0 To lstSource.ListCount - 1
        DoEvents
        Dim Mode As String
        If lstSource.ListCount > 0 And txtDestination <> "" Then
            If optMode(0).Value = True Then Mode = "all"
            If optMode(1).Value = True Then Mode = "sub"
            lblStatus = "Extracting " & lstSource.List(lp)
            lblStatus.Refresh
            Extract lstSource.List(lp), txtDestination, Mode
        Else
            MsgBox "Select a valid source and destination"
        End If
    Next lp
    Screen.MousePointer = vbDefault
    TimeTaken = DateDiff("s", Starttime, Time)
    MsgBox "Extract complete" & vbCrLf & " Time to extract: " & TimeTaken & " " & "seconds", vbInformation, "Tracking Log Extract tool"
End Sub

Private Sub cmdRefresh_Click()
    lstLogs.Clear
    If txtFrom <> "" And txtTo <> "" Then
        vFromYear = Left(txtFrom, 4)
        vFromMonth = Mid(txtFrom, 5, 2)
        vFromDay = Right(txtFrom, 2)
        vFromDate = vFromDay & "/" & vFromMonth & "/" & vFromYear
        
        vToYear = Left(txtTo, 4)
        vToMonth = Mid(txtTo, 5, 2)
        vToDay = Right(txtTo, 2)
        vToDate = vToDay & "/" & vToMonth & "/" & vToYear
            
        For lp = 0 To filLogs.ListCount - 1
            vLogFilename = Left(filLogs.List(lp), Len(filLogs.List(lp)) - 4)
            vLogYear = Left(vLogFilename, 4)
            vLogMonth = Mid(vLogFilename, 5, 2)
            vLogDay = Right(vLogFilename, 2)
            vLogDate = vLogDay & "/" & vLogMonth & "/" & vLogYear
            If CDate(vLogDate) >= CDate(vFromDate) And CDate(vLogDate) <= CDate(vToDate) Then
                lstLogs.AddItem filLogs.List(lp)
            End If
        Next lp
    Else
        
    End If
    BuildPath
    If lstLogs.ListCount > 0 Then cmdAdd.Enabled = True
End Sub

Private Sub cmdSetup_Click()
    ' Ask user for source path
    Sourcepath = InputBox("Enter the source directory (UNC Name required)", "Info required", "\\exhsrv1\c$\exchsrvr\tracking.log\")
    '
    ' Extract the server name from the input string, but first ensure that \\ is part of string
    If Left(Sourcepath, 2) <> "\\" Then         ' Looking for \\
        If Left(Sourcepath, 1) = "\" Then       ' Looking for \
            Sourcepath = "\" & Sourcepath       ' Add \
        Else
            Sourcepath = "\\" & Sourcepath      ' Add \\
        End If
    End If
    BuildPath
    filLogs.Path = Sourcepath
    filLogs.Refresh
End Sub

Sub BuildPath()
    pos1 = InStr(3, Sourcepath, "\")            ' Find first occurance of '\' after the initial \\
    Servername = Mid(Sourcepath, 3, pos1 - 3)   ' Build server name
    If Sourcepath <> "" Then
        txtDestination = "c:\temp\" & Servername & "_" & txtFrom & " to " & txtTo & ".csv"
    Else
        txtDestination = "c:\temp\" & Servername & "_" & txtTo & ".csv"
    End If
End Sub

Private Sub Form_Load()
    txtTo = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00")
End Sub

Private Sub txtFrom_Change()
    If txtFrom <> "" And txtTo <> "" Then
        cmdRefresh.Enabled = True
    Else
        cmdRefresh.Enabled = False
    End If
End Sub

Private Sub txtTo_Change()
    If txtFrom <> "" And txtTo <> "" Then
        cmdRefresh.Enabled = True
    Else
        cmdRefresh.Enabled = False
    End If

End Sub
