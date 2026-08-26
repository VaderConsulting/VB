VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse for Folder"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optType 
      Caption         =   "Lyrics (*.txt)"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton optType 
      Caption         =   "MIDI"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.ListBox lstFiles 
      Height          =   2010
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   4815
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Include sub-folders"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.DirListBox dirBrowse 
      Height          =   3240
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4455
   End
   Begin VB.DriveListBox drvBrowse 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   4455
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim sDir As String
    Dim nDirs As Integer, nFiles As Integer, lSize As Long
    Dim sSrchString As String
    Dim ismIDI As Boolean, isLyrics As Boolean
    Dim Filename As String, Songname As String
    Dim intLoop As Integer
    Dim intTemp As Integer
    Dim strTemp As String
    
    lstFiles.Clear
    
    sDir = dirBrowse.Path
    If optType(0).Value = True Then
        sSrchString = "*.mid"
    End If
    If optType(1).Value = True Then
        sSrchString = "*.txt"
    End If
    
    adoConn.Open DSN
    MousePointer = vbHourglass
    lblStatus = "Searching " & UCase(sDir) & "..."
    lSize = FindFile(sDir, sSrchString, nDirs, nFiles)
    MousePointer = vbDefault
    'MsgBox "Total Size = " & lSize & " bytes"
    lblStatus = Str(nFiles) & " files found in" & Str(nDirs) & " directories"
    For intLoop = 0 To lstFiles.ListCount - 1
        ismIDI = False
        isLyrics = False
        Filename = lstFiles.List(intLoop)
        intTemp = InStrRev(Filename, "\")
        strTemp = Right(Filename, Len(Filename) - intTemp)
        Songname = Left(strTemp, Len(strTemp) - 4)
        
        Filename = UCase(Replace(Filename, "'", "''"))
        Songname = UCase(Replace(Songname, "'", "''"))
        Filename = UCase(Replace(Filename, "%", "%%"))
        Songname = UCase(Replace(Songname, "%", "%%"))
        Filename = UCase(Replace(Filename, """", """"""))
        Songname = UCase(Replace(Songname, """", """"""))
        
        If UCase(Right(Filename, 3)) = "MID" Then ismIDI = True
        If UCase(Right(Filename, 3)) = "TXT" Then isLyrics = True
        
        SQL = ""
        SQL = "SELECT Filename FROM tblSongs WHERE Filename = '" & Filename & "'"
        adoRS.Open SQL, adoConn
        If adoRS.EOF Then
            SQL = ""
            SQL = SQL & "INSERT INTO tblSongs (Name, Filename, isMIDI, isLyrics) VALUES ("
            SQL = SQL & "'" & Songname & "',"
            SQL = SQL & "'" & Filename & "',"
            SQL = SQL & ismIDI & ","
            SQL = SQL & isLyrics & ""
            SQL = SQL & ")"
            lblStatus.Caption = "Inserting " & Songname & " into Db..."
            adoConn.Execute SQL
        Else
            lblStatus = Filename & "  already exists in Database"
        End If
        adoRS.Close
        DoEvents
        frmBrowse.Refresh
    Next intLoop
    adoConn.Close
    cmdOK.Enabled = False
    MsgBox "Complete.  Inserted " & Str(nFiles) & " files into database", vbInformation + vbOKOnly
    Unload Me
End Sub

Private Sub drvBrowse_Change()
    dirBrowse.Path = drvBrowse.Drive
End Sub

Private Function FindFile(ByVal sFol As String, sFile As String, nDirs As Integer, nFiles As Integer) As Long
    Dim tFld As Folder, tFil As File, Filename As String
    
    On Error Resume Next
    Set fld = fso.GetFolder(sFol)
    Filename = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    While Len(Filename) <> 0
        FindFile = FindFile + FileLen(fso.BuildPath(fld.Path, Filename))
        nFiles = nFiles + 1
        lstFiles.AddItem fso.BuildPath(fld.Path, Filename)  ' Load ListBox
        Filename = Dir()  ' Get next file
        DoEvents
    Wend
    lblStatus = "Searching " & fld.Path & "..."
    nDirs = nDirs + 1
    
    If chkSubFolders.Value = vbChecked Then
        If fld.SubFolders.Count > 0 Then
            For Each tFld In fld.SubFolders
               DoEvents
               FindFile = FindFile + FindFile(tFld.Path, sFile, nDirs, nFiles)
            Next
        End If
    End If
    
End Function


