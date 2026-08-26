VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Finder"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9705
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tabFiiles 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "All files (*.*)"
      TabPicture(0)   =   "frmMain.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtbFiles"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "FileInfo"
      TabPicture(1)   =   "frmMain.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "rtbVersionInfo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "All Applications"
      TabPicture(2)   =   "frmMain.frx":5C4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstApplications"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdExport"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Corporate Applications"
      TabPicture(3)   =   "frmMain.frx":5C66
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstCorporate"
      Tab(3).Control(1)=   "cmdExportCorporate"
      Tab(3).ControlCount=   2
      Begin VB.CommandButton cmdExportCorporate 
         Caption         =   "Export"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66600
         TabIndex        =   6
         Top             =   4680
         Width           =   855
      End
      Begin VB.ListBox lstCorporate 
         Height          =   3960
         ItemData        =   "frmMain.frx":5C82
         Left            =   -74760
         List            =   "frmMain.frx":5C84
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   9015
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66600
         TabIndex        =   4
         Top             =   4680
         Width           =   855
      End
      Begin VB.ListBox lstApplications 
         Height          =   3960
         ItemData        =   "frmMain.frx":5C86
         Left            =   -74760
         List            =   "frmMain.frx":5C88
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   9015
      End
      Begin RichTextLib.RichTextBox rtbFiles 
         Height          =   3960
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6985
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":5C8A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbVersionInfo 
         Height          =   3960
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6985
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":5D06
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   9840
      X2              =   0
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   9840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuScan 
         Caption         =   "Scan"
         Begin VB.Menu mnuDrive 
            Caption         =   "C"
            Index           =   0
         End
         Begin VB.Menu mnuDrive 
            Caption         =   "D"
            Index           =   1
         End
         Begin VB.Menu mnuDrive 
            Caption         =   "Path..."
            Index           =   2
         End
         Begin VB.Menu mnuDrive 
            Caption         =   "Network..."
            Index           =   3
         End
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Const ATTR_NORMAL = 0
   Const ATTR_READONLY = 1
   Const ATTR_HIDDEN = 2
   Const ATTR_SYSTEM = 4
   Const ATTR_VOLUME = 8
   Const ATTR_DIRECTORY = 16
   Const ATTR_ARCHIVE = 32
   Dim CorpAppExes() As String
   Dim CorpAppNames() As String
   Dim SearchName As String
   
Private Function FindFile(ByVal sFol As String, sFile As String, nDirs As Integer, nFiles As Integer) As Long
    Dim tFld As Folder, tFil As File, FileName As String, TotalSize As Long, FilenameToAdd As String
    Dim varPath As String
    Dim fso As New FileSystemObject
    Dim fld As Folder, DirName As String, Filesize As String, Fileinfo As String
    Dim Readonly As String, Hidden As String, System As String, Archive As String
    Dim FileVersionInfo As String, appName As String, appVersion As String
    Dim AppFound As Boolean, FoundCorpApp As Boolean
    Dim i As Integer
    
    On Error Resume Next  ' Required for files that are in use
    
    Set fld = fso.GetFolder(sFol)
    varPath = fso.BuildPath(fld.Path, sFile)
    FileName = UCase(Dir(varPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly Or vbArchive))
    While Len(FileName) <> 0 And FileName <> "." And FileName <> ".."
        If DirName <> fld.Path Then
            If Len(rtbFiles.Text) <> 0 Then rtbFiles.Text = rtbFiles.Text & vbCrLf
            rtbFiles.Text = rtbFiles.Text & " Directory of " & fld.Path & vbCrLf
            DirName = fld.Path
        End If
        FilenameToAdd = UCase(fso.BuildPath(fld.Path, FileName))
        TotalSize = TotalSize + FileLen(fso.BuildPath(fld.Path, FileName))
        nFiles = nFiles + 1
        Filesize = Format(Format(FileLen(FilenameToAdd), "###,###,### "), "@@@@@@@@@@@@@")
        If Trim(Filesize) = "" Then Filesize = "           0 "
        Fileinfo = Format(FileDateTime(FilenameToAdd), "dd mmm yyyy  hh:nn AMPM") & " " & Filesize & FileName
        
        Readonly = IIf(GetAttr(FilenameToAdd) And ATTR_READONLY, "R", "")
        Hidden = IIf(GetAttr(FilenameToAdd) And ATTR_HIDDEN, "H", "")
        System = IIf(GetAttr(FilenameToAdd) And ATTR_SYSTEM, "S", "")
        Archive = IIf(GetAttr(FilenameToAdd) And ATTR_ARCHIVE, "A", "")
        
        Fileinfo = Fileinfo & " " & Readonly & Hidden & System & Archive
        
        rtbFiles.Text = rtbFiles.Text & Fileinfo & vbCrLf
        
        FileVersionInfo = GetFileinfo(FilenameToAdd, appName, appVersion)
        If FileVersionInfo <> "" Then
            If Len(rtbVersionInfo.Text) <> 0 Then
                rtbVersionInfo.Text = rtbVersionInfo.Text & vbCrLf
            End If
            rtbVersionInfo.Text = rtbVersionInfo.Text & FileVersionInfo
            If appName <> "" Then
                ' Look through found applications, and check if it is already listed
                For lp = 0 To lstApplications.ListCount - 1
                    If lstApplications.List(lp) = appName & " " & appVersion Then AppFound = True
                Next lp
                ' Not listed, so add to list
                If Not AppFound Then lstApplications.AddItem appName & " " & appVersion
                AppFound = False
            End If
        End If
        
        ' Go through in-memory list of Corporate Applications, and check if this app is one of those
        For i = 0 To UBound(CorpAppExes()) - 1
            If UCase(FileName) = UCase(CorpAppExes(i)) Then
                FoundCorpApp = True
                Exit For
            End If
            'Debug.Print CorpAppExes(i)
        Next i
        
        If FoundCorpApp = True Then
        '    ' Yes it is - add it to the list, but only if it isn't already there.
        '    ' Look through found CORPORATE applications, and check if it is already listed
        '    For lp = 0 To lstApplications.ListCount - 1
        '        If lstCorporate.List(lp) = CorpAppNames(i) Then AppFound = True
        '    Next lp
        '    ' Not listed, so add to list
            If Not AppFound Then lstCorporate.AddItem CorpAppNames(i) & " -" & FileName & " (" & DirName & ")"
            FoundCorpApp = False
        '    AppFound = False
        End If
        FileName = Dir()  ' Get next file
        DoEvents
    Wend
    lblStatus = "Searching " & vbCrLf & fld.Path & "..."
    nDirs = nDirs + 1
    If fld.SubFolders.Count > 0 Then
        For Each tFld In fld.SubFolders
            DoEvents
            TotalSize = TotalSize + FindFile(tFld.Path, sFile, nDirs, nFiles)
        Next
    End If
    FindFile = TotalSize
End Function

Private Sub cmdExport_Click()
    Open "c:\apps (" & SearchName & ").htm" For Output As #1
    Screen.MousePointer = vbHourglass
    Print #1, "<html>"
    Print #1, "    <head>"
    Print #1, "    <title>Applications found</title>"
    Print #1, "    </head>"
    Print #1, "    <body>"
    Print #1, "        <div align='center'>"
    Print #1, "        <h2>Applications found on " & SearchName & "</h2>"
    Print #1, "        <table align='center' border='0' cellborder='0'>"
    
    For lp = 0 To lstApplications.ListCount - 1
        Print #1, "            <tr>"
        Print #1, "                <td>" & lstApplications.List(lp) & "</td>"
        Print #1, "            </tr>"
    Next lp
    
    Print #1, "        </table"
    Print #1, "        </div>"
    Print #1, "        <br>"
    Print #1, "        <div align='left'>"
    Print #1, "        Note:&nbsp;&nbsp;The presence of a single .exe file from an application does not necessarily mean the whole application is present."
    Print #1, "        </div>"
    Print #1, "    </body>"
    Print #1, "</html>"
    Screen.MousePointer = vbDefault
    Close 1
    MsgBox "List saved to c:\applications (" & SearchName & ").htm"
End Sub

Private Sub cmdExportCorporate_Click()
    Open "c:\c_apps (" & SearchName & ").htm" For Output As #1
    Screen.MousePointer = vbHourglass
    Print #1, "<html>"
    Print #1, "    <head>"
    Print #1, "    <title>Corporate Applications found</title>"
    Print #1, "    </head>"
    Print #1, "    <body>"
    Print #1, "        <div align='center'>"
    Print #1, "        <h2>Corporate Applications found on " & SearchName & "</h2>"
    Print #1, "        <table align='center' border='0' cellborder='0'>"
    
    For lp = 0 To lstCorporate.ListCount - 1
        Print #1, "            <tr>"
        Print #1, "                <td>" & lstCorporate.List(lp) & "</td>"
        Print #1, "            </tr>"
    Next lp
    
    Print #1, "        </table"
    Print #1, "        </div>"
    Print #1, "        <br>"
    Print #1, "        <div align='left'>"
    Print #1, "        Note:&nbsp;&nbsp;The presence of a single .exe file from an application does not necessarily mean the whole application is present."
    Print #1, "        </div>"
    Print #1, "    </body>"
    Print #1, "</html>"
    Screen.MousePointer = vbDefault
    Close 1
    MsgBox "List saved to c:\corporate applications (" & SearchName & ").htm"
End Sub

Private Sub mnuDrive_Click(Index As Integer)
    Dim nDirs As Integer, nFiles As Integer, lSize As Long
    Dim sDir As String, sSrchString As String
    Dim aPath As String
    Dim Apps() As String
    Dim strInfo As String
    Dim i As Integer
    
    aPath = App.Path
    If Right(aPath, 1) <> "\" Then aPath = aPath & "\"
    
    ' Load arrays with Corporate Application information
    ' csv is in format:  Executable, Names, Frontend, Backend, Description
    
    Open aPath & "BHP.csv" For Input As #1
        Do Until EOF(1)
            Line Input #1, strInfo
            Erase Apps
            strInfo = Replace(strInfo, Chr(34), "")
            Apps() = Split(strInfo, ",")
            i = i + 1
            ReDim Preserve CorpAppExes(i)
            ReDim Preserve CorpAppNames(i)
            CorpAppExes(i) = UCase(Apps(0))
            CorpAppNames(i) = UCase(Apps(1))
        Loop
    Close 1
    
    lstApplications.Clear
    lstCorporate.Clear
    rtbFiles.Text = ""
    rtbVersionInfo.Text = ""
    
    Select Case Index
        Case 0
            sDir = "c:\"
            SearchName = Environ$("computername") & " Drive C:"
        Case 1
            sDir = "d:\"
            SearchName = Environ$("computername") & " Drive D:"
        Case 2
            sDir = InputBox("Enter a search path", "Input required", "C:\Program Files")
            If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
            SearchName = Environ$("computername") & " " & sDir
        Case 3
            sDir = InputBox("Enter a Network path", "Input required", "")
            If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
    End Select
    Do
        SearchName = InputBox("Enter a Name for this search", "Input required", "")
    Loop Until InStr(1, SearchName, "\") = 0
    sSrchString = "*.*"
    
    'Get files
    If sDir <> "" And sDir <> "\" Then
        rtbFiles.Text = ""
        rtbVersionInfo = ""
        lstApplications.Clear
        lSize = FindFile(sDir, sSrchString, nDirs, nFiles)
        If lstApplications.ListCount > 0 Then
            cmdExport.Enabled = True
            End If
        If lstCorporate.ListCount > 0 Then
            cmdExportCorporate.Enabled = True
        End If
        rtbFiles.Text = rtbFiles.Text & vbCrLf & "Summary:" & vbCrLf & nFiles & " files in " & nDirs & " Directories - " & Format(lSize, "###,###,###") & " bytes (" & Int(lSize / 1024) & " Kb)"
    Else
        cmdExport.Enabled = False
        cmdExportCorporate.Enabled = False
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub
