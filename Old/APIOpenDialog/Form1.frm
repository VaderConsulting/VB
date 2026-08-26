VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive letter to UNC Path Convertor"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   795
   ClientWidth     =   9855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Subs"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ListBox lstSubs 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtFile 
      Height          =   4815
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lblDriveLetterSubs 
      Alignment       =   2  'Center
      Caption         =   "Drive Substitutions"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label lblInfo 
      Caption         =   "Open a text file to convert embedded Drive letters to UNC paths. (File | Open)"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuOpenFolder 
         Caption         =   "Open Folder"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_RETURNFSANCESTORS = 8
Private Const BIF_EDITBOX = 16
Private Const BIF_NEWDIALOGSTYLE = 64
Private Const BIF_BROWSEINCLUDEFILES = 16384

Private Const MAX_PATH = 260

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

 Public Filename As String

Private Sub cmdClear_Click()
    lstSubs.Clear
    cmdClear.Enabled = False
End Sub

Private Sub cmdSave_Click()
    ' Save the modified file.
    If Filename <> "" Then
        If Dir(Filename, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "" Then
            On Error GoTo FileOpen
                Name Filename As Filename & ".backup"
            On Error GoTo 0
            
            On Error GoTo FileSave
                Open Filename For Output As #1
                    Print #1, txtFile.Text
                Close 1
            On Error GoTo 0
        Else
            MsgBox "This file has been deleted or renamed.  (" & Filename & ")", vbExclamation, "Error"
        End If
    End If
    
    lblInfo.Caption = "Operation complete.  Select another file to convert, or close.  Click 'Clear Subs' if you wish to remove this list of drive substitutions."
    Filename = ""
    cmdSave.Enabled = False
    cmdStart.Enabled = False
    txtFile.Text = ""
    cmdClear.Enabled = True
    Exit Sub
' Error during rename
FileOpen:
    MsgBox "There has been an error renaming '" & Filename & "' to '" & Filename & ".backup'." & vbCrLf & vbCrLf & _
           "Ensure you have rights to modify this file, and that it is not already open." & vbCrLf & vbCrLf & _
           "Details: " & Err.Description & " (" & Err.Number & ")", vbCritical, "Error"
    Err.Clear
    Close
    Exit Sub

' Error during save
FileSave:
    MsgBox "There has been an error saving '" & Filename & "'." & vbCrLf & vbCrLf & _
           "Ensure you have rights to create this file, and that it does not already exist." & vbCrLf & vbCrLf & _
           "Details: " & Err.Description & " (" & Err.Number & ")", vbCritical, "Error"
    Err.Clear
    Close
End Sub

Private Sub cmdStart_Click()
    ' This is where the hard work is done.
    Dim lp As Integer, ServerandShare As String, DriveLetter As String
    
    For lp = 0 To lstSubs.ListCount - 1
        ServerandShare = Trim(Right(lstSubs.List(lp), Len(lstSubs.List(lp)) - 5))
        
        If Right(ServerandShare, 1) <> "\" Then
            ServerandShare = ServerandShare & "\"
        End If
        
        DriveLetter = Left(lstSubs.List(lp), 3)
        txtFile.Text = Replace(txtFile.Text, DriveLetter, ServerandShare)
    Next lp
    txtFile.Visible = True
    cmdStart.Enabled = False
    cmdSave.Enabled = True
    lblInfo.Caption = "Click the save button to save this file.  Note that you may manually modify the file if you wish."
End Sub

Private Sub lstSubs_Click()
    Dim SelectedFolder As String, CompletedFlag As Boolean
    
    szTitle = "Select target folder"
    SelectedFolder = GetFolderAPI(szTitle)
    
    If Right(SelectedFolder, 1) <> "\" Then
            SelectedFolder = SelectedFolder & "\"
        End If
    
    lstSubs.List(lstSubs.ListIndex) = Left(lstSubs.List(lstSubs.ListIndex), 5) & " " & SelectedFolder
    
    ' Check if each drive letter has a substitution
    CompletedFlag = True
    For lp = 0 To lstSubs.ListCount - 1
        If Right(lstSubs.List(lp), 1) = "?" Then
            CompletedFlag = False
        End If
    Next lp
    
    ' Ready to go?
    If CompletedFlag Then
        lblInfo.Caption = "Click the start button to replace the drive letters found with the server and share substitutions entered."
        cmdStart.Enabled = True
    End If
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Drive Letter to UNC Path Conversion Application" + vbCrLf + vbCrLf + _
    "Licensed royalty free to Rio Tinto Exploration" + vbCrLf + vbCrLf + _
    "(c) 2002 D. Robinson", vbInformation + vbOKOnly, "About"
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuOpenFile_Click()
    Dim OpenFile As OPENFILENAME, SelectedFile As String
    Dim lReturn As Long
    Dim sFilter As String, maxFileSize As Long
    
    maxFileSize = 1024 ' Kilobytes
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hWndOwner = frmMain.hWnd
    OpenFile.hInstance = App.hInstance
    sFilter = "All Files (*.*)" & Chr(0) & "*.*" & Chr(0) & "Text Files (*.txt)" & Chr(0) & "*.TXT" & Chr(0) & "Workspace Files (*.wor)" & Chr(0) & "*.WOR" & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = App.Path
    OpenFile.lpstrTitle = "Select file to convert"
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        'The User pressed the Cancel Button
    Else
        SelectedFile = Trim(OpenFile.lpstrFile)
        If SelectedFile <> "" Then
            Filename = SelectedFile
            ReplaceLetterswithUNCs Filename
        Else
            Filename = ""
        End If
        
        ' Remove CHR(0)'s
        Filename = Replace(Filename, Chr(0), "")
        
        ' Check filesize
        If FileLen(Filename) / 1024 > maxFileSize Then
            MsgBox "This file is too large to convert.  The maximum size to convert is 1Mb.", vbExclamation, "Error"
            Filename = ""
        End If
    End If

End Sub

Private Sub mnuOpenFolder_Click()
    Dim SelectedFolder As String
    
    szTitle = "Select target folder"
    SelectedFolder = GetFolderAPI(szTitle)
End Sub

Private Function GetFolderAPI(ByVal strTitle As String, Optional flags As Long = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_NEWDIALOGSTYLE + BIF_EDITBOX) As String
    'Opens a Treeview control that displays the directories in a computer
    
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    szTitle = strTitle
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = flags
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        GetFolderAPI = sBuffer
    End If
End Function

Private Sub ReplaceLetterswithUNCs(strFilename As String)
    Dim SingleLine As String, FileContents As String
    Dim ColonPos As Integer, LeftMost As String, RightMost As String, DriveLetter As String
    Dim lp As Integer, FoundFlag As Boolean
    
    Open strFilename For Input As #1
        Do Until EOF(1)
            Line Input #1, SingleLine
            
            ' Look for a drive letter
            If SingleLine Like "*?:\*" Then
                ColonPos = InStr(1, SingleLine, ":")
                LeftMost = Left(SingleLine, ColonPos - 2)
                RightMost = Right(SingleLine, Len(SingleLine) - (ColonPos))
                DriveLetter = Mid(SingleLine, ColonPos - 1, 3)
                
                ' Ensure it isn't already listed
                For lp = 0 To lstSubs.ListCount - 1
                    If Left(lstSubs.List(lp), 3) = DriveLetter Then
                        FoundFlag = True
                    End If
                Next lp
                
                ' If not already in list, add to list
                If Not FoundFlag Then
                    lstSubs.AddItem DriveLetter & " = ?"
                End If
                FoundFlag = False
            End If
            
            FileContents = FileContents & SingleLine & vbCrLf
        Loop
    Close 1
    txtFile.Text = FileContents
    lblInfo.Caption = "Select each of the drives below, and when prompted, enter the network path to replace the drive letter with."
End Sub
