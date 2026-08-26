VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SecSearch"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   12705
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   4680
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtFilename 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   12495
   End
   Begin VB.ListBox lstFiles 
      Height          =   4740
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   6615
   End
   Begin VB.CheckBox chkSubfolders 
      Caption         =   "Subfolders"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtElapsed 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtMatched 
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtSearched 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtDrive 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selected filename"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Headers found within document"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Files matching filespec"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   6615
   End
   Begin VB.Label lblHeaders 
      Height          =   4695
      Left            =   6840
      TabIndex        =   10
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "Elapsed"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Matched"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Searched"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const vbDot = 46
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const vbBackslash = "\"
Private Const ALL_FILES = "*.*"

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type FILE_PARAMS
   bRecurse As Boolean
   nCount As Long
   nSearched As Long
   sFileNameExt As String
   sFileRoot As String
End Type

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
                                           (ByVal lpFileName As String, _
                                            lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
                                          (ByVal hFindFile As Long, _
                                           lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Declare Function PathMatchSpec Lib "shlwapi" Alias "PathMatchSpecW" _
                                           (ByVal pszFileParam As Long, _
                                            ByVal pszSpec As Long) As Long

Private fp As FILE_PARAMS  'holds search parameters

Private Sub cmdSearch_Click()

   Dim tstart As Single   'timer var for this routine only
   Dim tend As Single     'timer var for this routine only
   
   txtSearched.Text = ""
   txtMatched.Text = ""
   txtElapsed.Text = ""
   lstFiles.Clear
   Screen.MousePointer = vbHourglass
   
   With fp
      .sFileRoot = QualifyPath(txtDrive.Text) 'start path
      .sFileNameExt = "*.doc"                 'file type(s) of interest
      .bRecurse = chkSubfolders.Value = 1     'True = recursive search
      .nCount = 0                             'results
      .nSearched = 0                          'results
   End With
  
   tstart = GetTickCount()
   Call SearchForFiles(fp.sFileRoot)
   tend = GetTickCount()
   Screen.MousePointer = vbDefault
   
   txtSearched.Text = Format$(fp.nSearched, "###,###,###,##0")
   txtMatched.Text = Format$(fp.nCount, "###,###,###,##0")
   txtElapsed.Text = FormatNumber((tend - tstart) / 1000, 2) & "  seconds"
                                    
End Sub


Private Sub SearchForFiles(sRoot As String)

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
  
   hFile = FindFirstFile(sRoot & ALL_FILES, WFD)
  
   If hFile <> INVALID_HANDLE_VALUE Then
   
      Do
                  
        'if a folder, and recurse specified, call
        'method again
         If (WFD.dwFileAttributes And vbDirectory) Then
            If Asc(WFD.cFileName) <> vbDot Then

             If fp.bRecurse Then
                  SearchForFiles sRoot & TrimNull(WFD.cFileName) & vbBackslash
               End If
            End If
            
         Else
         
           'must be a file..
            If MatchSpec(WFD.cFileName, fp.sFileNameExt) Then
               fp.nCount = fp.nCount + 1
               lstFiles.AddItem sRoot & TrimNull(WFD.cFileName)
               lstFiles.Refresh
            End If  'If MatchSpec
      
         End If 'If WFD.dwFileAttributes
      
         fp.nSearched = fp.nSearched + 1
         
         DoEvents
      
      Loop While FindNextFile(hFile, WFD)
   
   End If 'If hFile
  
   Call FindClose(hFile)

End Sub


Private Function QualifyPath(sPath As String) As String

   If Right$(sPath, 1) <> vbBackslash Then
      QualifyPath = sPath & vbBackslash
   Else
      QualifyPath = sPath
   End If
      
End Function


Private Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlen(StrPtr(startstr)))
   
End Function


Private Function MatchSpec(sFile As String, sSpec As String) As Boolean

   MatchSpec = PathMatchSpec(StrPtr(sFile), StrPtr(sSpec))
   
End Function

Private Sub ListBox1_Click()
    
End Sub

Public Function GetHeaderInfo(strFilename As String) As String
    Dim CharacterCount As Integer
    Dim iCharacters As Integer
    Dim iSectionCount As Integer
    Dim iSections As Integer
    Dim strHeader As String
    Dim oWord As Word.Application
    Dim oDoc As Document
    
    Screen.MousePointer = vbHourglass
    
    Set oWord = CreateObject("Word.Application")
    Set oDoc = oWord.Documents.Open(strFilename, , , , , , , , , , , False)
    
    iSectionCount = oDoc.Sections.Count
    
    For iSections = 1 To iSectionCount
        strHeader = strHeader & "[Section " & iSections & "]" & vbCrLf
        CharacterCount = oDoc.Sections.Item(iSections).Headers.Item(1).Range.Characters.Count
        
        For iCharacters = 1 To CharacterCount
            strHeader = strHeader & oDoc.Sections.Item(iSections).Headers.Item(1).Range.Characters.Item(iCharacters)
        Next iCharacters
        
        strHeader = strHeader & vbCrLf
        DoEvents
    Next iSections

    oDoc.Close False
    oWord.Quit
    
    Set oDoc = Nothing
    Set oWord = Nothing

    GetHeaderInfo = strHeader
    Screen.MousePointer = vbDefault
    
End Function


Private Sub lstFiles_Click()
    txtFilename.Text = lstFiles.List(lstFiles.ListIndex)
    lblHeaders.Caption = GetHeaderInfo(lstFiles.List(lstFiles.ListIndex))
End Sub
