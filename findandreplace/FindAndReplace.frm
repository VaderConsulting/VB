VERSION 5.00
Begin VB.Form frmFindAndReplace 
   Caption         =   "FindAndReplace"
   ClientHeight    =   5685
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtReplace 
      Height          =   1215
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txtFind 
      Height          =   1215
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   360
      Width           =   3975
   End
   Begin VB.ComboBox cboPattern 
      Height          =   315
      ItemData        =   "FindAndReplace.frx":0000
      Left            =   0
      List            =   "FindAndReplace.frx":0002
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.DriveListBox DriveList 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.DirListBox DirList 
      Height          =   1155
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.FileListBox FileList 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace With:"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblFind 
      Caption         =   "Find:"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmFindAndReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_QuitEarly As Boolean

' In the file file_name, replace occurrances of from_text
    ' with to_text. Return True if the string appeared and
    ' was replaced, False if the string was not in this file.

Private Function ReplaceInFile(ByVal file_name As String, ByVal from_text As String, ByVal to_text As String) As Boolean
    Dim fnum As Integer
    Dim file_text As String
    
    On Error GoTo ReplaceError
    
    ' Read the file.
    fnum = FreeFile
    Open file_name For Input As fnum
    file_text = Input$(LOF(fnum), #fnum)
    Close #fnum
    
    ' See if the text appears.
    If InStr(file_text, from_text) > 0 Then
        ' Replace the text.
        file_text = Replace(file_text, from_text, to_text)
        
        ' Rewrite the file.
        fnum = FreeFile
        Open file_name For Output As fnum
        Print #fnum, file_text;
        Close #fnum
        
        ReplaceInFile = True
    End If
    Exit Function
    
ReplaceError:
    Select Case MsgBox("Error " & Err.Number & " processing file " & file_name & vbCrLf & Err.Description & "Try again?", vbYesNoCancel)
        Case vbYes
            Resume
        Case vbNo
            Exit Function
        Case Else
            m_QuitEarly = True
            Exit Function
    End Select
End Function

Private Sub cmdReplace_Click()
    Dim from_text As String
    Dim to_text As String
    Dim dir_name As String
    Dim patterns As Variant
    Dim file_name As String
    Dim i As Integer
    Dim results As String
    
    ' Get the text to find and replace.
    from_text = txtFind.Text
    to_text = txtReplace.Text
    
    ' Get the directory name.
    dir_name = FileList.Path
    If Right$(dir_name, 1) <> "\" Then dir_name = dir_name & "\"
    
    ' Get the file patterns.
    patterns = Split(FileList.Pattern, ";")
    
    results = "Files:"
    
    ' Repeat for each pattern.
    m_QuitEarly = False
    For i = LBound(patterns) To UBound(patterns)
        ' Add the pattern to the file name.
        file_name = Dir$(dir_name & patterns(i))
        Do While Len(file_name) > 0
            ' Process this file.
            If ReplaceInFile(dir_name & file_name, from_text, to_text) Then
                results = results & " " & file_name
            End If
            If m_QuitEarly Then Exit For
            
            ' Get the next file.
            file_name = Dir$()
            Loop
        Next i
        
        MsgBox results
    End Sub

Private Sub DirList_Change()
    FileList.Path = DirList.Path
End Sub

Private Sub DriveList_Change()
    'On Error GoTo DriveError
    DirList.Path = DriveList.Drive
    Exit Sub
    
DriveError:
    DriveList.Drive = DirList.Path
    Exit Sub
End Sub
Private Sub Form_Load()
    cboPattern.AddItem "Text (*.txt)"
    cboPattern.AddItem "HTML (*.htm;*.html)"
    cboPattern.AddItem "VB Forms (*.frm)"
    cboPattern.AddItem "VB Modules (*.bas)"
    cboPattern.AddItem "VB Classes (*.cls)"
    cboPattern.AddItem "All VB Files (*.frm;*.bas;*.cls)"
    cboPattern.AddItem "C Header Files (*.h)"
    cboPattern.AddItem "All C Files (*.h;*.cpp;*.cxx)"
    cboPattern.AddItem "All Files (*.*)"
    cboPattern.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Const GAP = 60
    
    Dim wid As Single
    Dim hgt As Single
    Dim X As Single
    
    If WindowState = vbMinimized Then Exit Sub
    
    wid = DriveList.Width
    DriveList.Move GAP, GAP, wid
    cboPattern.Move GAP, ScaleHeight - cboPattern.Height, wid
    
    hgt = (cboPattern.Top - DriveList.Top - DriveList.Height - 3 * GAP) / 2
    If hgt < 100 Then hgt = 100
    DirList.Move GAP, DriveList.Top + DriveList.Height + GAP, wid, hgt
    FileList.Move GAP, DirList.Top + DirList.Height + GAP, wid, hgt
    
    hgt = (ScaleHeight - 2 * GAP - 2 * lblFind.Height - cmdReplace.Height) / 2
    If hgt < 120 Then hgt = 120
    X = cboPattern.Left + cboPattern.Width + GAP
    wid = ScaleWidth - X
    If wid < 120 Then wid = 120
    
    lblFind.Move X, 0
    txtFind.Move X, lblFind.Top + lblFind.Height, wid, hgt
    lblReplace.Move X, txtFind.Top + txtFind.Height + GAP
    txtReplace.Move X, lblReplace.Top + lblReplace.Height, wid, hgt
    X = X + (wid - cmdReplace.Width) / 2
    cmdReplace.Move X, txtReplace.Top + txtReplace.Height + GAP
End Sub

Private Sub cboPattern_Click()
    Dim pat As String
    Dim p1 As Integer
    Dim p2 As Integer
    
    pat = cboPattern.List(cboPattern.ListIndex)
    p1 = InStr(pat, "(")
    p2 = InStr(pat, ")")
    FileList.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub
