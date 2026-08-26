VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmINIEditor 
   Caption         =   "INI Editor"
   ClientHeight    =   6270
   ClientLeft      =   315
   ClientTop       =   1185
   ClientWidth     =   9450
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmINIEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   9450
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   8115
      TabIndex        =   9
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtPath 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   8175
      End
   End
   Begin VB.ListBox ListNames 
      Height          =   2790
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton GetSettings 
      BackColor       =   &H00FFFF00&
      Caption         =   "Open File"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.ListBox ListKey1 
      Height          =   3180
      Left            =   3120
      TabIndex        =   6
      Top             =   480
      Width           =   2895
   End
   Begin VB.ListBox ListKey2 
      Height          =   3180
      Left            =   6120
      TabIndex        =   5
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtInformation 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Add/Delete/Change"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   9255
      Begin VB.CommandButton cmdFindFirst 
         Caption         =   "Find First"
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Top             =   7000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0000FF00&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Find Next"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   6000
         TabIndex        =   11
         Top             =   1200
         Width           =   3135
         Begin VB.CommandButton cmdFindNext 
            BackColor       =   &H00FF00FF&
            Caption         =   "Find Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblFindNext 
            AutoSize        =   -1  'True
            Caption         =   "lblFindNext"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   780
         End
      End
      Begin VB.TextBox Text1 
         Height          =   1335
         HideSelection   =   0   'False
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   960
         Width           =   5775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   7000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Add/Delete/Change and then click OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblEntries 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   75
      End
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmINIEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iniFile As String
Dim File As Long
Dim buffer As String
Dim keys As String
Dim iniName As String
Dim finished As Boolean
Dim pos As Integer
Dim FavoriteFileName As String

Private Sub Form_Load()
    ChDir App.Path
    cmdFindNext.Enabled = False
End Sub

Private Sub cmdAdd_Click()
    Dim FileNum
    
    FileNum = FreeFile
    On Error Resume Next
    Open FavoriteFileName For Output As FileNum
    Print #FileNum, Text1.Text
    Close #FileNum
    FavoriteFileName = txtPath.Text
    Dim mHandle
    mHandle = FreeFile
    Open FavoriteFileName For Binary As #mHandle
    Close #mHandle
    Call GetNames
    
End Sub

Private Sub Command1_Click()
    FavoriteFileName = txtPath.Text
    Dim mHandle
    mHandle = FreeFile
    Open FavoriteFileName For Binary As #mHandle
    Close #mHandle
    
    Dim FileNum
    FileNum = FreeFile
    Open FavoriteFileName For Input As FileNum
    Text1.Text = Input(LOF(FileNum), FileNum)
    Close FileNum
End Sub

Private Sub GetSettings_Click()
    
    Dim strFilename As String
    FileDialog.CancelError = True
    'On Error GoTo errhandler
    FileDialog.Filter = _
    "INI Files|*.ini"
    On Error Resume Next
    FileDialog.CancelError = True
    FileDialog.ShowOpen
    
    txtPath.Text = FileDialog.FileName
    On Error GoTo errhandler
    Call GetNames
    txtInformation.Text = "Click the names in this listbox"
    Command1_Click
    Exit Sub
errhandler:
    If Err.Number = 53 Then
        MsgBox "File Not Found"
    End If
End Sub

Private Sub GetNames()
    ListNames.Clear
    iniFile = txtPath.Text
    File = FreeFile
    
    Open iniFile For Input As File
    Do Until EOF(File)
        Line Input #File, buffer
        If Left(buffer, 1) = "[" Then
            iniName = Mid(buffer, 2, Len(buffer) - 2)
            ListNames.AddItem (iniName)
        End If
        Loop
        Close File
    End Sub

Private Sub GetKeys()
    
    iniFile = txtPath.Text
    File = FreeFile
    ListKey1.Clear
    ListKey2.Clear
    
    finished = False
    iniName = "[" & ListNames.Text & "]"
    
    Open iniFile For Input As File
    Do
        
        Line Input #File, buffer
        
        If Left(buffer, Len(ListNames.Text) + 2) = iniName Then
            
            Do
                If Not EOF(File) Then
                    Line Input #File, keys
                    If Left(keys, 1) = "[" Or keys = "" Then 'check if another key starts or if this one ends or if its EOF
                    finished = True
                Else
                    
                    pos = InStr(keys, "=") - 1
                    
                    ListKey1.AddItem (Left(keys, pos))
                    ListKey2.AddItem (Mid(keys, pos + 2))
                End If
            Else
                finished = True
            End If
        Loop Until finished = True
        
    End If
    
Loop Until finished = True
Close File
End Sub

Private Sub ListNames_Click()
    Call GetKeys
End Sub

Private Sub ListKey1_Click()
    ListKey2.ListIndex = ListKey1.ListIndex
    lblEntries = ListKey1.Text & "=" & ListKey2.Text
    lblFindNext.Caption = "Find " & lblEntries & " if there are any!"
    cmdFindNext.Enabled = True
    cmdFindFirst_Click
End Sub
Private Sub cmdFindFirst_Click()
    If lblEntries.Caption <> "" Then
        If InStr(Text1, lblEntries.Caption) <> 0 Then
            Text1.SelStart = InStr(Text1, lblEntries.Caption) - 1
            Text1.SelLength = Len(lblEntries.Caption)
        Else
            MsgBox "Not found"
        End If
    End If
    
End Sub

Private Sub cmdfindNext_Click()
    ListKey2.ListIndex = ListKey1.ListIndex
    lblEntries = ListKey1.Text & "=" & ListKey2.Text
    If lblEntries.Caption <> "" Then
        Text1.SelStart = Text1.SelStart + 2
        If InStr(Text1.SelStart, Text1, lblEntries.Caption) <> 0 Then
            Text1.SelStart = InStr(Text1.SelStart, Text1, lblEntries.Caption) - 1
            Text1.SelLength = Len(lblEntries.Caption)
        Else
            MsgBox "Not Found"
        End If
    End If
End Sub
