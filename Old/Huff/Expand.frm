VERSION 4.00
Begin VB.Form frmExpand 
   AutoRedraw      =   -1  'True
   Caption         =   "Huffman Data Expansion Demo"
   ClientHeight    =   4635
   ClientLeft      =   1470
   ClientTop       =   1800
   ClientWidth     =   7260
   Height          =   5325
   Left            =   1410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   7260
   Top             =   1170
   Width           =   7380
   Begin VB.PictureBox picInvisible 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   4875
      TabIndex        =   15
      Top             =   4140
      Width           =   4935
   End
   Begin VB.PictureBox picVisible 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   4875
      TabIndex        =   14
      Top             =   3420
      Width           =   4935
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1320
      ScaleHeight     =   255
      ScaleWidth      =   3915
      TabIndex        =   12
      Top             =   2940
      Width           =   3945
   End
   Begin VB.Frame fraOutputFile 
      Caption         =   "Expanded File"
      Height          =   1155
      Left            =   120
      TabIndex        =   7
      Top             =   1410
      Width           =   5325
      Begin VB.TextBox txtOutLength 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtOutputPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   4035
      End
      Begin VB.Label lblOutLength 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Left            =   450
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblOutputPath 
         AutoSize        =   -1  'True
         Caption         =   "Output Path"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.Frame fraInputFile 
      Caption         =   "Compressed File"
      Height          =   1155
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5325
      Begin VB.TextBox txtInFileLength 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   1245
      End
      Begin VB.TextBox txtInputPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   4035
      End
      Begin VB.Label lblLength 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblInputPath 
         AutoSize        =   -1  'True
         Caption         =   "Input Path"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   525
      Left            =   5820
      TabIndex        =   1
      Top             =   3420
      Width           =   1245
   End
   Begin VB.CommandButton cmdExpand 
      Caption         =   "Expand File"
      Height          =   525
      Left            =   5820
      TabIndex        =   0
      Top             =   2670
      Width           =   1245
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      Caption         =   "Progress Messages"
      Height          =   195
      Left            =   2550
      TabIndex        =   13
      Top             =   2700
      Width           =   1380
   End
   Begin MSComDlg.CommonDialog Dlg1 
      Left            =   6450
      Top             =   240
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Compressed File"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLn1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Open &Expanded File"
         Index           =   1
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuLn2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmExpand"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExpand_Click()
    Dim hFileIn As Integer
    Dim hFileOut As Integer
    Dim sPathIn As String
    Dim sPathOut As String
    Dim iRootNode As Integer
    Dim TreeCounts As Integer
    Dim ProgressStep As Single
    
    
    hFileIn = FreeFile
    sPathIn = txtInputPath.Text
    Open sPathIn For Binary Access Read As #hFileIn
    If hFileIn = 0 Then
        MsgBox "Error opening input file", 48
        Exit Sub
    End If
    txtInFileLength = LOF(hFileIn)
    
    hFileOut = FreeFile
    sPathOut = txtOutputPath.Text
    Open sPathOut For Binary Access Write As #hFileOut
    If hFileOut = 0 Then
        MsgBox "Error opening output file", 48
        Exit Sub
    End If
    
    picProgress.Cls
    picProgress.Print "Reading Counts..."
    TreeCounts = InputCounts(hFileIn)
    If TreeCounts = 0 Then
        picProgress.Cls
        picProgress.Print "Fatal Error..Quiting..."
        Exit Sub
    Else
        ProgressStep = ((txtInFileLength * 1.6) - TreeCounts) / 100
        ProgressStep = ProgressStep + 1
    End If
    picProgress.Cls
    picProgress.Print "Building Decoding Tree..."
    iRootNode = BuildTree
    picProgress.Cls
    picProgress.Print "Expanding Data..."
    ExpandData hFileIn, hFileOut, iRootNode, CLng(ProgressStep)
    txtOutLength = LOF(hFileOut)
    Close
    picProgress.Cls
    picProgress.Print "Expansion Complete..."
End Sub

Private Sub Form_Load()
    Height = 4810
    picVisible.Visible = False
    picInvisible.Width = picVisible.Width
    picInvisible.Height = picVisible.Height
End Sub

Private Sub mnuExit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub mnuFileOpen_Click(Index As Integer)
    dlg1.Filter = "Huff Files (*.??_)|*.??_"
    dlg1.DialogTitle = "Open Compressed File"
    dlg1.ShowOpen
    txtInputPath.Text = dlg1.filename
End Sub

Private Sub mnuFileSave_Click(Index As Integer)
    Dim StartSearch As Integer
    Dim Found As Integer
    Dim LastSlash As Integer
    
    If txtInputPath.Text <> "" Then
        Do
            StartSearch = Found + 1
            Found = InStr(StartSearch, txtInputPath, "\")
                If Found <> 0 Then LastSlash = Found
        Loop Until Found = 0
        dlg1.InitDir = Mid(txtInputPath.Text, 1, Found)
        dlg1.filename = ""
    End If
    dlg1.Filter = "All Files (*.*)|*.*"
    dlg1.DialogTitle = "Save Expanded File As"
    dlg1.ShowSave
    txtOutputPath.Text = dlg1.filename
End Sub
