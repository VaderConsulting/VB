VERSION 4.00
Begin VB.Form frmHuff 
   AutoRedraw      =   -1  'True
   Caption         =   "Huffman Data Compression Demo"
   ClientHeight    =   4635
   ClientLeft      =   450
   ClientTop       =   1935
   ClientWidth     =   8880
   Height          =   5325
   Left            =   390
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   8880
   Top             =   1305
   Width           =   9000
   Begin VB.PictureBox picInvisible 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2070
      ScaleHeight     =   435
      ScaleWidth      =   4875
      TabIndex        =   24
      Top             =   4110
      Width           =   4935
   End
   Begin VB.PictureBox picVisible 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2070
      ScaleHeight     =   435
      ScaleWidth      =   4875
      TabIndex        =   23
      Top             =   3420
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtPercent 
      Height          =   285
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4140
      ScaleHeight     =   255
      ScaleWidth      =   2805
      TabIndex        =   19
      Top             =   2880
      Width           =   2835
   End
   Begin VB.CheckBox chkModel 
      Caption         =   "Print Model"
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   525
      Left            =   7470
      TabIndex        =   17
      Top             =   3420
      Width           =   1245
   End
   Begin VB.CheckBox chkNodes 
      Caption         =   "Print Nodes"
      Height          =   255
      Left            =   90
      TabIndex        =   16
      Top             =   2850
      Width           =   1455
   End
   Begin VB.CheckBox chkCounts 
      Caption         =   "Print Counts"
      Height          =   255
      Left            =   90
      TabIndex        =   15
      Top             =   2580
      Width           =   1455
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Compressed File"
      Height          =   1125
      Left            =   90
      TabIndex        =   9
      Top             =   1380
      Width           =   5175
      Begin VB.TextBox txtEncodeLength 
         Height          =   285
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox txtOutFile 
         Height          =   285
         Left            =   870
         TabIndex        =   10
         Top             =   270
         Width           =   4095
      End
      Begin VB.Label lblEncodeLength 
         Caption         =   "Length"
         Height          =   315
         Left            =   210
         TabIndex        =   14
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lblOutFile 
         AutoSize        =   -1  'True
         Caption         =   "Out Path"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   330
         Width           =   630
      End
   End
   Begin VB.Frame fraInput 
      Caption         =   "Input File"
      Height          =   1125
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   5175
      Begin VB.TextBox txtFileLength 
         Height          =   285
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox txtInFile 
         Height          =   285
         Left            =   870
         TabIndex        =   6
         Top             =   270
         Width           =   4095
      End
      Begin VB.Label lblInFile 
         AutoSize        =   -1  'True
         Caption         =   " In Path"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblFileLength 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   750
         Width           =   495
      End
   End
   Begin VB.ListBox lstWeights 
      Height          =   2010
      Left            =   7470
      TabIndex        =   3
      Top             =   420
      Width           =   1245
   End
   Begin VB.ListBox lstCounts 
      Height          =   2010
      Left            =   5940
      TabIndex        =   1
      Top             =   420
      Width           =   1245
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Compress"
      Height          =   525
      Left            =   7470
      TabIndex        =   0
      Top             =   2730
      Width           =   1245
   End
   Begin VB.Label lblPercent 
      AutoSize        =   -1  'True
      Caption         =   "% Compressed"
      Height          =   195
      Left            =   2250
      TabIndex        =   21
      Top             =   2670
      Width           =   1035
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "Messages"
      Height          =   195
      Left            =   5250
      TabIndex        =   20
      Top             =   2670
      Width           =   720
   End
   Begin MSComDlg.CommonDialog Dlg1 
      Left            =   5370
      Top             =   1080
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Label lblWeights 
      AutoSize        =   -1  'True
      Caption         =   "Node Weights"
      Height          =   195
      Left            =   7590
      TabIndex        =   4
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label lblCounts 
      AutoSize        =   -1  'True
      Caption         =   "ASCII Character Counts"
      Height          =   195
      Left            =   5730
      TabIndex        =   2
      Top             =   120
      Width           =   1680
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuInputFileOpen 
         Caption         =   "&Open Input File"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenCompressedFile 
         Caption         =   "Open &Compressed File"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmHuff"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompress_Click()
    Dim hFileIn As Integer
    Dim hFileOut As Integer
    Dim sPathIn As String
    Dim sPathOut As String
    Dim RootNode As Integer
    Dim Percent As Single
    Dim ProgressStep As Long
    
    hFileIn = FreeFile
    sPathIn = txtInFile.Text
    Open sPathIn For Binary Access Read As #hFileIn
    If hFileIn = 0 Then
        MsgBox "Error opening input file", 48
        txtInFile.SetFocus
        Exit Sub
    End If
    txtfilelength.Text = LOF(hFileIn)
    ProgressStep = LOF(hFileIn) / 100
    ProgressStep = ProgressStep + 1
    If txtfilelength.Text = 0 Then
        picProgress.Cls
        picProgress.Print "Input file has zero length..exiting..."
        Exit Sub
    End If
    
    hFileOut = FreeFile
    sPathOut = txtOutFile
    Open sPathOut For Binary Access Write As #hFileOut
    If hFileOut = 0 Then
        MsgBox "Error opening output file", 48
        txtOutFile.SetFocus
        Exit Sub
    End If
    
    picProgress.Cls
    picProgress.Print "Counting Characters..."
    ' Count occurrences of ASCII characters
    CharacterCounts hFileIn, ProgressStep
    Close #hFileIn
    picProgress.Cls
    picProgress.Print "Scaling Counts..."
    ScaleCounts                 ' Scale counts to fit integer
    picProgress.Cls
    picProgress.Print "Building Tree..."
    RootNode = BuildTree       ' Build Huffman tree
    If chkNodes.Value = Checked Then
        picProgress.Cls
        picProgress.Print "Printing Nodes..."
        PrintNodes RootNode
    End If
    ' Build Symbol table for Huffman tree
    ' First 0 is CodeSoFar variable
    ' Second 0 is Bits variable
    picProgress.Cls
    picProgress.Print "Building Codes..."
    ConvertTreeToCode 0, 0, RootNode
    If chkModel.Value = Checked Then
        picProgress.Cls
        picProgress.Print "Printing Model..."
        PrintModel RootNode
    End If
    picProgress.Cls
    picProgress.Print "Outputting Counts..."
    OutputCounts hFileOut
    hFileIn = FreeFile
    Open sPathIn For Binary Access Read As #hFileIn
    If hFileIn = 0 Then
        MsgBox "Error opening input file, second read", 48
        txtInFile.SetFocus
        Exit Sub
    End If
    picProgress.Cls
    picProgress.Print "Outputting Huffman Codes..."
    CompressFile hFileIn, hFileOut, ProgressStep
    Percent = (txtEncodeLength / txtfilelength) * 100
    txtPercent = 100 - Format(Percent, "#0.00")
    picProgress.Cls
    picProgress.Print "Compression Done..."
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Height = 4800
    picInvisible.Width = picVisible.Width
    picInvisible.Height = picVisible.Height
End Sub

Private Sub lstcounts_Click()
    lstWeights.ListIndex = lstcounts.ListIndex
End Sub

Private Sub lstWeights_Click()
    lstcounts.ListIndex = lstWeights.ListIndex
End Sub

Private Sub mnuInputFileOpen_Click(Index As Integer)
    Dim Path As String
    Dim Length As Integer
    
    Dlg1.Filter = "All Files(*.*)|*.*"
    Dlg1.DialogTitle = "Open File to be Compressed"
    Dlg1.ShowOpen
    txtInFile.Text = Dlg1.filename
    ' Create the default output file name.
    Length = Len(Dlg1.filename)
    Path = Mid(Dlg1.filename, 1, Length - 1) & "_"
    txtOutFile.Text = Path
End Sub

Private Sub mnuOpenCompressedFile_Click(Index As Integer)
    Dlg1.Filter = "All Files(*.*)|*.*"
    Dlg1.DialogTitle = "Open File for Compressed Output"
    Dlg1.filename = txtOutFile.Text
    Dlg1.ShowOpen
    txtOutFile.Text = Dlg1.filename
End Sub
