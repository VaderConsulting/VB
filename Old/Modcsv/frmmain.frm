VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Modify Domain Admin .csv files"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   3840
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtbCSV 
      Height          =   3255
      Left            =   7680
      TabIndex        =   16
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":0442
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtDrivePath 
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   13
      Text            =   "na"
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox txtDrivePath 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   12
      Text            =   "%siteserver%\General"
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtDrivePath 
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   11
      Text            =   "%siteserver%\Transfer"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox txtDrivePath 
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   10
      Text            =   "%siteserver%\Applications"
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtDrivePath 
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Text            =   "%siteserver%\Corporate"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtDrivePath 
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Text            =   "na"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblFilename 
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Drive Paths for Traveller users"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblDrivePath 
      Caption         =   "Q:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblDrivePath 
      Caption         =   "O:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblDrivePath 
      Caption         =   "X:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblDrivePath 
      Caption         =   "N:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblDrivePath 
      Caption         =   "T:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblDrivePath 
      Caption         =   "P:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblPath 
      Alignment       =   2  'Center
      Caption         =   "Path to Domain Admin .csv files"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdModify_Click()
  Dim D As String
  Dim p As String
  Dim CompleteCSV As String
  Dim pos1 As Long
  Dim pos2 As Long
  Dim Filename As String
  Dim TextToReplace As String
  Dim TextToInsert As String
  Dim DriveText As String
  Dim PrinterText As String
  Dim lp As Integer
  
  TextToInsert = ""
  TextToInsert = TextToInsert & "Traveller,"
  TextToInsert = TextToInsert & txtDrivePath(0) & ","
  TextToInsert = TextToInsert & txtDrivePath(1) & ","
  TextToInsert = TextToInsert & txtDrivePath(2) & ","
  TextToInsert = TextToInsert & txtDrivePath(3) & ","
  TextToInsert = TextToInsert & txtDrivePath(4) & ","
  TextToInsert = TextToInsert & txtDrivePath(5) & ","
  
  CompleteCSV = ""
  
  ' Remove \ from right of txtPath
  If Right(txtPath, 1) = "\" Then txtPath = Left(txtPath, Len(txtPath) - 1)
  p = txtPath & "\"
  
  If Trim(p) = "" Then Exit Sub
  Screen.MousePointer = vbHourglass
  ' Create dir for new files if necessary
  If Dir(p & "New", vbDirectory) = "" Then
    MkDir (p) & "New"
  End If
  
  p = p & "*.csv"
  D = Dir(p)
  Do Until D = ""
    Filename = Left(p, Len(p) - 6) & "\" & D
    TextToReplace = ""
    PrinterText = ""
    If InStr(1, UCase(Filename), "X") > 0 Then
      rtbCSV.Filename = Filename
      lblFilename = Filename
      frmMain.Refresh
      pos1 = rtbCSV.Find("Traveller")
      pos2 = rtbCSV.Find(vbCrLf, pos1)
      If pos1 > 0 And pos2 > 0 Then
        TextToReplace = Mid(rtbCSV.Text, pos1 + 1, (pos2 - pos1))
        ' Now extract the printer paths.. these are to be added to the new drive paths
        For lp = Len(TextToReplace) To 1 Step -1
          If Mid(TextToReplace, lp, 1) = "," Then i = i + 1
          If i = 3 Then Exit For
        Next lp
        If i <> 1 Then
          DriveText = Left(TextToReplace, lp)
          PrinterText = Mid(TextToReplace, lp + 1, 255)
        End If
        i = 0
        rtbCSV.Text = Replace(rtbCSV.Text, DriveText, TextToInsert)
        
        ' Save file back
        Open txtPath & "\New\" & D For Output As #1
          Print #1, rtbCSV.Text
        Close 1
      End If
    End If
    D = Dir
  Loop
  rtbCSV.Text = ""
  Screen.MousePointer = vbDefault
  MsgBox "New files have been saved to " & txtPath & "\New"
  lblFilename = ""
End Sub

Private Sub Form_Load()
  txtPath = App.Path
  
End Sub

Private Sub txtDrivePath_Change(Index As Integer)
  If txtDrivePath(Index) = "" Then txtDrivePath(Index) = "NA"
End Sub
