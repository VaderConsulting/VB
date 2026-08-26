VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CSV Maker"
   ClientHeight    =   12585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12585
   ScaleWidth      =   16350
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtSiteLoc 
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   4320
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid flxCSV 
      Height          =   7935
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   13996
      _Version        =   393216
      Rows            =   45
      Cols            =   11
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open in Notepad"
      Height          =   615
      Left            =   3240
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3975
   End
   Begin VB.FileListBox filSource 
      Height          =   2235
      Left            =   3240
      Pattern         =   "*.bat"
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.DirListBox dirSource 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtSites 
      Height          =   285
      Left            =   6120
      TabIndex        =   2
      Text            =   "Sites"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtPDC 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "\\CBDXAAA"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "SiteLoc:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Destination Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblDir 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label lblScript 
      Caption         =   "Admin$\system32\repl\export\scripts"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()
    Shell "C:\Winnt\System32\Notepad.exe " & filSource.Path & "\" & filSource.List(filSource.ListIndex), vbNormalFocus
End Sub

Private Sub cmdRefresh_Click()
    dirSource.Path = txtSource.Text
End Sub

Private Sub dirSource_Change()
    filSource.Path = dirSource.Path
End Sub

Private Sub Form_Load()
    lblDir = txtPDC & "\" & lblScript & "\" & txtSites
    txtSource = txtPDC & "\" & lblScript
    cmdRefresh_Click
    flxCSV.Row = 0
    flxCSV.Col = 1
    flxCSV.ColWidth(1) = 2000
    flxCSV.Text = "Group"
    flxCSV.Col = 2
    flxCSV.Text = "P:"
    flxCSV.Col = 3
    flxCSV.Text = "T:"
    flxCSV.Col = 4
    flxCSV.Text = "N:"
    flxCSV.Col = 5
    flxCSV.Text = "X:"
    flxCSV.Col = 6
    flxCSV.Text = "O:"
    flxCSV.Col = 7
    flxCSV.Text = "Q:"
    flxCSV.Col = 8
    flxCSV.Text = "Printer 1"
    flxCSV.Col = 9
    flxCSV.Text = "Printer 2"
    flxCSV.Col = 10
    flxCSV.Text = "Printer 3"
    For cell = 1 To 10
        flxCSV.ColWidth(cell) = 1500
    Next cell
End Sub

Private Sub txtPDC_Change()
    lblDir = txtPDC & "\" & lblScript & "\" & txtSites
    txtSource = txtPDC & "\" & lblScript
End Sub

Private Sub txtSiteLoc_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtSites_Change()
    lblDir = txtPDC & "\" & lblScript & "\" & txtSites
    txtSource = txtPDC & "\" & lblScript
End Sub

