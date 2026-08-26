VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "SAX Workbench"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Filter Options"
      Height          =   2775
      Left            =   240
      TabIndex        =   32
      Top             =   960
      Width           =   7635
      Begin VB.ComboBox ComboIndentChar 
         Height          =   315
         Left            =   6000
         TabIndex        =   48
         Text            =   "Space"
         Top             =   1860
         Width           =   1155
      End
      Begin VB.TextBox Indentation 
         Height          =   285
         Left            =   6000
         TabIndex        =   47
         Text            =   "2"
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Frame Frame6 
         Caption         =   "Formatting"
         Height          =   1635
         Left            =   2640
         TabIndex        =   41
         Top             =   960
         Width           =   2235
         Begin VB.OptionButton OptionCompressed 
            Caption         =   "Compressed"
            Height          =   315
            Left            =   120
            TabIndex        =   44
            Top             =   900
            Width           =   1815
         End
         Begin VB.OptionButton OptionIndented 
            Caption         =   "Indented"
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   540
            Width           =   1875
         End
         Begin VB.OptionButton OptionNone 
            Caption         =   "None"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   300
            Value           =   -1  'True
            Width           =   1635
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Name Folding"
         Height          =   1635
         Left            =   300
         TabIndex        =   37
         Top             =   960
         Width           =   2235
         Begin VB.OptionButton OptionNoFolding 
            Caption         =   "None"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   300
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptionProper 
            Caption         =   "Proper casing"
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton OptionLower 
            Caption         =   "Lower case"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   900
            Width           =   1515
         End
         Begin VB.OptionButton OptionUpper 
            Caption         =   "Upper case"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1875
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   315
         Left            =   6900
         TabIndex        =   36
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox OutFile 
         Height          =   315
         Left            =   1320
         TabIndex        =   34
         Top             =   300
         Width           =   5475
      End
      Begin VB.CheckBox AttrsToElements 
         Caption         =   "Attributes to elements"
         Height          =   255
         Left            =   5040
         TabIndex        =   33
         Top             =   1020
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Indent Char"
         Height          =   255
         Left            =   4980
         TabIndex        =   46
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Indentation"
         Height          =   255
         Left            =   4980
         TabIndex        =   45
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Output to file:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About..."
      Height          =   375
      Left            =   6720
      TabIndex        =   31
      Top             =   3900
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Counters"
      Height          =   2715
      Left            =   240
      TabIndex        =   20
      Top             =   960
      Width           =   3975
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   4
         Left            =   1920
         TabIndex        =   27
         ToolTipText     =   "Total number of characters in each and every element and attribute name"
         Top             =   1920
         Width           =   1875
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   26
         ToolTipText     =   "Total number of characters in element and attribute text content"
         Top             =   1500
         Width           =   1875
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   2
         Left            =   1920
         TabIndex        =   25
         ToolTipText     =   "Total number of text nodes"
         Top             =   1080
         Width           =   1875
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   24
         ToolTipText     =   "Total number of attributes"
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   22
         ToolTipText     =   "Total number of elements"
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Name Chars"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   30
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Text Chars"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   29
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Text Nodes"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   28
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Attributes"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Elements"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Left            =   6300
      TabIndex        =   19
      Top             =   120
      Width           =   435
   End
   Begin VB.Frame Frame1 
      Caption         =   "Overall"
      Height          =   2715
      Left            =   4440
      TabIndex        =   16
      Top             =   960
      Width           =   3315
      Begin VB.CommandButton FullReport 
         Caption         =   "Generate Report"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   2100
         Width           =   2655
      End
      Begin VB.Label Percent 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   56.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1455
         Left            =   360
         TabIndex        =   18
         Top             =   420
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Counters"
      Height          =   2715
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   3855
      Begin VB.TextBox TextResult 
         Height          =   315
         Index           =   4
         Left            =   1740
         TabIndex        =   10
         ToolTipText     =   "The subset of tests that failed during output comparison"
         Top             =   2220
         Width           =   1935
      End
      Begin VB.TextBox TextResult 
         Height          =   315
         Index           =   3
         Left            =   1740
         TabIndex        =   9
         ToolTipText     =   "The subset of passed tests reported an error in the input.  You need to verify that they reported the right thing."
         Top             =   1740
         Width           =   1935
      End
      Begin VB.TextBox TextResult 
         Height          =   315
         Index           =   2
         Left            =   1740
         TabIndex        =   8
         ToolTipText     =   "The total number of tests that failed."
         Top             =   1260
         Width           =   1935
      End
      Begin VB.TextBox TextResult 
         Height          =   315
         Index           =   1
         Left            =   1740
         TabIndex        =   7
         ToolTipText     =   "The number of tests passed"
         Top             =   780
         Width           =   1935
      End
      Begin VB.TextBox TextResult 
         Height          =   315
         Index           =   0
         Left            =   1740
         TabIndex        =   6
         ToolTipText     =   "Total number of tests run"
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Output Failed"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Negative Passes"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Failed"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Passed"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Total"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5741
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Confromance"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Statistics"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Filter"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "http://localhost/oasis/xmlconf/xmlconf.xml"
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Status 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oasis As OasisTest

Private Sub Command1_Click()
    If (Text1.text = "") Then
        Status.Caption = "Please enter URL."
        Text1.SetFocus
        GoTo Error
    End If
    If (Command1.Caption = "Cancel") Then
        oasis.Cancel = True
        GoTo Error
    End If
    Command1.Caption = "Cancel"
    DoEvents
    
    Select Case TabStrip1.SelectedItem.Index
    Case 1
        Set oasis = New OasisTest
        Call oasis.run(Text1.text)
    Case 2
        Dim stats As Statistics
        Set stats = New Statistics
        Call stats.run(Text1.text)
    Case 3
        Dim f As Filter
        Set f = New Filter
        f.NamesToUpper = OptionUpper.value
        f.NamesToLower = OptionLower.value
        f.NamesToProper = OptionProper.value
        f.AttrsToElements = AttrsToElements.value
        If (Form1.OptionIndented) Then f.Format = True
        If (Form1.OptionCompressed) Then f.Compress = True
        f.Indentation = Form1.Indentation.text
        If (ComboIndentChar.text = "Space") Then
            f.IndentChar = " "
        Else
            f.IndentChar = Chr(9)
        End If
        Call f.run(Text1.text, OutFile.text)
    End Select
Error:
    Command1.Caption = "Run"
End Sub

Private Sub Command2_Click()
    CommonDialog1.Filter = "XML (*.xml)|*.xml|All (*.*)|*.*"
    CommonDialog1.ShowOpen
    Text1.text = CommonDialog1.FileName
End Sub

Private Sub Command3_Click()
    frmAbout.Show
End Sub

Private Sub Command4_Click()
    CommonDialog1.Filter = "XML (*.xml)|*.xml|All (*.*)|*.*"
    CommonDialog1.ShowOpen
    OutFile.text = CommonDialog1.FileName
End Sub

Private Sub Form_Load()
    Set oasis = New OasisTest
    Frame3.Visible = False
    Frame1.Caption = "Overall Conformance"
    Frame4.Visible = False
    ComboIndentChar.AddItem ("Space")
    ComboIndentChar.AddItem ("Tab")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oasis.Quit = True
    oasis.Cancel = True
End Sub

Private Sub FullReport_Click()
    oasis.GenerateReport
End Sub

Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.Index
    Case 1
        Frame3.Visible = False
        FullReport.Visible = True
        Frame1.Caption = "Overall Conformance"
        Frame1.Visible = True
        Frame2.Visible = True
        Frame4.Visible = False
        Percent.Caption = ""
        Form1.Percent.BackColor = &H8000000F
    Case 2
        Frame3.Visible = True
        FullReport.Visible = False
        Frame1.Caption = "Tagginess Ratio"
        Frame1.Visible = True
        Frame2.Visible = False
        Frame4.Visible = False
        Percent.Caption = ""
        Percent.Caption = ""
        Form1.Percent.BackColor = &H8000000F
    Case 3
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = True
    End Select
End Sub

