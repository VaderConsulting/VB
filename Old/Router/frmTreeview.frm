VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmTreeview 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Treeview - existing signals"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLog 
      Caption         =   "View log"
      Height          =   375
      Left            =   10320
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtbPreview 
      Height          =   5175
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9128
      _Version        =   327681
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmTreeview.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   11520
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.FileListBox filFiles 
      Height          =   5160
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9128
      _Version        =   327682
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmTreeview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Rem close the form
    Unload Me
End Sub

Private Sub cmdLog_Click()
    Rem show the router logfile
    rtbPreview.filename = gLog
End Sub

Private Sub filFiles_Click()
    Rem show the selected file in the preview window
    rtbPreview.filename = filFiles.Path & "\" & filFiles.List(filFiles.ListIndex)
End Sub

Private Sub Form_Load()
    Rem set up the treeview control
    Dim nodX As Node
    Set nodX = TreeView1.Nodes.Add(, , "R", "Signals root")
    nodX.Expanded = True
    For a = frmSend.chkPersonnel().LBound To frmSend.chkPersonnel().UBound
        If frmSend.chkPersonnel(a).Caption <> "" Then
            Set nodX = TreeView1.Nodes.Add("R", tvwChild, , frmSend.chkPersonnel(a).Caption)
        End If
    Next a
    Set nodX = TreeView1.Nodes.Add(, , "P", "Archive")
    nodX.Expanded = True
    Rem set the pattern to a nonexistant one so that there are no files displayed
    Rem at start up
    filFiles.Pattern = "*.QWE"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As Node)
    Rem show the files in the directory of the selected node
    filFiles.Pattern = "*." & frmSend.txtExtension.Text
    If Node.Index > 1 And Node.Index - 1 < frmSend.chkPersonnel().UBound Then
        filFiles.Path = gDir(Node.Index - 2)
    End If
    If Node.Text = "Archive" Then
        filFiles.Path = gArchive
        filFiles.Refresh
    End If
    If Node.Text = "Signals root" Then
        filFiles.Pattern = "*.QWE"
    End If
    rtbPreview.filename = ""
End Sub
