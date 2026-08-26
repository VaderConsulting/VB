VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Song Search"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   10185
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNew 
      Caption         =   "New..."
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import DeluxeCD Database"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8640
      TabIndex        =   1
      Top             =   7920
      Width           =   1455
   End
   Begin ComctlLib.TreeView tvwMusic 
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   12726
      _Version        =   327682
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   6
      ImageList       =   "imlMusic"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList imlMusic 
      Left            =   0
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImport_Click()
    Dim counter As Integer
    Dim db As Database
    Dim tb As Recordset
    
    Set db = OpenDatabase("c:\win95\deluxecd.mdb")
    Set tb = db.OpenRecordset("Titles")
    tb.MoveFirst
    tb.MoveLast
    tb.MoveFirst
    On Error Resume Next
    For counter = 0 To tb.RecordCount - 1
        If tb(1) <> "" Then
            tvwMusic.Nodes.Add , , "P-" & tb(1), tb(1)
            MkDir (gRootDir & tb(1))
            tvwMusic.Nodes.Add "P-" & tb(1), tvwChild, "C-" & tb(2), tb(2)
            MkDir (gRootDir & tb(1) & "\" & tb(2))
        End If
        tb.MoveNext
    Next counter
    tb.Close
    db.Close
End Sub

Private Sub Form_Load()
    Dim nodename As String
    Dim root As String
    Dim rtncode As Integer
    Dim counter As Integer
    'populate root of treeview control
    root = gRootDir
    rtncode = PopulateNodes(root)
    ' populate children of root
    For counter = 1 To tvwMusic.Nodes.Count
        root = gRootDir & tvwMusic.Nodes(counter).Text & "\"
        rtncode = PopulateNodes(root, tvwMusic.Nodes(counter).Text)
    Next counter
    If Dir("c:\win95\deluxecd.mdb") <> "" Then
        cmdImport.Enabled = True
    End If
End Sub

Function PopulateNodes(root As String, Optional parent As String) As Integer
    Dim nodename As String
    nodename = Dir(root, vbDirectory)
    Do Until nodename = ""
        If nodename <> "." And nodename <> ".." Then
            If (GetAttr(root) And vbDirectory) = vbDirectory Then
                If parent <> "" Then ' child node
                    tvwMusic.Nodes.Add "P-" & parent, tvwChild, "C-" & nodename, nodename
                Else               ' root node
                    tvwMusic.Nodes.Add , , "P-" & nodename, nodename
                End If
            End If
        End If
        nodename = Dir
    Loop
End Function

Private Sub tvwMusic_NodeClick(ByVal Node As ComctlLib.Node)
    If Left(Node.Key, 2) = "P-" Then Exit Sub ' exit sub if not a child (album)
    
End Sub
