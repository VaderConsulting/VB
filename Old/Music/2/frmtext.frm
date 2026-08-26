VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.Form frmText 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6960
   ClientLeft      =   1080
   ClientTop       =   1590
   ClientWidth     =   9600
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6960
   ScaleWidth      =   9600
   Begin VB.ListBox lstFiles 
      Height          =   840
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   6030
      Width           =   8295
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Save"
      Height          =   315
      Left            =   8370
      TabIndex        =   4
      Top             =   5940
      Width           =   1185
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   810
      TabIndex        =   2
      Top             =   5580
      Width           =   7485
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   8370
      TabIndex        =   1
      Top             =   5580
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   8370
      TabIndex        =   0
      Top             =   6570
      Width           =   1185
   End
   Begin RichTextLib.RichTextBox txtText 
      Height          =   5505
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   9710
      _Version        =   327681
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FRMTEXT.frx":0000
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
   Begin MSComDlg.CommonDialog cdBrowse 
      Left            =   8280
      Top             =   6210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Label lblFilename 
      Alignment       =   1  'Right Justify
      Caption         =   "Filename"
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   5580
      Width           =   735
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim datmarty As Database
    Dim tabSongs As Recordset
    Dim tabBands As Recordset

Private Sub cmdBrowse_Click()
    cdBrowse.filename = ""
    cdBrowse.Filter = "Text files (*.TXT)|*.TXT|All files (*.*)|*.*"
    cdBrowse.ShowOpen
    lstFiles.AddItem cdBrowse.filename
End Sub

Private Sub cmdClose_Click()
    frmStart.Show
    frmText.Hide
End Sub

Private Sub cmdSelect_Click()
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    tabSongs.MoveFirst
    Do Until tabSongs.EOF
        If tabSongs(1) = frmStart.txtSong Then
            tabSongs.Edit
            tabSongs(13) = 1
            frmStart.chkFormat(2) = 1
            tabSongs(19) = txtText
            tabSongs.Update
            cmdClose_Click
        End If
        tabSongs.MoveNext
    Loop
End Sub

Private Sub Form_Activate()
    frmText.Caption = "Text - " & frmStart.txtSong
    subListUpdate
End Sub

Private Sub Form_Load()
    Width = gVGAWidth
    Height = gVGAHeight
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub
Private Sub subListUpdate()
    lstFiles.Clear
    Textname = Dir(CurDir & "\*.TXT")
    Rem If Textname <> "" Then lstFiles.AddItem CurDir & "\" & Textname
    If Textname <> "" Then lstFiles.AddItem Textname
    Do Until Picturename = ""
        Textname = Dir
        Rem lstFiles.AddItem CurDir & "\" & Textname
        lstFiles.AddItem Textname
    Loop
    On Error Resume Next
    lstFiles.RemoveItem 0
End Sub

Private Sub lstFiles_Click()
    txtFilename = lstFiles.List(lstFiles.ListIndex)
End Sub

Private Sub txtFilename_Change()
    Screen.MousePointer = vbHourglass
    txtText.LoadFile txtFilename
    Screen.MousePointer = vbNormal
    txtText.Refresh
End Sub
