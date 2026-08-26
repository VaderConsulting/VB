VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.Form frmView 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6960
   ClientLeft      =   1140
   ClientTop       =   1590
   ClientWidth     =   9600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6960
   ScaleWidth      =   9600
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   315
      Left            =   7110
      TabIndex        =   1
      Top             =   6570
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
      Height          =   6495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   11456
      _Version        =   327681
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FRMVIEW.frx":0000
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
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    frmStart.Show
    frmView.Hide
End Sub

Private Sub cmdPrint_Click()
    Printer.Print txtText
    Printer.NewPage
    Printer.EndDoc
End Sub


Private Sub Form_Activate()
    Width = gVGAWidth
    Height = gVGAHeight
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
    Caption = frmStart.txtSong
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    tabSongs.MoveFirst
    Do Until tabSongs.EOF
        If tabSongs(1) = frmStart.txtSong Then
            txtText = tabSongs(19)
            Exit Do
        End If
        tabSongs.MoveNext
    Loop
End Sub

