VERSION 5.00
Begin VB.Form frmBarcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New ID"
   ClientHeight    =   3360
   ClientLeft      =   2700
   ClientTop       =   1380
   ClientWidth     =   2295
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   2295
   Begin VB.Frame fmeType 
      Caption         =   "Type"
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2055
      Begin VB.OptionButton optType 
         Caption         =   "Location Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optType 
         Caption         =   "Unescorted Visitor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optType 
         Caption         =   "Escorted Visitor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optType 
         Caption         =   "Standard"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdLookup 
      Caption         =   "Lookup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Large"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Small"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Done"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblPersonalNumber 
      Alignment       =   2  'Center
      Caption         =   "Personal Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
    frmPrint.lblNameRank = ""
    frmPrint.lblBillet = ""
    Text1 = ""
    Text1.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload frmPrint
    Unload Me
End Sub

Private Sub cmdLookup_Click()
    frmPrint.lblNameRank = ""
    frmPrint.lblBillet = ""
    Do Until a = 232
        If gPersonalNumber(a) = Val(Text1.Text) Then
            frmPrint.lblNameRank = gRank(a) & " " & gName(a)
            frmPrint.lblBillet = gBilletNumber(a) & " " & gBilletText(a)
            Exit Sub
        End If
        a = a + 1
        DoEvents
    Loop
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
    frmPrint.PrintForm
    Printer.EndDoc
End Sub

Private Sub Form_Activate()
    optSize(0) = 1
    Text1.SetFocus
    Left = (Screen.Width - Width - frmPrint.Width) / 2
    Top = (Screen.Height - Height) / 2
    frmPrint.Left = Left + Width
    frmPrint.Top = Top
End Sub

Private Sub Form_Load()
    frmPrint.Show
    frmPrint.lblType = "Gangway Pass"
End Sub

Private Sub optSize_Click(Index As Integer)
    frmPrint.picBarcode.ScaleMode = 3
    
    Select Case Index
        Case 0
            frmPrint.picBarcode.Height = frmPrint.picBarcode.Height * (1.4 * 40 / frmPrint.picBarcode.ScaleHeight)
            frmPrint.picBarcode.FontSize = 8
        Case 1
            frmPrint.picBarcode.Height = frmPrint.picBarcode.Height * (2.4 * 40 / frmPrint.picBarcode.ScaleHeight)
            frmPrint.picBarcode.FontSize = 10
        Case 2
            frmPrint.picBarcode.Height = frmPrint.picBarcode.Height * (3 * 40 / frmPrint.picBarcode.ScaleHeight)
            frmPrint.picBarcode.FontSize = 14
    End Select
    
    Call Text1_Change

End Sub

Private Sub optType_Click(Index As Integer)
    Select Case Index
        Case 0
            frmPrint.lblType = "Gangway Pass"
            cmdLookup.Enabled = True
        Case 1
            frmPrint.lblType = "Escorted Visitor"
            cmdLookup.Enabled = False
        Case 2
            frmPrint.lblType = "Unescorted Visitor"
            cmdLookup.Enabled = False
        Case 3
            frmPrint.lblType = "Location"
            cmdLookup.Enabled = False
    End Select
    frmPrint.lblBillet = ""
    frmPrint.lblNameRank = ""
    Text1 = ""
    Text1.SetFocus
End Sub

Private Sub Text1_Change()
    Call ModuleBarcode.DrawBarcode(Text1, frmPrint.picBarcode)
    
    MinWidth = 2 * Text1.Left + Text1.Width
    pw = 2 * frmPrint.picBarcode.Left + frmPrint.picBarcode.Width
    fw = MinWidth
    If pw > fw Then fw = pw
    'frmBarcode.Width = fw
    frmPrint.picBarcode.Left = ((frmPrint.Width - frmPrint.picBarcode.Width) / 2) + 120

End Sub


