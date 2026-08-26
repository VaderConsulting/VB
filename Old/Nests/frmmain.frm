VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Robin"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   10845
   Begin VB.Frame fmeSearch 
      Caption         =   "Search Area"
      Height          =   975
      Left            =   2040
      TabIndex        =   12
      Top             =   480
      Width           =   2055
      Begin VB.OptionButton optNest 
         Caption         =   "Nest number"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optAnywhere 
         Caption         =   "Anywhere"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.ListBox lstNest 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   10695
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse >>>"
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtNest 
      BackColor       =   &H00C0FFFF&
      DataSource      =   "datNest"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   10320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Label lblSearch 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Search String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblGrade 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grade"
      Height          =   255
      Left            =   8760
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblPlate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thickness"
      Height          =   255
      Left            =   7920
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Description"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblDwg 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Drawing"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label lblNest 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nest number"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblMaterial 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Material"
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblDrawing 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Drawing"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
    frmMain.Height = 7665
End Sub

Private Sub cmdSearch_Click()
    Static Error_Presented
    For varSearch = 1 To lstNest.ListCount - 1
        If optNest.Value = True Then
            If InStr(1, Left$(lstNest.List(varSearch), 8), txtNest, vbTextCompare) <> 0 Then
                lstNest.ListIndex = varSearch
                cmdBrowse_Click
                txtNest.SelStart = 0
                txtNest.SelLength = Len(txtNest)
                Exit Sub
            Else
            End If
        End If
        If optAnywhere.Value = True Then
            If InStr(1, lstNest.List(varSearch), txtNest, vbTextCompare) <> 0 Then
                lstNest.ListIndex = varSearch
                cmdBrowse_Click
                txtNest.SelStart = 0
                txtNest.SelLength = Len(txtNest)
                Exit Sub
            Else
            End If
        End If
    Next varSearch
    If Error_Presented = 0 Then
        MsgBox "Nest Number not found", vbOKOnly + vbInformation, "Message"
        Error_Presented = 1
        txtNest.SelStart = 0
        txtNest.SelLength = Len(txtNest)
    End If
End Sub

Public Sub Form_Load()
    frmSplash.Show
    Dim gdBNest As Database
    Dim grsDrawings As Recordset
    Dim fldNest As Field
    Dim fldName As Field
    Dim fldDescription As Field
    Dim fldPlate As Field
    Dim fltGrade As Fields
    Dim a(5) As String
    a(1) = String(10, " ")
    a(2) = String(35, " ")
    a(3) = String(20, " ")
    a(4) = String(7, " ")
    a(5) = String(15, " ")
    Height = 2235
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - 7665) / 2
    gdb = "C:\DATA\ACCESS\Nest.MDB"
    Set gdBNest = OpenDatabase(gdb)
    Set grsDrawings = gdBNest.OpenRecordset("Drawings")
    grsDrawings.MoveFirst
    Do Until grsDrawings.EOF
        LSet a(1) = grsDrawings.Fields(0) & ""
        LSet a(2) = grsDrawings.Fields(1) & ""
        LSet a(3) = grsDrawings.Fields(2) & ""
        LSet a(4) = grsDrawings.Fields(3) & ""
        LSet a(5) = grsDrawings.Fields(4) & ""
        varAll = a(1) & a(2) & a(3) & a(4) & a(5)
        lstNest.AddItem varAll
        grsDrawings.MoveNext
    Loop
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub txtNest_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtNest_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmdSearch_Click
    End If
End Sub
