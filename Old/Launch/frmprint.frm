VERSION 5.00
Begin VB.Form frmPrint 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   9000
      TabIndex        =   25
      Top             =   7440
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   2400
      Top             =   6480
      Width           =   4815
   End
   Begin VB.Label lblPolicy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2520
      TabIndex        =   27
      Top             =   6480
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   4920
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   3975
   End
   Begin VB.Label lblExplain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Explanation of Scores:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   9735
   End
   Begin VB.Line linExplain 
      X1              =   120
      X2              =   9840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   11
      Top             =   840
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   3
      Top             =   4680
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   2
      Top             =   5160
      Width           =   8500
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   1
      Top             =   5640
      Width           =   8505
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   12
      Left            =   480
      TabIndex        =   0
      Top             =   6120
      Width           =   8505
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    frmPERS.Show
    If frmPERS.chkHelp.Value = 1 Then frmScores.Show
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    frmPrint.BackColor = &H8000000E
    cmdPrint.Visible = False
    cmdDone.Visible = False
    PrintForm
    Printer.EndDoc
    frmPrint.BackColor = &H8000000B
    cmdPrint.Visible = True
    cmdDone.Visible = True
    MsgBox "Print Complete", vbInformation + vbOKOnly, "Explanation of scores"
End Sub

Private Sub Form_Activate()
    If gRank = "SMN" Or gRank = "AB" Then
        Criteria = 6
        file = gDrive2 & "\bulletin\lan\SMN-AB.txt"
    End If
    
    If gRank = "LS" Then
        Criteria = 8
        file = gDrive2 & "\bulletin\lan\LS.txt"
    End If
    
    If gRank = "PO" Or gRank = "CPO" Then
        Criteria = 11
        file = gDrive2 & "\bulletin\lan\PO-CPO.txt"
    End If
    
    Open file For Input As #1
    For a = 0 To Criteria
        lblCriteria(a) = frmPERS.lblCriteria(a)
        For b = 1 To 9
            Line Input #1, txtexplain
            score = Val(frmPERS.txtScore(a))
            If score = b Then lblExplain(a + 1) = txtexplain
        Next b
    Next a
    Close 1
    lblPolicy = vbCrLf
    lblPolicy = lblPolicy & "Current RAN policy does not allow us to focus on PERS scores," & vbCrLf
    lblPolicy = lblPolicy & "but rather on the individual's performance." & vbCrLf
    lblPolicy = lblPolicy & "Because of this, the author encourages the reportee" & vbCrLf
    lblPolicy = lblPolicy & "discuss the meaning of their PERS with Divisional staff."
End Sub
