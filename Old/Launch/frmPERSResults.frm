VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmPERSResults 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PERS Results"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart chtData 
      Height          =   1935
      Left            =   7200
      OleObjectBlob   =   "frmPERSResults.frx":0000
      TabIndex        =   71
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton cmdHints 
      Caption         =   "Hints"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Display helpful tips on usage"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   7560
      TabIndex        =   20
      ToolTipText     =   "Print to the default printer"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   11
      Left            =   6360
      TabIndex        =   70
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   10
      Left            =   6360
      TabIndex        =   69
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   6360
      TabIndex        =   68
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   6360
      TabIndex        =   67
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   6360
      TabIndex        =   66
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   65
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   64
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   63
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   62
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   61
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   60
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   59
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   11
      Left            =   5400
      TabIndex        =   58
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   10
      Left            =   5400
      TabIndex        =   57
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   56
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   55
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   54
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   53
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   52
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   51
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   50
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   49
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   48
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblWeighting 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   47
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   11
      Left            =   4800
      TabIndex        =   46
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   10
      Left            =   4800
      TabIndex        =   45
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   4800
      TabIndex        =   44
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   4800
      TabIndex        =   43
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   4800
      TabIndex        =   42
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   41
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   40
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   39
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   38
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   37
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   36
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   35
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   34
      Top             =   8520
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   33
      Top             =   8160
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   32
      Top             =   7800
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   31
      Top             =   7440
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   30
      Top             =   7080
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   29
      Top             =   6720
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   28
      Top             =   6360
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   27
      Top             =   6000
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   5640
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   9735
   End
   Begin VB.Label lblJustify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   4560
      Width           =   9735
   End
   Begin VB.Line Line1 
      X1              =   6000
      X2              =   7200
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7680
      TabIndex        =   21
      ToolTipText     =   "Double click to change date"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblTotal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblCriteria 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Criteria"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label lblWeightings 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Weighting"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblScoreTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblFinalTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Final"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Double click to change Surname"
      Top             =   120
      Width           =   7335
   End
   Begin VB.Image imgCrest 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   8040
      Picture         =   "frmPERSResults.frx":2356
      Stretch         =   -1  'True
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmPERSResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDone_Click()
    frmPERS.Show
    If frmPERS.chkHelp.Value = 1 Then frmScores.Show
    Unload Me
End Sub

Private Sub cmdHints_Click()
    msg = "Double click date to modify PERS date" & vbCrLf
    msg = msg & "Double click surname to modify spelling"
    MsgBox msg, vbInformation + vbOKOnly, "Hints"
End Sub

Private Sub cmdPrint_Click()
    If frmPERS.chkHelp = 1 Then ' only if they have asked for help compiling scores
        Rem Dim justify(11)
        For a = 0 To 11 ' twelve times becuase 12 criteria
            If lblCriteria(a) <> "" Then ' only if the criteria applies
                If Right$(lblScore(a), 1) = "*" Then ' only if they have a * score
                    txt = "Please enter justification for this score." & vbCrLf
                    txt = txt & gCriteria(a + 1, Val(Left$(lblScore(a), 1)))
                    lblJustify(a) = InputBox(txt, "Justification", gCriteria(a + 1, Val(Left$(lblScore(a), 1))))
                End If
            End If
        Next a
    End If
    cmdPrint.Visible = False
    cmdDone.Visible = False
    cmdHints.Visible = False
    frmPERSResults.Height = 9300
    PrintForm
    Printer.EndDoc
    frmPERSResults.Height = 4860
    cmdPrint.Visible = True
    cmdDone.Visible = True
    cmdHints.Visible = True
    MsgBox "Print Complete", vbInformation + vbOKOnly, "PERS Result"
End Sub

Private Sub Form_Activate()
    lblName = "PERS Result for " & gRank & " " & gSurname
    lbldate = Date
    chtData.Visible = frmPERS.chtData.Visible
    chtData.ColumnCount = frmPERS.chtData.ColumnCount
    With chtData.Plot.Axis(VtChAxisIdY).ValueScale
        .Auto = False
        .MajorDivision = 0
        .MinorDivision = 0
        .Maximum = 9
        .Minimum = 0
    End With
    
    'Height = 6810
    'Height = 10650
    Height = 4860
    cmdPrint.SetFocus
    
    ' get data from previous form
    For a = 0 To 11
        lblCriteria(a) = frmPERS.lblCriteria(a)
        lblScore(a) = frmPERS.txtScore(a)
        If Val(lblScore(a)) < 3 Or Val(lblScore(a)) > 7 Then
            lblScore(a) = lblScore(a) & "*"
        End If
        lblWeighting(a) = frmPERS.txtWeight(a)
        lblFinal(a) = frmPERS.txtFinal(a)
        lblTotal = frmPERS.lblTotal
        If lblCriteria(a) <> "" Then
            lblScore(a).Visible = True
            lblWeighting(a).Visible = True
            lblFinal(a).Visible = True
        Else
            lblScore(a).Visible = False
            lblWeighting(a).Visible = False
            lblFinal(a).Visible = False
        End If
    Next a
    chtData.ColumnCount = frmPERS.chtData.ColumnCount
    For a = 1 To chtData.ColumnCount
        chtData.Column = a
        chtData.Data = Left$(lblScore(a - 1), 1)
    Next a
End Sub

Private Sub lbldate_DblClick()
    result = MsgBox("Is this date correct ?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Confirm date")
    If result = vbNo Then
        vardate = InputBox("Enter correct date for PERS", "Date correction", lbldate.Caption)
        lbldate = vardate
        lbldate.Refresh
    End If
End Sub

Private Sub lblName_DblClick()
    result = MsgBox("Is this spelling correct ? [" & gSurname & "]", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Confirm spelling")
    If result = vbNo Then
        gSurname = InputBox("Enter correct spelling for Surname", "Spelling correction", gSurname)
        lblName = "PERS Result for " & gRank & " " & gSurname
        lblName.Refresh
    End If
End Sub
