VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmFlow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6030
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Down"
      Height          =   615
      Left            =   4800
      TabIndex        =   40
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Up"
      Height          =   615
      Left            =   4800
      TabIndex        =   39
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "Start"
      Height          =   855
      Left            =   4800
      Picture         =   "frmFlow.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   120
      Width           =   1095
   End
   Begin ComctlLib.Slider sldStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   9720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   327682
      Max             =   7
   End
   Begin ComctlLib.ProgressBar pbrStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   9480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   7
   End
   Begin VB.CommandButton cmdCont 
      Caption         =   "Continue"
      Height          =   855
      Left            =   4800
      Picture         =   "frmFlow.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblStep 
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   600
      Y1              =   7560
      Y2              =   6360
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Refund"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   38
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   9
      Left            =   120
      Top             =   7560
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   600
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label lblN 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Index           =   7
      Left            =   1080
      TabIndex        =   37
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   135
      Index           =   7
      Left            =   1920
      TabIndex        =   36
      Top             =   6840
      Width           =   255
   End
   Begin VB.Line lneFlow 
      Index           =   52
      X1              =   1800
      X2              =   2280
      Y1              =   6840
      Y2              =   6360
   End
   Begin VB.Line lneFlow 
      Index           =   51
      X1              =   1800
      X2              =   1320
      Y1              =   6840
      Y2              =   6360
   End
   Begin VB.Line lneFlow 
      Index           =   50
      X1              =   2280
      X2              =   1800
      Y1              =   6360
      Y2              =   5880
   End
   Begin VB.Line lneFlow 
      Index           =   49
      X1              =   1320
      X2              =   1800
      Y1              =   6360
      Y2              =   5880
   End
   Begin VB.Line lneFlow 
      Index           =   48
      X1              =   1800
      X2              =   1800
      Y1              =   7320
      Y2              =   6840
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Replace?"
      Height          =   465
      Index           =   16
      Left            =   1320
      TabIndex        =   35
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Caption         =   "Claim Progress"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   9120
      Width           =   5895
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buy New"
      Height          =   255
      Index           =   15
      Left            =   3600
      TabIndex        =   29
      Top             =   7200
      Width           =   975
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   8
      Left            =   3600
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Repair"
      Height          =   255
      Index           =   14
      Left            =   4800
      TabIndex        =   28
      Top             =   6240
      Width           =   975
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   7
      Left            =   4800
      Top             =   6120
      Width           =   975
   End
   Begin VB.Line lneFlow 
      Index           =   47
      X1              =   4800
      X2              =   4560
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label lblN 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Index           =   6
      Left            =   4200
      TabIndex        =   27
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   135
      Index           =   6
      Left            =   4560
      TabIndex        =   26
      Top             =   6480
      Width           =   255
   End
   Begin VB.Line lneFlow 
      Index           =   46
      X1              =   4080
      X2              =   4080
      Y1              =   7080
      Y2              =   6840
   End
   Begin VB.Line lneFlow 
      Index           =   45
      X1              =   4080
      X2              =   4560
      Y1              =   6840
      Y2              =   6360
   End
   Begin VB.Line lneFlow 
      Index           =   44
      X1              =   4080
      X2              =   3600
      Y1              =   6840
      Y2              =   6360
   End
   Begin VB.Line lneFlow 
      Index           =   43
      X1              =   4560
      X2              =   4080
      Y1              =   6360
      Y2              =   5880
   End
   Begin VB.Line lneFlow 
      Index           =   42
      X1              =   3600
      X2              =   4080
      Y1              =   6360
      Y2              =   5880
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Repair?"
      Height          =   255
      Index           =   13
      Left            =   3720
      TabIndex        =   25
      Top             =   6240
      Width           =   735
   End
   Begin VB.Line lneFlow 
      Index           =   41
      X1              =   5040
      X2              =   4560
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Image imgFlow 
      Height          =   480
      Index           =   1
      Left            =   5040
      Picture         =   "frmFlow.frx":0884
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image imgFlow 
      Height          =   480
      Index           =   0
      Left            =   5040
      Picture         =   "frmFlow.frx":0CC6
      Top             =   4920
      Width           =   480
   End
   Begin VB.Line lneFlow 
      Index           =   40
      X1              =   3000
      X2              =   3000
      Y1              =   7560
      Y2              =   5160
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "In Stock?"
      Height          =   255
      Index           =   12
      Left            =   1440
      TabIndex        =   24
      Top             =   7680
      Width           =   735
   End
   Begin VB.Line lneFlow 
      Index           =   39
      X1              =   1320
      X2              =   1800
      Y1              =   7800
      Y2              =   7320
   End
   Begin VB.Line lneFlow 
      Index           =   38
      X1              =   2280
      X2              =   1800
      Y1              =   7800
      Y2              =   7320
   End
   Begin VB.Line lneFlow 
      Index           =   37
      X1              =   1800
      X2              =   1320
      Y1              =   8280
      Y2              =   7800
   End
   Begin VB.Line lneFlow 
      Index           =   36
      X1              =   1800
      X2              =   2280
      Y1              =   8280
      Y2              =   7800
   End
   Begin VB.Line lneFlow 
      Index           =   35
      X1              =   1800
      X2              =   1800
      Y1              =   8520
      Y2              =   8280
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   135
      Index           =   5
      Left            =   1920
      TabIndex        =   23
      Top             =   8280
      Width           =   255
   End
   Begin VB.Label lblN 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Index           =   5
      Left            =   2280
      TabIndex        =   22
      Top             =   7920
      Width           =   255
   End
   Begin VB.Line lneFlow 
      Index           =   34
      X1              =   2520
      X2              =   2280
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   6
      Left            =   2520
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Order New"
      Height          =   255
      Index           =   11
      Left            =   2520
      TabIndex        =   21
      Top             =   7680
      Width           =   975
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   5
      Left            =   1320
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Replace"
      Height          =   255
      Index           =   10
      Left            =   1320
      TabIndex        =   20
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "< 7 Days?"
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   19
      Top             =   5040
      Width           =   735
   End
   Begin VB.Line lneFlow 
      Index           =   33
      X1              =   1320
      X2              =   1800
      Y1              =   5160
      Y2              =   4680
   End
   Begin VB.Line lneFlow 
      Index           =   32
      X1              =   2280
      X2              =   1800
      Y1              =   5160
      Y2              =   4680
   End
   Begin VB.Line lneFlow 
      Index           =   31
      X1              =   1800
      X2              =   1320
      Y1              =   5640
      Y2              =   5160
   End
   Begin VB.Line lneFlow 
      Index           =   30
      X1              =   1800
      X2              =   2280
      Y1              =   5640
      Y2              =   5160
   End
   Begin VB.Line lneFlow 
      Index           =   29
      X1              =   1800
      X2              =   1800
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   135
      Index           =   4
      Left            =   1920
      TabIndex        =   18
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label lblN 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   2280
      TabIndex        =   17
      Top             =   5280
      Width           =   255
   End
   Begin VB.Line lneFlow 
      Index           =   28
      X1              =   3000
      X2              =   2280
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line lneFlow 
      Index           =   27
      X1              =   4080
      X2              =   4080
      Y1              =   4680
      Y2              =   4200
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   4
      Left            =   3600
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Claim invalid"
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   16
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Replace?"
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   15
      Top             =   5040
      Width           =   735
   End
   Begin VB.Line lneFlow 
      Index           =   26
      X1              =   3600
      X2              =   4080
      Y1              =   5160
      Y2              =   4680
   End
   Begin VB.Line lneFlow 
      Index           =   25
      X1              =   4560
      X2              =   4080
      Y1              =   5160
      Y2              =   4680
   End
   Begin VB.Line lneFlow 
      Index           =   24
      X1              =   4080
      X2              =   3600
      Y1              =   5640
      Y2              =   5160
   End
   Begin VB.Line lneFlow 
      Index           =   23
      X1              =   4080
      X2              =   4560
      Y1              =   5640
      Y2              =   5160
   End
   Begin VB.Line lneFlow 
      Index           =   22
      X1              =   4080
      X2              =   4080
      Y1              =   5880
      Y2              =   5640
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   135
      Index           =   3
      Left            =   4200
      TabIndex        =   14
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label lblN 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Index           =   3
      Left            =   4560
      TabIndex        =   13
      Top             =   5280
      Width           =   255
   End
   Begin VB.Line lneFlow 
      Index           =   21
      X1              =   5040
      X2              =   4560
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warranty?"
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   12
      Top             =   3840
      Width           =   735
   End
   Begin VB.Line lneFlow 
      Index           =   20
      X1              =   1320
      X2              =   1800
      Y1              =   3960
      Y2              =   3480
   End
   Begin VB.Line lneFlow 
      Index           =   19
      X1              =   2280
      X2              =   1800
      Y1              =   3960
      Y2              =   3480
   End
   Begin VB.Line lneFlow 
      Index           =   18
      X1              =   1800
      X2              =   1320
      Y1              =   4440
      Y2              =   3960
   End
   Begin VB.Line lneFlow 
      Index           =   17
      X1              =   1800
      X2              =   2280
      Y1              =   4440
      Y2              =   3960
   End
   Begin VB.Line lneFlow 
      Index           =   16
      X1              =   1800
      X2              =   1800
      Y1              =   4680
      Y2              =   4440
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   135
      Index           =   2
      Left            =   1920
      TabIndex        =   11
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblN 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Index           =   2
      Left            =   2280
      TabIndex        =   10
      Top             =   4080
      Width           =   255
   End
   Begin VB.Line lneFlow 
      Index           =   15
      X1              =   3600
      X2              =   2280
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   3
      Left            =   3600
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Not ours"
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.Line lneFlow 
      Index           =   14
      X1              =   4080
      X2              =   4080
      Y1              =   2760
      Y2              =   2040
   End
   Begin VB.Line lneFlow 
      Index           =   13
      X1              =   4080
      X2              =   3480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblN 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   3480
      TabIndex        =   8
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   135
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice?"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Line lneFlow 
      Index           =   12
      X1              =   2520
      X2              =   3000
      Y1              =   2040
      Y2              =   1560
   End
   Begin VB.Line lneFlow 
      Index           =   11
      X1              =   3480
      X2              =   3000
      Y1              =   2040
      Y2              =   1560
   End
   Begin VB.Line lneFlow 
      Index           =   10
      X1              =   3000
      X2              =   2520
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line lneFlow 
      Index           =   9
      X1              =   3000
      X2              =   3480
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line lneFlow 
      Index           =   8
      X1              =   3000
      X2              =   3000
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line lneFlow 
      Index           =   7
      X1              =   2520
      X2              =   2280
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblN 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   135
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manual"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.Line lneFlow 
      Index           =   6
      X1              =   1800
      X2              =   1800
      Y1              =   3480
      Y2              =   2520
   End
   Begin VB.Line lneFlow 
      Index           =   5
      X1              =   1800
      X2              =   2280
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line lneFlow 
      Index           =   4
      X1              =   1800
      X2              =   1320
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line lneFlow 
      Index           =   3
      X1              =   2280
      X2              =   1800
      Y1              =   2040
      Y2              =   1560
   End
   Begin VB.Line lneFlow 
      Index           =   2
      X1              =   1320
      X2              =   1800
      Y1              =   2040
      Y2              =   1560
   End
   Begin VB.Line lneFlow 
      Index           =   1
      X1              =   1800
      X2              =   1800
      Y1              =   1560
      Y2              =   1320
   End
   Begin VB.Line lneFlow 
      Index           =   0
      X1              =   1800
      X2              =   1800
      Y1              =   840
      Y2              =   600
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   2
      Left            =   2520
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Listed?"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   1
      Left            =   1320
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scan/Enter"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblFlow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape shpFlow 
      Height          =   495
      Index           =   0
      Left            =   1320
      Shape           =   2  'Oval
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape shpHilite 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCont_Click()
    gStepNo = 3
    StepChange
    frmData.lblStatus = "CONTINUING CLAIM PROGRESS"
    rsStock.MoveFirst
    a = Int(frmData.lstRecord.List(0))
    rsStock.Move a
    If rsStock.Fields(5) <> "" Then lblDateOut = CDate(rsStock.Fields(5) & "")
    WarrantyCheck
    cmdOperation.Visible = True
    cmdCont.Visible = False
End Sub

Private Sub cmdOperation_Click()
    If cmdOperation.Caption = "Start" Then
        For a = 0 To 8
            frmData.lblField(a) = ""
        Next a
        frmData.imgDate.Visible = False
        gStepNo = 1
        StepChange
        cmdOperation.Caption = "Stop"
        cmdOperation.Picture = frmPic.imgStop.Picture
        frmData.txtSerial.Enabled = True
        frmData.txtDb.Enabled = False
        frmData.lblStatus = "Waiting for input"
        frmData.txtSerial.SetFocus
        frmData.lstMatch.Clear
        Exit Sub
    End If
    If cmdOperation.Caption = "Stop" Then
        gStepNo = 0
        StepChange
        cmdOperation.Caption = "Start"
        cmdOperation.Picture = frmPic.imgStart.Picture
        frmData.lblStatus = "Waiting for Start"
        frmData.txtSerial.Enabled = False
        frmData.txtSerial = ""
        frmData.pbrSearch.Value = 0
    End If
End Sub

Private Sub Command1_Click()
    gStepNo = gStepNo + 1
    StepChange
End Sub

Private Sub Command2_Click()
    gStepNo = gStepNo - 1
    StepChange
End Sub

Private Sub Form_Load()
    shpHilite.Left = 1080
    shpHilite.Top = 0
    pbrStatus.Max = gStepMax
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLeft = X
    lblTop = Y
End Sub

Private Sub lblStep_Change()
    StepChange
    Select Case gStepNo
        Case 5
        choice = MsgBox("The client has the choice of a refund or a replacement item.  Replace the item ?", vbYesNo + vbQuestion, "Question")
        If choice = vbYes Then
            gStepNo = 6
            StepChange
        Else
            gStepNo = 19
            StepChange
        End If
        Case 6
            frmData.lblStatus = "Checking if replacement item is in stock"
            CheckStock
        Case 9
            frmOrders.Show
            frmOrders.txtField(2) = frmData.lblField(2)
            frmOrders.txtField(0) = frmData.lblField(1)
            frmOrders.txtField(1) = frmData.lblField(4)
            frmOrders.txtField(3) = Format(Date, "Long date")
        Case 11
            gStepNo = 12
            StepChange
        Case 12
            ReplaceReply = MsgBox("Does the client want the item replaced ?", vbQuestion + vbYesNo, "Question")
            If ReplaceReply = vbNo Then
                gStepNo = 16
                StepChange
            End If
            If ReplaceReply = vbYes Then
                gStepNo = 13
                StepChange
            End If
        Case 13
            RepairReply = MsgBox("Does the client want the item repaired ?", vbQuestion + vbYesNo, "Question")
            If RepairReply = vbYes Then
                gStepNo = 17
                StepChange
            End If
            If RepairReply = vbNo Then
                gStepNo = 14
                StepChange
            End If
        Case 19
            frmData.lblStatus = "Refund the full price of the item to the client."
    End Select
End Sub


