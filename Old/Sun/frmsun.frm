VERSION 5.00
Begin VB.Form frmSun 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4440
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4440
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txtTAY 
      Height          =   285
      Left            =   1680
      TabIndex        =   22
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtTAX 
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtEA 
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtE 
      Height          =   285
      Left            =   1680
      TabIndex        =   18
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtMA 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtJulian 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtSunrise 
      Height          =   285
      Left            =   4920
      TabIndex        =   12
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSunrise 
      Caption         =   "Sunrise"
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtDay 
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtMonth 
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLongitude 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "144.9"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtLatitude 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "-37.85"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   6360
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblEA 
      Caption         =   "Eccentric Anomaly"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblE 
      Caption         =   "Eccentricity"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblMA 
      Caption         =   "Mean Anomaly"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblJulian 
      Caption         =   "Julian day"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblDay 
      Caption         =   "Day"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblMonth 
      Caption         =   "Month"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLongitude 
      Caption         =   "Longitude"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblLatitude 
      Caption         =   "Latitude"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmSun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
    End
End Sub

Public Sub cmdSunrise_Click()
    Dim MeanAnomaly As Double
    Dim Eccentricity As Double
    Dim JulianDay As Double
    Dim EccentricAnomaly As Double
    Dim TrueAnomalyXCoord As Double
    Dim TrueAnomalyYCoord As Double
    
    vDay = Val(txtDay)
    vMonth = Val(txtMonth)
    vYear = Val(txtYear)
    
    JulianDay = CInt(367 * vYear - 7 * (vYear + Int((vMonth + 9) / 12)) / 4 + Int(275 * vMonth / 9) + vDay - 730530)
    txtJulian = JulianDay
    MeanAnomaly = 356.047 + 0.9856002585 * JulianDay
    txtMA = MeanAnomaly
    Eccentricity = 0.016709 - (1.151 * 10 ^ -9) * JulianDay
    txtE = Eccentricity
    Text1 = MeanAnomaly
    EccentricAnomaly = MeanAnomaly + Eccentricity * 180 / PI * SinD(MeanAnomaly) * (1 + Eccentricity * CosD(MeanAnomaly))
    Text2 = MeanAnomaly
    txtEA = EccentricAnomaly
    TrueAnomalyXCoord = CosD(EccentricAnomaly) - Eccentricity
    txtTAX = TrueAnomalyXCoord
    TrueAnomalyYCoord = SinD(EccentricAnomaly) * Sqr(1 - Eccentricity ^ 2)
    Text1 = MeanAnomaly
    txtTAY = TrueAnomalyYCoord
End Sub

Public Sub Form_Load()
    txtYear = DatePart("yyyy", Now)
    txtMonth = DatePart("m", Now)
    txtDay = DatePart("d", Now)
End Sub
