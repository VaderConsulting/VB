VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose date"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin MSACAL.Calendar calDates 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _Version        =   524288
      _ExtentX        =   7858
      _ExtentY        =   5318
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2000
      Month           =   7
      Day             =   27
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblFrom 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    If lblFrom = "0" Then
        '        frmMain.txtFrom = calDates.Value
        frmMain.txtFrom = Format(Year(calDates.Value), "0000") & Format(Month(calDates.Value), "00") & Format(Day(calDates.Value), "00")
    Else
        '        frmMain.txtTo = calDates.Value
        frmMain.txtTo = Format(Year(calDates.Value), "0000") & Format(Month(calDates.Value), "00") & Format(Day(calDates.Value), "00")
    End If
    Unload Me
End Sub
