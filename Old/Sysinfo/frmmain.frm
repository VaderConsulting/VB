VERSION 5.00
Object = "{1B773E42-2509-11CF-942F-008029004347}#3.3#0"; "sysmon.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin SystemMonitorCtl.SystemMonitor monServers 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   11775
      _Version        =   196611
      _ExtentX        =   20770
      _ExtentY        =   13150
      DisplayType     =   1
      ReportValueType =   0
      MaximumScale    =   100
      MinimumScale    =   0
      ShowLegend      =   -1  'True
      ShowToolbar     =   -1  'True
      ShowScaleLabels =   -1  'True
      ShowHorizontalGrid=   0   'False
      ShowVerticalGrid=   0   'False
      ShowValueBar    =   -1  'True
      ManualUpdate    =   0   'False
      Highlight       =   0   'False
      ReadOnly        =   0   'False
      MonitorDuplicateInstances=   -1  'True
      UpdateInterval  =   1
      BackColorCtl    =   -2147483633
      ForeColor       =   -1
      BackColor       =   -1
      GridColor       =   8421504
      TimeBarColor    =   255
      Appearance      =   -1
      BorderStyle     =   0
      GraphTitle      =   ""
      YAxisLabel      =   ""
      LogFileName     =   ""
      AmbientFont     =   -1  'True
      LegendColumnWidths=   $"frmMain.frx":0000
      LegendSortDirection=   6619253
      LegendSortColumn=   0
      CounterCount    =   1
      MaximumSamples  =   100
      SampleCount     =   0
      Counter00001.Path=   "\\server\Processor(_Total)\% Processor Time"
      Counter00001.Color=   255
      Counter00001.Width=   1
      Counter00001.LineStyle=   0
      Counter00001.ScaleFactor=   2147483647
      Selected        =   "\\server\Processor(_Total)\% Processor Time"
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1320
      Top             =   8160
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   7560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Label1 = monServers.Counters(1).Value
End Sub
