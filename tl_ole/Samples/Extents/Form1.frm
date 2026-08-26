VERSION 5.00
Object = "*\AIOleObject Extent.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   3255
   ClientTop       =   3375
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   2310
   Begin prjIOleObjectExtent.UC_Extents UC_Extents1 
      Height          =   960
      Left            =   420
      TabIndex        =   0
      Top             =   870
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1693
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Resize the control in design-time"
      Height          =   585
      Left            =   248
      TabIndex        =   1
      Top             =   210
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' Adding items to UserControl context menu
'
' Sample Form
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 02/09/2000
'
'*********************************************************************************************
Option Explicit

