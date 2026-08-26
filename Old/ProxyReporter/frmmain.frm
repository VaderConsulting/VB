VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proxy Reporter"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLogfile 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "C:\Documents\VB\ProxyReporter\WS010820.log"
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
