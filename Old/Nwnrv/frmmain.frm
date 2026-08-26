VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NWN Resource Viewer"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNWNDir 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "C:\NeverwinterNights\NWN"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblNWNDir 
      Caption         =   "NWN Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

