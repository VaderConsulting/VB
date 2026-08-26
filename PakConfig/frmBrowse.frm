VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ListBox lstItems 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblItemType 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   7200
      X2              =   0
      Y1              =   3620
      Y2              =   3620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   7200
      X2              =   0
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
