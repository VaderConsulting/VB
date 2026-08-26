VERSION 5.00
Begin VB.Form frmInstruc 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructions"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "frmInstruc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FOR BEST RESULTS IN RUNNING THE GAME PLEASE RUN IT FROM YOUR HARD DISK!!!"
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblInstructoins 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblnfo2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmInstruc.frx":0442
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1935
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Image imgPlay 
      BorderStyle     =   1  'Fixed Single
      Height          =   3060
      Left            =   3240
      Picture         =   "frmInstruc.frx":05C0
      Top             =   4200
      Width           =   3060
   End
   Begin VB.Image imgClose 
      BorderStyle     =   1  'Fixed Single
      Height          =   3060
      Left            =   120
      Picture         =   "frmInstruc.frx":1284
      Top             =   4200
      Width           =   3060
   End
End
Attribute VB_Name = "frmInstruc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************'
'Filename :       frmInstruc.frm                             '
'Author:          Coenraad Steenkamp                         '
'Created:         09/05/2000                                 '
'Operating System Windows 9x & NT ,2000 ,Me                  '
'                                                            '
'Description:                                                '
'This from shows instructions on how to play the game and    '
'some tips.                                                  '
'                                                            '
'************************************************************'
Option Explicit
Private Sub imgClose_Click()
   'Declare Local Variables
   Dim strAnswer As String
   'Ask's the user if he wants to Exit
   strAnswer = MsgBox("Are you sure?", vbYesNo, "Confirmation...")
   If strAnswer = vbYes Then
      'Unload the instruction form
      Unload frmInstruc
      'Show the Intro form
      frmIntro.Show
   End If
End Sub
Private Sub imgPlay_Click()
   'Declare Local Variables
   Dim strAnswer As String
   'Ask's the user if he wants to Exit
   strAnswer = MsgBox("Are you sure?", vbYesNo, "Confirmation...")
   If strAnswer = vbYes Then
      'Unload the instruction form
      Unload frmInstruc
      'Show the Intro form
      frmPlay.Show
   End If
End Sub
