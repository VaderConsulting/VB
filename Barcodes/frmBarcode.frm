VERSION 5.00
Begin VB.Form frmBarcode 
   BorderStyle     =   0  'None
   Caption         =   "Barcode"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   LinkTopic       =   "Form2"
   ScaleHeight     =   1590
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBarcode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1212
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   120
      Width           =   3852
   End
End
Attribute VB_Name = "frmBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

