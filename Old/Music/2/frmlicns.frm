VERSION 5.00
Begin VB.Form frmLicense 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "License Agreement"
   ClientHeight    =   4080
   ClientLeft      =   3795
   ClientTop       =   2925
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMLICNS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4080
   ScaleWidth      =   6570
   Begin VB.CommandButton cmdDisagree 
      Caption         =   "I Disagree"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgree 
      Caption         =   "I Agree"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtAgreement 
      Height          =   3375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FRMLICNS.frx":08CA
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End Sub
Private Sub cmdAgree_Click()
    frmStart.Show
    frmSplash.Show 1
    frmLicense.Hide
End Sub
Private Sub cmdDisagree_Click()
    End
End Sub
Private Sub Form_Activate()
'Chips, Bits and Bytes License Agreement
'
'By selecting the button marked 'Agree' below, you are agreeing to the terms and conditions outlined in this License Agreement.
'1. Microsoft Corporation has no connection with the Author, and cannot be held responsible in any way for any damage whatsoever that may occur from the use, misuse or abuse of this program and associated files.
'2. The Author cannot be held responsible in any way for any damage whatsoever that may occur from the use, misuse or abuse of this program and associated files.
'3. Any requests for modifications, additions or alterations of any sort will be treated on a first come first served basis.  There is no guarantee that any requests of this sort will be honoured, and no time period can be given for the results of these requests."
'4. The Author warrants that this program is free from known bugs, (unless specifically stated in accompanying documentation), but if any programming errors are present, the Author cannot be held responsible for any damage arising from the use of this program.
'5. The User agrees to use this program at their own risk."
'
'This program is copyright (c) 1997 D. Robinson.
End Sub

Private Sub Form_Load()
    Width = 6690
    Height = 4485
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
