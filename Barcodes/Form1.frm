VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Barcode Generator"
   ClientHeight    =   1695
   ClientLeft      =   2715
   ClientTop       =   1395
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   3990
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Large"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Small"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Text            =   "12345678"
      Top             =   360
      Width           =   3732
   End
   Begin VB.Label Label1 
      Caption         =   "Enter text for barcode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDraw_Click()
    Load frmBarcode
    frmBarcode.Show
    Call DrawBarcode(Text1, frmBarcode.picBarcode)
    
    MinWidth = 2 * Text1.Left + Text1.Width
    pw = 2 * frmBarcode.picBarcode.Left + frmBarcode.picBarcode.Width
    fw = MinWidth
    If pw > fw Then fw = pw
    frmBarcode.Width = fw
    frmBarcode.PrintForm
    Unload frmBarcode
End Sub


Private Sub cmdExit_Click()
    End

End Sub

Private Sub Form_Activate()

    optSize(1) = 1

End Sub

Private Sub optSize_Click(Index As Integer)
    frmBarcode.picBarcode.ScaleMode = 3
    
    Select Case Index
        Case 0
            frmBarcode.picBarcode.Height = frmBarcode.picBarcode.Height * (1.4 * 40 / frmBarcode.picBarcode.ScaleHeight)
            frmBarcode.picBarcode.FontSize = 8
        Case 1
            frmBarcode.picBarcode.Height = frmBarcode.picBarcode.Height * (2.4 * 40 / frmBarcode.picBarcode.ScaleHeight)
            frmBarcode.picBarcode.FontSize = 10
        Case 2
            frmBarcode.picBarcode.Height = frmBarcode.picBarcode.Height * (3 * 40 / frmBarcode.picBarcode.ScaleHeight)
            frmBarcode.picBarcode.FontSize = 14
    End Select
    
    
    'Call Text1_Change

End Sub

'Private Sub Text1_Change()
'    Call DrawBarcode(Text1, frmBarcode.picBarcode)
'
'    MinWidth = 2 * Text1.Left + Text1.Width
'    pw = 2 * frmBarcode.picBarcode.Left + frmBarcode.picBarcode.Width
'    fw = MinWidth
'    If pw > fw Then fw = pw
'    Form1.Width = fw
'
'End Sub


