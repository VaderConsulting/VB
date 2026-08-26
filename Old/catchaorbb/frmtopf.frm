VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTopF 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Top Five Players"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "frmTopF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDlg1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
   End
   Begin VB.CommandButton cmdClose 
      Height          =   1455
      Left            =   2760
      Picture         =   "frmTopF.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdFont 
      Height          =   1455
      Left            =   480
      Picture         =   "frmTopF.frx":096A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblNames2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "The top Five Players For Catch a Orbb"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1080
      TabIndex        =   20
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "5"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   17
      Top             =   3360
      Width           =   90
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "4"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   16
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "3"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   15
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "2"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   14
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   13
      Top             =   1440
      Width           =   90
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "No."
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   5
      Left            =   3120
      TabIndex        =   11
      Top             =   3360
      Width           =   45
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   4
      Left            =   3120
      TabIndex        =   10
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   3
      Left            =   3120
      TabIndex        =   9
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   2
      Left            =   3120
      TabIndex        =   8
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   1440
      Width           =   45
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Score"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   0
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   420
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   3360
      Width           =   45
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   45
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Name"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   420
   End
End
Attribute VB_Name = "frmTopF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************'
'Filename :       frmPlay.frm                                '
'Author:          Coenraad Steenkamp                         '
'Created:         09/05/2000                                 '
'Operating System Windows 9x & NT ,2000 ,Me                  '
'                                                            '
'Description:                                                '
'This form show's the top 5 five scores to the user .First   '
'the scores are loaded from the disk into an array and then  '
'it is sorted by the bubble sort mehod.The from also have a  '
'option to change the font.                                  '
'                                                            '
'************************************************************'
Private Sub cmdClose_Click()
   'Declare Local Variables
   Dim strAnswer As String
   'Ask's the user if he wants to Exit
   strAnswer = MsgBox("Are you sure?", vbYesNo, "Confirmation...")
   If strAnswer = vbYes Then
      'Unloads my Top Five Form
      Unload frmTopF
      frmPlay.Show
   End If
End Sub
Private Sub cmdFont_Click()
    'Declare Local variables
    Dim iX As Integer
    'Shows the Commondlg Font Box
    CommonDlg1.ShowFont
    'Start the loop from 0 to 5 to set the Fonts
    For iX = 0 To 5
        'Set the fonts for the Names
        lblName(iX).FontName = CommonDlg1.FontName
        lblName(iX).FontSize = CommonDlg1.FontSize
        lblName(iX).FontBold = CommonDlg1.FontBold
        lblName(iX).FontItalic = CommonDlg1.FontItalic
        lblName(iX).FontUnderline = CommonDlg1.FontUnderline
        lblName(iX).FontStrikethru = CommonDlg1.FontStrikethru
        'Set the fonts for the Score
        lblScore(iX).FontName = CommonDlg1.FontName
        lblScore(iX).FontSize = CommonDlg1.FontSize
        lblScore(iX).FontBold = CommonDlg1.FontBold
        lblScore(iX).FontItalic = CommonDlg1.FontItalic
        lblScore(iX).FontUnderline = CommonDlg1.FontUnderline
        lblScore(iX).FontStrikethru = CommonDlg1.FontStrikethru
        'Set the fonts for the Label Number
        lblLabel(iX).FontName = CommonDlg1.FontName
        lblLabel(iX).FontSize = CommonDlg1.FontSize
        lblLabel(iX).FontBold = CommonDlg1.FontBold
        lblLabel(iX).FontItalic = CommonDlg1.FontItalic
        lblLabel(iX).FontUnderline = CommonDlg1.FontUnderline
        lblLabel(iX).FontStrikethru = CommonDlg1.FontStrikethru
        'End's the for loop
    Next iX
End Sub
Private Sub Form_Unload(Cancel As Integer)
   'Error handeling goto EH
   On Error GoTo EH
   'Declare variables
   Dim iX As Integer
   Dim iRecLength As Integer
   miFileNum = FreeFile
   iRecLength = Len(gudRec)
   'Open the file from the Disk
   Open "A:\Orbb.txt" For Random As #miFileNum Len = iRecLength
   'Error Handeling for "Disk not ready"
EH:
   If Err.Number = 71 Then
      intRet = MsgBox("The drive is not ready.", _
               vbExclamation + vbAbortRetryIgnore, "Disk not in Drive!")
      Select Case intRet
         Case vbRetry
            Resume
         Case vbAbort
            End
         Case vbIgnore
            End
      End Select
   End If
   'Use the LOF()Function(Length of File)to Calculate number of records in the file
   giNumRecords = LOF(miFileNum) / iRecLength
   For iX = 1 To 5
      gudRec.StrName = frmTopF.lblName(iX)
      gudRec.iScore = frmTopF.lblScore(iX)
   Next iX
   'Put the records on the disk
   Put miFileNum, giNumRecords, gudRec
End Sub
