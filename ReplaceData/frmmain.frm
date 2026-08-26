VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Replace data in file"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtReplaceString 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtSearchString 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Replace String"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4080
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   4080
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Search String"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   3
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

Private Sub cmdGo_Click()
    Dim strFileData As String
    
    Open txtFilename & ".bak" For Output As #1
        Open txtFilename For Input As #2
            Do Until EOF(2)
                Line Input #2, strFileData
                strFileData = Replace(strFileData, txtSearchString, txtReplaceString, 1, -1, vbTextCompare)
                Print #1, strFileData
            Loop
        Close 2
    Close 1
    
    Kill txtFilename
    Name txtFilename & ".bak" As txtFilename
    End
End Sub

Private Sub Form_Load()
    Dim strArgs() As String
    
    If Command$ <> "" Then
        strArgs() = Split(Command$, " ")
        If strArgs(0) <> "" And strArgs(1) <> "" And strArgs(2) <> "" Then
            txtFilename = strArgs(0)
            txtSearchString = strArgs(1)
            txtReplaceString = strArgs(2)
            cmdGo_Click
        End If
    End If
End Sub

Private Sub txtFilename_Change()
    If txtFilename <> "" And txtSearchString <> "" And txtReplaceString <> "" Then
        cmdGo.Enabled = True
    Else
        cmdGo.Enabled = False
    End If
End Sub

Private Sub txtReplaceString_Change()
    If txtFilename <> "" And txtSearchString <> "" And txtReplaceString <> "" Then
        cmdGo.Enabled = True
    Else
        cmdGo.Enabled = False
    End If
End Sub

Private Sub txtSearchString_Change()
    If txtFilename <> "" And txtSearchString <> "" And txtReplaceString <> "" Then
        cmdGo.Enabled = True
    Else
        cmdGo.Enabled = False
    End If
End Sub
