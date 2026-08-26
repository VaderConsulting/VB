VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Generator"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3360
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLength 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Text            =   "14"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "1"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtPasswords 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Length"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Number"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    txtPasswords = ""
End Sub

Private Sub cmdCreate_Click()
    Dim lp As Integer
        Dim lengthloop As Integer
        Dim pw As String
        Dim i As Integer
        Dim NoOfOptions As Integer
        
        NoOfOptions = 3
        Randomize (Timer)
        For lp = 1 To CInt(txtNumber.Text)
            For lengthloop = 1 To CInt(txtLength.Text)
                'Do
                i = (Rnd(1) * (NoOfOptions - 1)) + 1
                'Loop Until i > 1
                '
                Select Case i
                    Case 1
                        pw = pw & Chr((Rnd(1) * 9) + 48) ' Numeric chars
                    Case 2
                        pw = pw & Chr((Rnd(1) * 25) + 65) ' Uppercase chars
                    Case 3
                        pw = pw & Chr((Rnd(1) * 25) + 97) ' Lowercase chars
                    Case 4
                        pw = pw & Chr((Rnd(1) * 15) + 32) ' Other characters (1)
                    Case 5
                        pw = pw & Chr((Rnd(1) * 5) + 91) ' Other characters (2)
                    Case 6
                        pw = pw & Chr((Rnd(1) * 3) + 123) ' Other characters (3)
                End Select
                'pw = pw & " [" & i & "] "
            Next
            txtPasswords.Text = txtPasswords.Text & pw & vbCrLf
            pw = ""
        Next
End Sub
