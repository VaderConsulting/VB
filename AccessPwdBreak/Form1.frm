VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "MS Access 97 Password Breaker"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Cdlg1 
      Left            =   3060
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   345
      Left            =   5610
      TabIndex        =   7
      Top             =   1380
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   465
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   1380
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Break the Password"
      Height          =   465
      Left            =   180
      TabIndex        =   0
      Top             =   2040
      Width           =   1605
   End
   Begin VB.Label Label4 
      Caption         =   "author@example.com"
      Height          =   225
      Left            =   4230
      TabIndex        =   8
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "The Pass Word is .."
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   2850
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "KNR's Access 97 Password Breaker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   150
      TabIndex        =   3
      Top             =   30
      Width           =   6165
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Path of the Access 97 Mdb here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   1080
      Width           =   6165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TChar(256) As String
Dim GetPwd As String * 14
Dim PWDs As String
Dim MyValCol
Dim MyPwr
Dim StrColVals As String
Dim StrMyPwr As String
Dim n%
Private Sub Command1_Click()
    Dim MdbnameStr As String
    MdbnameStr = UCase(Trim(Text1))
    If MdbnameStr = "" Then
        MsgBox "Pls Enter Path of the Mdb "
        Text1.SetFocus
        Exit Sub
    End If
    If Dir(MdbnameStr) = "" Then
        MsgBox "The file " & UCase(MdbnameStr) & " not found in the Specified Path.."
        Exit Sub
    End If
    StrColVals = "134,251,236,55,93,68,156,250,198,94,40,230,19,182"
    StrMyPwr = "1,2,4,8,16,32,64,128"
    MyValCol = Split(StrColVals, ",")
    MyPwr = Split(StrMyPwr, ",")
    Call GetAccessPwd(MdbnameStr)
End Sub
Function GetAccessPwd(MdbName As String)
    PWDs = ""
    If GetAccPWord(MdbName) = False Then Exit Function
    For n = 0 To 13
        WordOutChArray (n)
        If Asc(Chr(TChar(Asc(Mid(GetPwd, n + 1, 1))))) <> 0 Then
            PWDs = PWDs & Chr(TChar(Asc(Mid(GetPwd, n + 1, 1))))
        End If
    Next
    If Trim(PWDs) = "" Then
        MsgBox "Sorry there is No Password was set to this Database.."
    Else
        Text2 = PWDs
    End If
End Function
Function GetAccPWord(Pfile As String) As Boolean
    GetAccPWord = True
    Open Pfile For Binary As #1
        Seek #1, 67
        Get #1, , GetPwd
    Close #1
End Function
Function WordOutChArray(nChar As Integer) As Boolean
    Dim n, Powers, PowVal As Integer
    Dim MLoop, Place As Integer
    Dim SLoop As Integer
    Dim tmpCh As String
    For n = 0 To 255
        TChar(n) = n
    Next
    For Powers = 0 To 7
        PowVal = MyPwr(Powers)
        If (MyValCol(nChar) And PowVal) Then
            For MLoop = 0 To (256 / (PowVal * 2)) - 1
                Place = PowVal * MLoop * 2
                For SLoop = 0 To PowVal - 1
                    tmpCh = TChar(Place + SLoop)
                    TChar(Place + SLoop) = TChar(Place + SLoop + PowVal)
                    TChar(Place + SLoop + PowVal) = tmpCh
                Next
            Next
        End If
    Next
End Function
Private Sub Command2_Click()
    End
End Sub
Private Sub Command3_Click()
    Cdlg1.Filter = "*.mdb|*.mdb"
    Cdlg1.ShowOpen
    Text1 = ""
    If Cdlg1.FileName <> "" Then
        Text1 = Cdlg1.FileName
    End If
End Sub
