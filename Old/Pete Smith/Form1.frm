VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbFile 
      Height          =   7095
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12515
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Open "\\cbdxaaa\netlogon\sites\perthxk.csv" For Input As #1
        Line Input #1, stuff
        Line Input #1, TRAV
        Do Until EOF(1)
            Input #1, Group
            For lp = 1 To 9
                Input #1, junk
            Next lp
            Debug.Print Group
            cmd = "NET GROUP " & Group & " /DOMAIN >>c:\temp\curhouse.txt"
            Open "c:\temp\runme.bat" For Append As #2
                Print #2, cmd
            Close 2
        Loop
    Close 1
End Sub

Private Sub Command2_Click()
    rtbFile.FileName = "c:\temp\curhouse.txt"
    pos = 1
    Do Until pos = Len(rtbFile.Text)
        k = Mid(rtbFile.Text, pos, 8)
        If UCase(Left(k, 2)) = "PD" Then
            pos = pos + 8
            Open "c:\temp\junk.txt" For Append As #2
                Print #2, "NET USER " & k & " /DOMAIN >>lotsastuff.txt"
            Close 2
        End If
        pos = pos + 1
    Loop
    MsgBox "Done"
End Sub
