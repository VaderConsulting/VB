VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Open "h:\computers.txt" For Input As #1
        Do Until EOF(1)
            Line Input #1, CName
            Label1 = CName
            Label1.Refresh
            cmd = "ping " & CName & ">h:\ping.txt"
            Open "h:\pin.bat" For Output As #2
                Print #2, cmd
            Close 2
            Shell "h:\pin.bat", vbHide
            T = Time
            Do Until DateDiff("s", T, Time) > 2
                DoEvents
            Loop
            Open "h:\ping.txt" For Input As #3
                Do Until EOF(3)
                    Line Input #3, IP
                    If InStr(1, IP, "[") > 0 Then
                        l = InStr(1, IP, "[")
                        l = l + 1
                        le = Mid(IP, l, 40)
                        r = InStr(1, le, "]")
                        
                        IPAddress = Left(le, r - 1)
                        Open "h:\outputIP.csv" For Append As #4
                            Print #4, IPAddress & "," & CName
                        Close 4
                    End If
                Loop
            Close 3
        Loop
    Close 1
    MsgBox "Done"
End Sub
