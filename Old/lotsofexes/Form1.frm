VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    lp = 0
    Open "c:\temp\serv1.txt" For Input As #1
        Do Until EOF(1)
            lp = lp + 1
            Line Input #1, serv1
            Open "c:\temp\bat\1\bat1" & lp & ".bat" For Output As #2
                Print #2, serv1
            Close 2
        Loop
    Close 1
    lp = 0
    Open "c:\temp\serv2.txt" For Input As #1
        Do Until EOF(1)
            lp = lp + 1
            Line Input #1, serv2
            Open "c:\temp\bat\2\bat2" & lp & ".bat" For Output As #2
                Print #2, serv2
            Close 2
        Loop
    Close 1
    lp = 0
    Open "c:\temp\serv.txt" For Input As #1
        Do Until EOF(1)
            lp = lp + 1
            Line Input #1, serv3
            Open "c:\temp\bat\3\bat3" & lp & ".bat" For Output As #2
                Print #2, "copy c:\temp\bat\wake.bat " & serv3 & "\c$\temp\wakeupall.bat"
            Close 2
        Loop
    Close 1
    lp = 0
    Open "c:\temp\serv.txt" For Input As #1
        Do Until EOF(1)
            lp = lp + 1
            Line Input #1, serv4
            Open "c:\temp\bat\4\bat4" & lp & ".bat" For Output As #2
                Print #2, "soon " & serv4 & " 60 c:\temp\wakeupall.bat"
            Close 2
        Loop
    Close 1
    Open "c:\temp\serv.txt" For Input As #1
        Do Until EOF(1)
            lp = lp + 1
            Line Input #1, serv
            Open "c:\temp\bat\clean\" & lp & ".bat" For Output As #2
                Print #2, "del " & serv & "\c$\temp\wake.exe"
                Print #2, "del " & serv & "\c$\temp\macs.csv"
            Close 2
        Loop
    Close 1
    End
End Sub
