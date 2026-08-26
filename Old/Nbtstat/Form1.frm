VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Clean up IP's"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdGetMAC 
      Caption         =   "Get MAC's"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstFull 
      Height          =   4155
      Left            =   3960
      TabIndex        =   4
      Top             =   840
      Width           =   4095
   End
   Begin VB.ListBox lstBlank 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Text            =   "h:\macstoinsert.csv"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "h:\blankMACs.csv"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExport_Click()
    Open txtOutput For Output As #4
        For a = 0 To lstFull.ListCount - 1
            Print #4, lstFull.List(a)
            lblStatus = "Exporting... " & a & " of " & lstFull.ListCount - 1
            lblStatus.Refresh
        Next a
    Close 4
End Sub

Private Sub cmdGetMAC_Click()
    For a = 0 To lstBlank.ListCount - 1
        lblStatus = a & " of " & lstBlank.ListCount - 1
        lblStatus.Refresh
        varPos1 = InStr(1, lstBlank.List(a), ",")
        varHost = Left(lstBlank.List(a), varPos1 - 1)
        varPos2 = InStr(varPos1 + 1, lstBlank.List(a), ",")
        varIP = Mid(lstBlank.List(a), varPos2 + 1, 20)
        For b = 1 To Len(varIP) ' remove " " from the IP
            If Mid(varIP, b, 1) = " " Then
                varLeft = Left(varIP, b - 1)
                varRight = Right(varIP, Len(varIP) - (b + 1))
                varIP = varLeft & varRight
            End If
        Next b
        'varCmd = "NBTSTAT -A " & varIP & " >c:\temp\nbtinfo.txt"
        varCmd = "NBTSTAT -a " & varHost & " >c:\temp\nbtinfo.txt"
        Open "c:\temp\nbt.bat" For Output As #2
            Print #2, varCmd
        Close 2
        On Error Resume Next
            Kill "c:\temp\nbtinfo.txt"
        On Error GoTo 0
        Shell "c:\temp\nbt.bat", vbHide
        Dim t As Date
        Dim tt As Date
        t = Time
        tt = DateAdd("s", 1, t)
        Do Until Time > tt
            DoEvents
        Loop
        Open "c:\temp\nbtinfo.txt" For Input As #3
            Do Until EOF(3)
                varMAC = ""
                Line Input #3, varStuff
                If varStuff = "Host not found." Then Exit Do
                varPos3 = InStr(1, varStuff, "MAC Address = ")
                If varPos3 <> 0 Then
                    varMAC = Mid(varStuff, varPos3 + 14, 25)
                    lstFull.AddItem varHost & "," & varMAC & "," & varIP
                End If
                
                lstFull.Refresh
            Loop
        Close 3
        
    Next a
End Sub

Private Sub cmdRefresh_Click()
    Open txtInput For Input As #1
        Do Until EOF(1)
            Line Input #1, varData
            lstBlank.AddItem varData
        Loop
    Close 1
End Sub
