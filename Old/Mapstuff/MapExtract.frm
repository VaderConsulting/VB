VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   13950
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List5 
      Height          =   3960
      Left            =   11640
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   120
      Pattern         =   "*.tab"
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin VB.ListBox List4 
      Height          =   3960
      Left            =   8760
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin VB.ListBox List3 
      Height          =   3960
      Left            =   5880
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   12360
      TabIndex        =   0
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Filename"
      Height          =   255
      Left            =   11640
      TabIndex        =   11
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Northwest"
      Height          =   255
      Left            =   8760
      TabIndex        =   10
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Northeast"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Southeast"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Southwest"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Open "c:\temp\mapinfo2.csv" For Output As #2
        Print #2, "filename,sw_x,sw_y,se_x,sw_y,ne_x,ne_y,nw_x,nw_y"
        For lp = 0 To File1.ListCount - 1
            Open File1.Path & "\" & File1.List(lp) For Input As #1
                Line Input #1, junk
                Line Input #1, junk
                Line Input #1, junk
                Line Input #1, junk
                Line Input #1, info5
                Line Input #1, junk
                Line Input #1, info1
                Line Input #1, info2
                Line Input #1, info3
                Line Input #1, info4
                If InStr(1, info5, "gif") <> 0 Then
                    sw_x = Mid(info1, 3, 13)
                    sw_y = Mid(info1, 17, 14)
                    se_x = Mid(info2, 3, 13)
                    se_y = Mid(info2, 17, 14)
                    ne_x = Mid(info3, 3, 13)
                    ne_y = Mid(info3, 17, 14)
                    nw_x = Mid(info4, 3, 13)
                    nw_y = Mid(info4, 17, 14)
                    file_name = Mid(info5, 18, Len(info5) - 19)
                    Print #2, file_name & "," & sw_x & "," & sw_y & "," & se_x & "," & se_y & "," & ne_x & ","; ne_y & "," & nw_x & ","; nw_y
                    List1.AddItem sw_x & " " & sw_y
                    List2.AddItem se_x & " " & se_y
                    List3.AddItem ne_x & " " & ne_y
                    List4.AddItem nw_x & " " & nw_y
                    List5.AddItem file_name
                    
                End If
            Close 1
        Next lp
    Close 2
End Sub

Private Sub Form_Load()
    File1.Path = "d:\mapinfo"
End Sub
