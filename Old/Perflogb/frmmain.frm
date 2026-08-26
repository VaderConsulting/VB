VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PerfLogb - Performance Log Extract tool for Batch files"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin MSChart20Lib.MSChart chtLogs 
      Height          =   6735
      Left            =   4080
      OleObjectBlob   =   "frmMain.frx":0742
      TabIndex        =   4
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox lstFiles 
      Height          =   6300
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtLogPath 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "\\CBDXAAI\kixlog$\batch"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblPath 
      Caption         =   "Log Files path"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Base 1
    Dim TotalTime(3000) As Integer
    Dim Username(3000) As String
    Dim Sitename(3000) As String

Private Sub cmdOpen_Click()
    lstFiles.Clear
    Logs = Dir(txtLogPath & "\*.log")
    While Logs <> ""
        lstFiles.AddItem Logs
        Logs = Dir()
    Wend
    'Workstations = 1
    If lstFiles.ListCount > 0 Then
        For lst = 0 To lstFiles.ListCount - 1
            logfilename = lstFiles.List(lst)
            Workstations = 1
            Open txtLogPath & "\" & logfilename For Input As #1
                While Not EOF(1)
                    On Error Resume Next                 ' required to get around files still open
                    Line Input #1, dat
                    pos = InStr(1, dat, ":")
                    l = Left(dat, pos - 1)
                    r = Mid(dat, pos + 2, Len(dat) - pos)
                    Select Case l
                        Case "Logon Start"
                            If startflag <> 1 Then
                                st = r
                            End If
                            startflag = 1
                        Case "Logon Finish"
                            If finishflag <> 1 Then
                                fin = r
                            End If
                            finishflag = 1
                        Case "Username"
                            If userflag <> 1 Then
                                user = r
                            End If
                            userflag = 1
                        Case "SiteLoc"
                            If siteflag <> 1 Then
                                siteloc = Mid(dat, pos + 1, Len(dat) - pos)
                            End If
                            siteflag = 1
                    End Select
                Wend
                startflag = 0
                finishflag = 0
                userflag = 0
                siteflag = 0
                If st <> "" And fin <> "" And user <> "" And siteloc <> "" Then
                    TotalTime(lst + 1) = DateDiff("s", st, fin)
                    Username(lst + 1) = user
                    Sitename(lst + 1) = siteloc
                    Open "c:\temp\PerfLogB.csv" For Append As #2
                    lne = Left(lstFiles.List(lst + 1), Len(lstFiles.List(lst + 1)) - 4) & ","
                    lne = lne & Username(lst + 1) & ","
                    lne = lne & Sitename(lst + 1) & ","
                    lne = lne & TotalTime(lst + 1)
                    If Val(TotalTime(lst + 1)) > 0 Then
                        Print #2, lne
                    End If
                    Close 2
                End If
            Close 1
        Next lst
    End If
    MsgBox "Extract complete"
    End
End Sub

Public Sub Form_Load()
    If Dir("c:\temp\PerfLogB.csv") <> "" Then Kill "c:\temp\PerfLogB.csv"
    cmdOpen_Click
End Sub

