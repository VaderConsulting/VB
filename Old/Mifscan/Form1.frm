VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstMonitor 
      Height          =   7665
      Left            =   5280
      TabIndex        =   7
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ListBox lstPC 
      Height          =   7665
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtMonitor 
      Height          =   285
      Left            =   5160
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.DirListBox dirFiles 
      Height          =   8415
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblMonitor 
      Caption         =   "Monitor Serial Number"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblPC 
      Caption         =   "PC Serial Number"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dirFiles_Click()
    dirmonitor = Dir(dirFiles.List(dirFiles.ListIndex) & "\Monitor.mif")
    dirPC = Dir(dirFiles.List(dirFiles.ListIndex) & "\PC.mif")
    dir0 = Dir(dirFiles.List(dirFiles.ListIndex) & "\compaq0.mif")
    dir2 = Dir(dirFiles.List(dirFiles.ListIndex) & "\compaq2.mif")
    vardir = dirFiles.List(dirFiles.ListIndex)
    If dirmonitor <> "" Then
        Open vardir & "\" & dirmonitor For Input As #1
            Do Until EOF(1)
                Line Input #1, varData
                If InStr(1, varData, "Serial Number") <> 0 Then
                    Line Input #1, varData
                    Line Input #1, varData
                    Line Input #1, varData
                    varPos = InStr(1, varData, Chr(34))
                    MonitorSerial1 = Mid(varData, varPos + 1, Len(varData) - (varPos + 1))
                    a = 0
                End If
            Loop
        Close 1
    End If
    If dirPC <> "" Then
        Open vardir & "\" & dirPC For Input As #1
            Do Until EOF(1)
                Line Input #1, varData
                If InStr(1, varData, "Serial Number") <> 0 Then
                    Line Input #1, varData
                    Line Input #1, varData
                    Line Input #1, varData
                    varPos = InStr(1, varData, Chr(34))
                    PCSerial1 = Mid(varData, varPos + 1, Len(varData) - (varPos + 1))
                    a = 0
                End If
            Loop
        Close 1
    End If
    If dir0 <> "" Then
        Open vardir & "\" & dir0 For Input As #1
            Do Until EOF(1)
                Line Input #1, varData
                If InStr(1, varData, "Serial Number") <> 0 Then
                    Line Input #1, varData
                    Line Input #1, varData
                    Line Input #1, varData
                    varPos = InStr(1, varData, Chr(34))
                    PCSerial2 = Mid(varData, varPos + 1, Len(varData) - (varPos + 1))
                    a = 0
                End If
            Loop
        Close 1
    End If
    If dir2 <> "" Then
        Open vardir & "\" & dir2 For Input As #1
            Do Until EOF(1)
                Line Input #1, varData
                If InStr(1, varData, "Serial Number") <> 0 Then
                    Line Input #1, varData
                    Line Input #1, varData
                    Line Input #1, varData
                    varPos = InStr(1, varData, Chr(34))
                    MonitorSerial2 = Mid(varData, varPos + 1, Len(varData) - (varPos + 1))
                    a = 0
                End If
            Loop
        Close 1
        If MonitorSerial1 <> "" Or MonitorSerial2 <> "" Then
            If MonitorSerial1 <> "" Then
                lstMonitor.AddItem MonitorSerial1
            Else
                lstMonitor.AddItem MonitorSerial2
            End If
        End If
        If PCSerial1 <> "" Or PCSerial2 <> "" Then
            If PCSerial1 <> "" Then
                lstPC.AddItem PCSerial1
            Else
                lstPC.AddItem PCSerial2
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    dirFiles.Path = "F:\Mifs"
End Sub
