VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nucleus"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3030
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   855
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   480
   End
   Begin VB.Label lblProcesses 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim Hostname As String
  Dim strTemp As String
  Dim lp As Integer
  Dim Applications() As String
  ' Start here
  
  ' Ensure no other copies of app running
  If App.PrevInstance Then
    Unload Me
    End
  End If
  ' Get computer name
  Me.Show
  Me.Refresh
  Hostname = Environ$("computername")
  lblProcesses.Caption = "Nucleus is ready"
  lblProcesses.Refresh
  If Command$ = "" Then
    strTemp = GetValue("", "Software\CSC\Nucleus", "Applications", False)
    Applications() = Split(strTemp, ",")
    For lp = 0 To UBound(Applications())
      Debug.Print Applications(lp)
      ' Execute Task
      lblProcesses.Caption = "Spawning " & Applications(lp)
      lblProcesses.Refresh
      ExecuteTask Applications(lp)
    Next lp
    tmrMain.Enabled = True
  Else
    On Error Resume Next
    If Dir(Command$) <> "" Then
      ExecuteTask Command$
      tmrMain.Enabled = True
    Else
      ' Can't find this app. (FILE NOT FOUND)
      Unload Me
      End
    End If
    On Error GoTo 0
  End If
  lblProcesses.Caption = "Nucleus is ready"
  lblProcesses.Refresh
End Sub

Private Sub tmrMain_Timer()
  Dim oZip As CGZipFiles, ZipFileName As String, D As String
  Static Counter As Long
  ' Check if all running applications are complete
  FindCompletedProcesses
  lblProcesses.Caption = "Waiting for " & Processes.Count & " processes"
  frmMain.Refresh
  If (Processes.Count = 0) Or _
     (Counter = 60) Then 'Either no processes or time is up
    ' Get all files in c:\Nucleus, and zip up
    tmrMain.Enabled = False
    If Dir("C:\Nucleus\*.*") <> "" Then ' files exist to zip up, so proceed
      Set oZip = New CGZipFiles
    
      ' Give Zip File a Name / Path
      ZipFileName = "c:\Nucleus\Nucleus.zip"
      oZip.ZipFileName = ZipFileName
      oZip.UpdatingZip = False ' ensures a new zip is created
      D = Dir("c:\Nucleus\*.*")
      Do Until D = ""
        oZip.AddFile "C:\Nucleus\" & D
        D = Dir
      Loop
      ' This next line ensures the app finishes 'cleanly'
      On Error Resume Next
      ' Make zip file
      res = oZip.MakeZipFile
      If res = 0 Then ' no error creating zip file
        If Dir("c:\Nucleus\Nucleus.zip") <> "" Then ' zip succeeded, so proceed
          ' Get zipped files and copy to \\CBDXAAI\Nucleus
          FileCopy "c:\Nucleus\Nucleus.zip", "\\CBDXAAI\Nucleus\" & Environ$("computername") & ".zip"
        End If
        ' Remove all files in c:\Nucleus - but safely
        If Dir("\\CBDXAAI\Nucleus\" & Environ$("computername") & ".zip") <> "" Then
          If Err.Number = 0 Then
            Kill "c:\Nucleus\*.*"
          End If
        End If
      End If
    End If
    'All done, end program
    Set oZip = Nothing
    lblProcesses.Caption = "Nucleus shutting down"
    lblProcesses.Refresh
    Unload Me
  End If
End Sub
