VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text to Doc Conversion"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdBrowseDestination 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "Browse for Destination File"
      Top             =   480
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowseSource 
      Caption         =   "..."
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Browse for Source File"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtDestination 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Destination Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Source Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Word As Object

Private Sub cmdBrowseDestination_Click()
    ' Allow the user to select Cancel
    On Error GoTo catchError
    ' Work with Common Dialog control
    With cdlFiles
        ' Continue until a filename is entered
        lblStatus.Caption = "Waiting for user selection"
        Do
            .InitDir = "S:\"
            .DialogTitle = "Select Destination File"
            .Filter = "Documents (*.doc)|*.doc|All Files (*.*)|*.*"
            .CancelError = True
            .ShowOpen
        Loop Until cdlFiles.Filename <> ""
        lblStatus.Caption = "Select 'Create'"
        txtDestination = cdlFiles.Filename
        
    End With
    Exit Sub
catchError:
    Select Case Err.Description
        Case "Cancel was selected."
            ' Nothing to do, the user selected Cancel
        Case Else
            lblStatus.Caption = "Error: " & Err.Description
    End Select
    
End Sub

Private Sub cmdBrowseSource_Click()
    ' Prepare variables
    Dim ExtPos As Integer
    Dim Filename As String
    Dim NewFilename As String
    ' Allow the user to select Cancel
    On Error GoTo catchError
    ' Work with Common Dialog control
    With cdlFiles
        ' Continue until a filename is entered
        lblStatus.Caption = "Waiting for user selection"
        Do
            .InitDir = "S:\"
            .DialogTitle = "Select Source File"
            .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            .CancelError = True
            .ShowOpen
        Loop Until cdlFiles.Filename <> ""
        lblStatus.Caption = "Select 'Create', or Browse for Destination File"
        txtSource = cdlFiles.Filename
        ' Enable the destination options
        cmdBrowseDestination.Enabled = True
        txtDestination.Enabled = True
        ' Determine the filename
        ExtPos = InStrRev(cdlFiles.Filename, ".")
        Filename = cdlFiles.Filename
        If ExtPos <> 0 Then
            ExtPos = ExtPos - 1
            Filename = Left(Filename, ExtPos)
        End If
        NewFilename = Filename & ".doc"
        txtDestination.Text = NewFilename
        
    End With
    Exit Sub
catchError:
    Select Case Err.Description
        Case "Cancel was selected."
            ' Nothing to do, the user selected Cancel
        Case Else
            lblStatus.Caption = "Error: " & Err.Description
    End Select
    
End Sub

Private Sub cmdCreate_Click()
    ' Ignore errors
    On Error Resume Next
    ' Save these settings as defaults for later retrieval
    Open "H:\textToDocFilenames.txt" For Output As #1
        Print #1, txtSource.Text
        Print #1, txtDestination.Text
    Close 1
    ' Open the pre-existing text file
    Set Word = CreateObject("Word.Application")
    Word.Visible = False
    Word.Application.Documents.Open txtSource.Text
    ' Save this as a Word Document, current version
    Word.ActiveDocument.SaveAs txtDestination.Text, wdFormatDocument
    ' Close the active Document
    Word.ActiveDocument.Close
    Word.Application.Quit
    MsgBox "Your Text Document '" & txtSource.Text & "' has been saved as '" & txtDestination.Text & "'", vbOKOnly + vbInformation
    lblStatus.Caption = "Select 'Exit' to finish"
    cmdCreate.Enabled = False
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub txtDestination_Change()
    cmdCreate.Enabled = True
End Sub

Private Sub Form_Load()
    ' Initialize variables
    Dim s As String
    Dim SourceFilename As String
    Dim DestinationFilename As String
    Me.Show
    Me.Refresh
    ' Ignore errors
    On Error Resume Next
    ' Get directory of H:, looking for specific file
    s = Dir("H:\textToDocFilenames.txt", vbNormal)
    If s <> "" Then
        ' Found the file, extract default filename for source file
        Open "H:\textToDocFilenames.txt" For Input As #1
            Line Input #1, SourceFilename
            Line Input #1, DestinationFilename
        Close 1
        txtSource.Text = SourceFilename
        txtDestination.Text = DestinationFilename
    End If
End Sub
