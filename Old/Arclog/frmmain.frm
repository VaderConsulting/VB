VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ArcLog"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstStatus 
      Height          =   1035
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Timer tmrCheck 
      Interval        =   60000
      Left            =   2400
      Top             =   840
   End
   Begin VB.ListBox lstSource 
      Height          =   645
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdArchive 
      Caption         =   "Archive"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "5"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtDestination 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "\\monitor01\logs"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "\\CBDXAAI\Kixlog$"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblErrors 
      Caption         =   "Errors in current batch:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   5055
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   840
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   7920
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7920
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblWait 
      Caption         =   "Time (minutes)"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblDestination 
      Caption         =   "Destination"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblSource 
      Caption         =   "Source"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    ' Program variables
    Public SourceDir As String
    Public DestinationDir As String
    Public ArcTime As Integer
    
Private Sub Form_Load()
    '  ***** Determine command line arguments *****
    If Command$ <> "" Then
        tmrCheck.Enabled = False ' Turn off timer, because we are operating from command line
        pos1 = InStr(1, Command$, "/s=")       ' Get position of /s=
        If pos1 <> 0 Then
            pos2 = InStr(pos1, Command$, " ")  ' Get position of space after /s=
            If pos2 = 0 Then pos2 = 100
        End If
        pos3 = InStr(1, Command$, "/d=")       ' Get position of /d=
        If pos3 <> 0 Then
            pos4 = InStr(pos3, Command$, " ")  ' Get position of space after /d=
            If pos4 = 0 Then pos4 = 100
        End If
        pos5 = InStr(1, Command$, "/t=")       ' Get position of /t=
        '                  The '3' in the following lines indicate that the
        '                  command line argument is preceded by an argument prefix
        '                  that is 3 characters long.  ie /s=
        '
        If pos1 <> 0 Then SourceDir = Mid(Command$, pos1 + 3, pos2 - pos1 - 3)
        If pos3 <> 0 Then DestinationDir = Mid(Command$, pos3 + 3, pos4 - pos3 - 3)
        If pos5 <> 0 Then ArcTime = Val(Mid(Command$, pos5 + 3, 3))
    End If
    ' Set defaults
    If SourceDir = "" Then SourceDir = "\\CBDXAAI\Kixlog$"
    If DestinationDir = "" Then DestinationDir = "\\MONITOR01\logs"
    If ArcTime = 0 Then ArcTime = 5
    ' The following will proceed with an automated 'run' and then end the program.
    If Command$ <> "" Then
        cmdArchive_Click
        Unload Me
        End
    End If
    txtSource = SourceDir
    txtDestination = DestinationDir
    txtTime = ArcTime
End Sub

Private Sub cmdAbout_Click()
    msg = msg & "ArcLog is a file archiving utility that may be used with command line" & vbCrLf
    msg = msg & "arguments, or without arguments to be used with GUI controls." & vbCrLf
    msg = msg & "It is used to move files from one directory to another, but only if the files" & vbCrLf
    msg = msg & "have not been accessed for a period of time.  This period is measured in minutes" & vbCrLf
    msg = msg & "and may be specified at the command prompt, along with other options." & vbCrLf
    msg = msg & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "Usage:" & vbCrLf
    msg = msg & "      Arclog.exe /s=SOURCE_DIR /d=DESTINATION_DIR /t=TIME" & vbCrLf & vbCrLf
    msg = msg & "Defaults:" & vbCrLf
    msg = msg & "SOURCE = \\CBDXAAI\Kixlog$" & vbCrLf
    msg = msg & "DESTINATION = \\MONITOR01\C$\Data\Logs" & vbCrLf
    msg = msg & "TIME = 5" & vbCrLf
    MsgBox msg, vbInformation + vbOKOnly, "About"
End Sub

Private Sub cmdArchive_Click()
    tmrCheck.Enabled = False ' Disable timer whilst in this procedure
    SourceDir = txtSource
    DestinationDir = txtDestination
    ArcTime = txtTime
    Screen.MousePointer = vbHourglass
    lblStatus = "Retrieving list..."
    lblStatus.Refresh
    lstSource.Clear
    If Dir(SourceDir & "\renamed.log") <> "" Then
        Kill SourceDir & "\Renamed.log"
    End If
    lstStatus.Clear
    ' Get a list of all files in source directory that are older than Arctime,
    ' and move them to the destination directory.
    directory = Dir(SourceDir & "\*.*")
    Do Until directory = ""
        lblStatus = "Checking times on files found..."
        lblStatus.Refresh
        Filetime = FileDateTime(SourceDir & "\" & directory)
        CompareTime = DateAdd("s", ArcTime * 60, Date & " " & Time)
        If Filetime <= CompareTime Then
            lblStatus = "Processing " & directory
            lstSource.AddItem directory
            lblNumber = lstSource.ListCount & " source files"
            ' Rename file to prevent writes to it, but with error checking
            On Error Resume Next
            Do
                ' Debug.Print "Renaming " & SourceDir & "\" & directory & " as " & SourceDir & "\Renamed.log"
                Name SourceDir & "\" & directory As SourceDir & "\Renamed.log"
                errorNumber = Err.Number
                Err.Clear
                DoEvents
                errorCount = errorCount + 1
                If errorCount >= 10 Then                              ' Only retry the copy 10 times
                   errorCount = 0                                     ' The main reason this will fail is that
                   lblStatus = "Aborting file copy after 10 attempts" ' the file exists in the destination directory
                   lstStatus.AddItem "Abort on " & directory & " error: " & errorNumber & " (" & Error$(errorNumber) & ")"
                   lstStatus.ListIndex = lstStatus.ListCount - 1
                   lstStatus.Refresh
                   GoTo Continue
                End If
            Loop Until errorNumber = 0
            errorCount = 0
            ' Copy file to destination directory
            ' Debug.Print "Copying " & SourceDir & "\" & "Renamed.log" & " as " & DestinationDir & "\" & directory
            FileCopy SourceDir & "\" & "Renamed.log", DestinationDir & "\" & directory
            ' Delete renamed file
            Kill SourceDir & "\Renamed.log"
Continue:
            frmMain.Refresh
            DoEvents
        End If
        ' Get next file from source directory
        directory = Dir
    Loop
    lstStatus.Clear
    lstStatus.AddItem Time & " " & lblNumber & " processed"
    lstSource.Clear
    lblNumber = "0 Items"
    lblStatus = "Idle"
    Screen.MousePointer = vbDefault
    tmrCheck.Enabled = True ' Re-enable timer
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub tmrCheck_Timer()
    Static Counter      ' Keep track of how many times we get here
    Counter = Counter + 1
    If Counter = 3 Then ' At the third time, start processing files
        Counter = 0
        cmdArchive_Click
    End If
End Sub
