VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit personal shortcuts"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   ClipControls    =   0   'False
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWarning 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   1920
      Top             =   720
   End
   Begin VB.TextBox txtWarning 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "frmEdit.frx":0742
      Top             =   1560
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.CommandButton cmdExample 
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdManual 
      Caption         =   "Edit manually"
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      ToolTipText     =   "Load notepad with personal.txt"
      Top             =   3360
      Width           =   1095
   End
   Begin ComctlLib.TabStrip tabLaunch 
      Height          =   375
      Left            =   1440
      TabIndex        =   21
      Top             =   240
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   661
      TabWidthStyle   =   2
      Style           =   1
      ShowTips        =   0   'False
      TabFixedWidth   =   494
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   26
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "1"
            Key             =   ""
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "2"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "3"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "4"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "5"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "6"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "7"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "8"
            Key             =   ""
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "9"
            Key             =   ""
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "10"
            Key             =   ""
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "11"
            Key             =   ""
            Object.Tag             =   "5"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "12"
            Key             =   ""
            Object.Tag             =   "6"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "13"
            Key             =   ""
            Object.Tag             =   "7"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "14"
            Key             =   ""
            Object.Tag             =   "8"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "15"
            Key             =   ""
            Object.Tag             =   "9"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "16"
            Key             =   ""
            Object.Tag             =   "10"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "17"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab18 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "18"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab19 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "19"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab20 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "20"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab21 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "21"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab22 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "22"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab23 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "23"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab24 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "24"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab25 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "25"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab26 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "26"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPersonal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Text            =   "1"
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      ToolTipText     =   "Program path"
      Top             =   840
      Width           =   6975
   End
   Begin VB.TextBox txtDirectory 
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      ToolTipText     =   "The program's working directory"
      Top             =   1560
      Width           =   6975
   End
   Begin VB.TextBox txtIcon 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      ToolTipText     =   "Picture to show on button"
      Top             =   1920
      Width           =   6975
   End
   Begin VB.TextBox txtTooltip 
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      ToolTipText     =   "Helpful text to display above button"
      Top             =   2280
      Width           =   6975
   End
   Begin VB.CommandButton cmdElipsis 
      Caption         =   "..."
      Height          =   255
      Index           =   0
      Left            =   9600
      TabIndex        =   8
      ToolTipText     =   "Browse for program"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdElipsis 
      Caption         =   "..."
      Height          =   255
      Index           =   2
      Left            =   9600
      TabIndex        =   7
      ToolTipText     =   "Browse for icon"
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtArgs 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      ToolTipText     =   "File to load with the program"
      Top             =   1200
      Width           =   6975
   End
   Begin VB.CommandButton cmdElipsis 
      Caption         =   "..."
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   5
      ToolTipText     =   "Browse for file"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Clear all boxes"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Show info"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      ToolTipText     =   "Same as OK, but does not close this window"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "Close window"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   8880
      TabIndex        =   0
      ToolTipText     =   "Save info, and close window"
      Top             =   3360
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdlBrowse 
      Left            =   840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fmeFullPath 
      Caption         =   "Full program executable"
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   9855
      Begin VB.Label lblFullPath 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Label lblPath 
      Caption         =   "Program path"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblDirectory 
      Caption         =   "Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblIcon 
      Caption         =   "Icon (*.ico or *.bmp)"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblTip 
      Caption         =   "Tool tip (helpful text to display)"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblArguments 
      Caption         =   "File to load (optional)"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
    msg = "Launcher program V " & App.Major & "." & App.Minor & "." & App.Revision & " written for HMAS Arunta by LSET Dave Robinson"
    MsgBox msg, vbOKOnly
End Sub

Private Sub cmdCancel_Click()
    Close #1
    frmEdit.Hide
End Sub

Private Sub cmdClear_Click()
    txtPath = ""
    txtArgs = ""
    txtDirectory = ""
    txtIcon = ""
    txtTooltip = ""
    cmdExample.Picture = frmEdit.Icon
End Sub

Private Sub cmdElipsis_Click(Index As Integer)
    On Error Resume Next
    Rem browse for program
    If Index = 0 Then
        cdlBrowse.FileName = ""
        cdlBrowse.InitDir = "d:\"
        cdlBrowse.Filter = "All Files (*.*)|*.*"
        cdlBrowse.ShowOpen
        If Not (cdlCancel) Then
            txtPath = cdlBrowse.FileName
            lp = Len(txtPath)
            Do Until a$ = "\" Or lp < 1
                a$ = Mid$(txtPath, lp, 1)
                lp = lp - 1
            Loop
            txtDirectory = Left$(txtPath, lp)
        End If
    End If
    Rem browse for file
    If Index = 1 Then
        cdlBrowse.FileName = ""
        cdlBrowse.InitDir = "h:\"
        cdlBrowse.Filter = "All Files (*.*)|*.*"
        cdlBrowse.ShowOpen
        If Not (cdlCancel) Then
            txtArgs = cdlBrowse.FileName
        End If
    End If
    Rem browse for icon
    If Index = 2 Then
        cdlBrowse.FileName = ""
        cdlBrowse.InitDir = "J:\icons"
        cdlBrowse.Filter = "Icon files (*.ico)|*.ico|Bitmap Files (*.bmp)|*.bmp"
        cdlBrowse.ShowOpen
        If Not (cdlCancel) Then
            txtIcon = cdlBrowse.FileName
        End If
    End If
    frmEdit.SetFocus
End Sub

Private Sub cmdexample_Click()
    On Error GoTo Wrong
    If txtPath <> "" Then
        If txtDirectory <> "" Then ChDir (txtDirectory)
        Shell lblFullPath, vbNormalFocus
    End If
    GoTo Done
Wrong:
    msg = "Error " & Err & " (" & Err.Description & ")" & vbCrLf & vbCrLf
    msg = msg & "Ensure the program exists and you have read and execute rights to it."
    MsgBox msg, vbCritical + vbOKOnly, "Critical Error"
Done:
End Sub

Private Sub cmdApply_Click()
    subApply
    If tmpIcon(txtPersonal) <> "" Then cmdExample.Picture = LoadPicture(tmpIcon(txtPersonal))
End Sub

Private Sub cmdManual_Click()
    Shell ("Notepad.exe " & gPersonal), vbNormalFocus
End Sub

Private Sub cmdOK_Click()
    gError = 0
    subApply
    If gError <> 1 Then
        frmEdit.Hide
        frmStart.Show
        gError = 0
    End If
End Sub
Public Sub subApply()
    For a = 1 To 10
        If tmpFile(a) <> "" Then
            If UCase$(Right$(tmpIcon(a), 3)) <> "ICO" Then
                If UCase$(Right$(tmpIcon(a), 3)) <> "BMP" Then
                    MsgBox "Please use a .ico or .bmp file for the shortcut " & a, vbCritical + vbOKOnly
                    tabLaunch.TabIndex = a
                    txtIcon.SetFocus
                    gError = 1
                End If
            End If
            If tmpDirs(a) = "" Then
                MsgBox "Please enter a working directory for the shortcut " & a, vbCritical + vbOKOnly
                tabLaunch.TabIndex = a
                txtDirectory.SetFocus
                gerr0r = 1
            End If
            If tmpTip(a) = "" Then
                MsgBox "Please enter some descriptive text for the shortcut " & a, vbCritical + vbOKOnly
                tabLaunch.TabIndex = a
                txtTooltip.SetFocus
                gError = 1
            End If
        End If
    Next a
    Close 1
    personalpath = Left$(gPersonal, Len(gPersonal) - 12)
    Rem delete existing backup of personal.txt
    If Dir(personalpath & "personal.bak") <> "" Then Kill (personalpath & "personal.bak")
    Rem rename existing file to personal.bak
    Name gPersonal As personalpath & "personal.bak"
    Open personalpath & "Personal.txt" For Output As #1
    For lp = 1 To 26
        If tmpProgram(lp) <> "" Then
            Write #1, lp
            Print #1, tmpProgram(lp)
            Print #1, tmpFile(lp)
            Print #1, tmpDirs(lp)
            Print #1, tmpIcon(lp)
            Print #1, tmpTip(lp)
            gProgram(lp + 77) = tmpProgram(lp)
            gFile(lp + 77) = tmpFile(lp)
            gDir(lp + 77) = tmpDirs(lp)
            frmStart.cmdButton(lp + 77).Picture = LoadPicture(tmpIcon(lp))
            frmStart.cmdButton(lp + 77).ToolTipText = tmpTip(lp)
            frmStart.cmdButton(lp + 77).Visible = True
        Else
            gProgram(lp + 77) = ""
            gDir(lp + 77) = ""
            Rem this next line clears the unwanted icon
            frmStart.cmdButton(lp + 77).Picture = frmEdit.Picture
            frmStart.cmdButton(lp + 77).ToolTipText = ""
            frmStart.cmdButton(lp + 77).Visible = False
        End If
    Next lp
    Close 1
    Rem cmdExample.Picture = LoadPicture(tmpIcon(txtPersonal))
    Rem frmStart.Refresh
    Rem add code to update pictures on frmstart.cmdbuttons(x)
End Sub


Private Sub Form_Activate()
    Rem this next line enables error trapping in case H: drive doesn't exist
    Rem On Error GoTo nohome
    Rem check for personal.txt file
    On Error Resume Next
    If Dir(gPersonal) <> "" Then
        Rem file exists - open it and get all available personal programs
        Open gPersonal For Input As #1
        a = 1
        Do Until EOF(1)
            Line Input #1, lne
            If lne = "" Then Exit Do
            tmp = CInt(lne)
            Line Input #1, tmpProgram(tmp)
            Line Input #1, tmpFile(tmp)
            Line Input #1, tmpDirs(tmp)
            Line Input #1, tmpIcon(tmp)
            Line Input #1, tmpTip(tmp)
            a = a + 1
        Loop
        Close 1
    End If
    On Error GoTo 0
    txtPersonal = 1
    subUpdate
    GoTo skip
nohome:
    If UCase(Command$) <> "DEBUG" Then
        msg = "Launcher cannot modify your user programs:" & vbCrLf
        msg = msg & "Home drive not found."
        MsgBox msg, vbCritical + vbOKOnly
        frmEdit.Hide
        frmStart.Show
    Else
        MsgBox "Debug mode.  For more complete testing ensure H: drive exists."
    End If
skip:
    On Error Resume Next
    frmEdit.tabLaunch.SelectedItem = 1
End Sub

Private Sub tabLaunch_Click()
    Rem keep current settings
    tmpProgram(txtPersonal) = txtPath
    tmpFile(txtPersonal) = txtArgs
    tmpDirs(txtPersonal) = txtDirectory
    tmpIcon(txtPersonal) = txtIcon
    tmpTip(txtPersonal) = txtTooltip
    
    Rem clear warning label
    txtWarning.Visible = False
    Rem change selected text number
    txtPersonal = tabLaunch.SelectedItem.Index
    
    Rem now select the new settings
    a = txtPersonal
    txtPath = tmpProgram(a)
    txtArgs = tmpFile(a)
    txtDirectory = tmpDirs(a)
    txtIcon = tmpIcon(a)
    txtTooltip = tmpTip(a)
    
    If txtIcon <> "" Then
        cmdExample.Picture = LoadPicture(tmpIcon(a))
    Else
        cmdExample.Picture = frmEdit.Icon
    End If
    frmEdit.Refresh
End Sub

Private Sub subUpdate()
    If txtPersonal = "" Then txtPersonal = 1
    a = txtPersonal
    txtPath = tmpProgram(a)
    txtArgs = tmpFile(a)
    txtDirectory = tmpDirs(a)
    txtIcon = tmpIcon(a)
    txtTooltip = tmpTip(a)
    frmEdit.Refresh
    Rem clear old icon
    cmdExample.Picture = frmEdit.Icon
    Rem clear tooltip
    cmdExample.ToolTipText = ""
    If txtIcon <> "" Then
        If Dir(txtIcon) <> "" Then
            cmdExample.Picture = LoadPicture(tmpIcon(txtPersonal))
        Else
            MsgBox "Problem loading icon.  Please check icon path"
        End If
    End If
    If txtTooltip <> "" Then cmdExample.ToolTipText = txtTooltip
End Sub

Private Sub tmrWarning_Timer()
    txtWarning.Visible = False
End Sub

Private Sub txtArgs_Change()
    tmpFile(txtPersonal) = txtArgs
    lblFullPath = txtPath & " " & txtArgs
End Sub

Private Sub txtDirectory_Change()
    tmpDirs(txtPersonal) = txtDirectory
End Sub

Private Sub txtIcon_Change()
    tmpIcon(txtPersonal) = txtIcon
End Sub

Private Sub txtPath_Change()
    tmpProgram(txtPersonal) = txtPath
    lblFullPath = txtPath & " " & txtArgs
    If InStr(1, txtPath, " ") > 0 Then
        txtWarning.Visible = True
        tmrWarning.Enabled = True
    End If
End Sub

Private Sub txtTooltip_Change()
    tmpTip(txtPersonal) = txtTooltip
    cmdExample.ToolTipText = txtTooltip
End Sub

