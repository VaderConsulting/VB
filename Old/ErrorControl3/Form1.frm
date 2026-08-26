VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Control"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   14880
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   2040
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0464
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabDescription 
      Height          =   3975
      Left            =   10080
      TabIndex        =   3
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Description"
      TabPicture(0)   =   "Form1.frx":08BE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtDescription"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   360
         Width           =   4455
      End
   End
   Begin MSComctlLib.StatusBar sbrBrowser 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   9195
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlBrowse 
      Left            =   1320
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08DA
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D2C
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":117E
            Key             =   "Home"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   14655
      ExtentX         =   25850
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList imlErrors 
      Left            =   600
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15D0
            Key             =   "Error"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":172A
            Key             =   "Question"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1884
            Key             =   "Logo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CD6
            Key             =   "Unknown"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":91D8
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F472
            Key             =   "Description"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   120
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView tvwError 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7011
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imlErrors"
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Begin VB.Menu mnuOpenDatabase 
            Caption         =   "Database"
         End
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuValidate 
         Caption         =   "Validate Database"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuErrorCodes 
         Caption         =   "Error Codes"
         Begin VB.Menu mnuGeneralErrors 
            Caption         =   "General"
         End
         Begin VB.Menu mnuICE 
            Caption         =   "ICE Reference"
         End
         Begin VB.Menu mnuInitialisation 
            Caption         =   "Initialisation"
         End
         Begin VB.Menu mnuLogo 
            Caption         =   "Logo Requirements"
         End
         Begin VB.Menu mnuWise 
            Caption         =   "Windows Installer Error Codes"
         End
      End
      Begin VB.Menu mnuFormatted 
         Caption         =   "Formatted Strings"
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuTreeWhatsThis 
      Caption         =   "Whats This Help"
      Visible         =   0   'False
      Begin VB.Menu mnuTreeView 
         Caption         =   "What's This"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastPath As String
Public Filename As String
Public WebPosition As Long
'Public MSI As WindowsInstaller.Database
Public MSIFilename As String

Private ThisControl As Control

Private Sub Browser_StatusTextChange(ByVal Text As String)
    sbrBrowser.SimpleText = Text
End Sub

Private Sub Form_Activate()
    ' "C:\Documents\VB6\ErrorControl\ActivePerl.txt"
    UpdateCaption
End Sub

Sub UpdateCaption()
    If Filename = "" Then
        Me.Caption = "Error Control"
    Else
        Me.Caption = "Error Control - " & Filename
    End If
    Browser.Navigate2 App.Path & "\notopic_0pk4.htm"
End Sub

Private Sub Form_Load()
    Dim Args As String, i As Integer
    LastPath = GetSetting(App.EXEName, "Setup", "Last Path", App.Path)
    
    'tabDescription.TabEnabled(1) = False
    
    Me.Show
    Me.Refresh
    
    Args = Trim(Command$ & "")
    
    If Args <> "" Then
        Args = Replace(Args, Chr(34), "")
        If Dir(Args) <> "" Then
            Filename = Args
            Screen.MousePointer = vbHourglass
            OpenTextFile Filename
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    DoExit
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    DoExit
End Sub

Sub DoExit()
    SaveSetting App.EXEName, "Setup", "Last Path", LastPath
    End
End Sub

Private Sub mnuFormatted_Click()
    ShowBrowser True
    Browser.Navigate2 App.Path & "\Help\msi\tref_0qp0.htm"
End Sub

Private Sub mnuGeneralErrors_Click()
    ShowBrowser True
    Browser.Navigate2 App.Path & "\Help\msi\code_13ub.htm"
End Sub

Private Sub mnuICE_Click()
    ShowBrowser True
    Browser.Navigate App.Path & "\Help\msi\ices_40th.htm"
End Sub

Private Sub mnuInitialisation_Click()
    ShowBrowser True
    Browser.Navigate App.Path & "\Help\msi\msi_5svn.htm"
End Sub

Sub DoOpenDatabase()
    Dim ReturnCode As Integer
    With cdlFiles
        .InitDir = LastPath
        .Filter = "Installer Databases (*.msi)|*.msi|All Files (*.*)|*.*"
        .Filename = ""
        .CancelError = True
        
        On Error Resume Next
            Do
                .ShowOpen
                If Err.Number <> 0 Then Exit Sub
            Loop Until .Filename <> ""
        On Error GoTo 0
        
        MSIFilename = .Filename
        
        txtDescription.Text = ""
        ShowBrowser False
        
        Screen.MousePointer = vbHourglass
        'Set MSI = OpenMSI(MSIFilename, msiOpenDatabaseModeTransact, ReturnCode)
        Screen.MousePointer = vbDefault
        tvwError.Nodes.Clear
        txtDescription.Text = ""
        Browser.Visible = True
        Browser.Navigate2 App.Path & "\notopic_0pk4.htm"
        mnuValidate_Click
        OpenTextFile App.Path & "\validate.txt"
    End With
End Sub

Sub DoOpenTextFile()
    With cdlFiles
        .InitDir = LastPath
        .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .Filename = ""
        .CancelError = True
        
        On Error Resume Next
            Do
                .ShowOpen
                If Err.Number <> 0 Then Exit Sub
            Loop Until .Filename <> ""
        On Error GoTo 0
        
        Filename = .Filename
        
        txtDescription.Text = ""
        ShowBrowser False
        
        Screen.MousePointer = vbHourglass
        OpenTextFile (Filename)
        Screen.MousePointer = vbDefault
        
    End With
End Sub

Private Sub mnuLogo_Click()
    ShowBrowser True
    Browser.Navigate2 App.Path & "\Help\msi\tcod_908j.htm"
End Sub

Sub OpenTextFile(strFilename As String)
    Dim strLine As String, strInput() As String, iceLine As String
    Dim i As Integer, lp As Integer, MessageText As String
    Dim strTemp As String, intTemp As String, IceNo As Integer, IceString As String
    Dim n As Node, j As Integer
    Dim IceErrors(255) As Integer
    Dim LogoError(255) As String
    Dim ThisNode As String
    Dim RootNode As String
    Dim FileDescriptionString As String
    Dim FileIceString As String
    Dim strInputLine1 As String
    Dim URL As String
    Dim ErrorType As String
    
    UpdateCaption
    
    Filename = App.Path & "\validate.txt"
    
    LastPath = GetFilePath(Filename)
    
    Me.Refresh
    sbrBrowser.SimpleText = "Clearing data"
    tvwError.Nodes.Clear
    Me.Refresh
    
    Set n = tvwError.Nodes.Add(, , "ERROR", "", "Error")
    n.Expanded = True
    Set n = tvwError.Nodes.Add(, , "WARNING", "", "Question")
    n.Expanded = True
    Set n = tvwError.Nodes.Add(, , "INFO", "", "Logo")
    n.Expanded = True
    'Set n = tvwError.Nodes.Add(, , "Unknown", "", "Unknown")
    'n.Expanded = True
    
    Open Filename For Input As #1
        i = 1
        Do
            
            strInputLine1 = ""
            
            Line Input #1, strInputLine1
            
            'If Left(strInputLine1, 11) = "Evaluation:" Then
                'MessageText = Replace(strInputLine1, vbTab, vbCrLf)
                iceLine = strInputLine1

                IceNo = CInt(Mid(iceLine, 4, 6))
                IceString = Trim(Mid(iceLine, 4, 6))

                On Error Resume Next
                    ErrorType = Trim(Mid(iceLine, 13, 10))
                    RootNode = ErrorType
                    'RootNode = GetRoot(MessageText)
                    Set n = tvwError.Nodes.Add(RootNode, tvwChild, "ICE" & IceString, "ICE" & IceString, "Folder")
                    n.Tag = "ICE" & IceNo

                    ThisNode = tvwError.Nodes.Item(RootNode & MessageText).Text
                    If ThisNode <> "" Then
                        ThisNode = ""
                    Else
                        IceErrors(IceNo) = IceErrors(IceNo) + 1
                        Set n = tvwError.Nodes.Add("ICE" & IceString, tvwChild, RootNode & MessageText, MessageText, "Description")
                        n.Tag = "ICE" & IceString
                        If IceErrors(IceNo) = 1 Then
                            tvwError.Nodes.Item("ICE" & IceString).Text = "ICE" & IceString & " (" & IceErrors(IceNo) & " error)"
                        Else
                            tvwError.Nodes.Item("ICE" & IceString).Text = "ICE" & IceString & " (" & IceErrors(IceNo) & " errors)"
                        End If
                    End If
                On Error GoTo 0
            
                sbrBrowser.SimpleText = "Loading data: " & MessageText

                MessageText = ""
                IceNo = 0
                IceString = ""
'            Else
'                ' ************ Logo errors - treated VERY differently
'                If Left(strInputLine1, 15) = "The shared file" Then
'
'                    On Error Resume Next
'                        Set n = tvwError.Nodes.Add("Logo", tvwChild, "Shared Files in Application Folder.", "Shared Files in Application Folder.", "Folder")
'                    On Error GoTo 0
'                    Set n = tvwError.Nodes.Add("Shared Files in Application Folder.", tvwChild, "A" & strInputLine1, strInputLine1, "Description")
'
'                ElseIf (Left(strInputLine1, 8) = "The file" = True) And (strInputLine1 Like "*into a directory*" = False) Then
'
'                    On Error Resume Next
'                        Set n = tvwError.Nodes.Add("Logo", tvwChild, "Shared DLL Reference Count flag not set.", "Shared DLL Reference Count flag not set.", "Folder")
'                        Set n = tvwError.Nodes.Add("Shared DLL Reference Count flag not set.", tvwChild, "B" & strInputLine1, strInputLine1, "Description")
'                    On Error GoTo 0
'
'
'                ElseIf Left(strInputLine1, 8) = "The file" And strInputLine1 Like "*into a directory*" Then
'
'                    On Error Resume Next
'                        Set n = tvwError.Nodes.Add("Logo", tvwChild, "File installed into a directory other than Program Files.", "File installed into a directory other than Program Files.", "Folder")
'                        Set n = tvwError.Nodes.Add("File installed into a directory other than Program Files.", tvwChild, "C" & strInputLine1, strInputLine1, "Description")
'                    On Error GoTo 0
'
'                ElseIf strInputLine1 Like "Win16 executable*" Then
'
'                    On Error Resume Next
'                        Set n = tvwError.Nodes.Add("Logo", tvwChild, "16 bit file included in package.", "16 bit file included in package.", "Folder")
'                        Set n = tvwError.Nodes.Add("16 bit file included in package.", tvwChild, "D" & strInputLine1, strInputLine1, "Description")
'                    On Error GoTo 0
'
'                ElseIf Left(strInputLine1, 26) = "No files have been flagged" Then
'
'                    Set n = tvwError.Nodes.Add("Logo", tvwChild, strInputLine1, strInputLine1, "Description")
'
'                Else
'
'                    On Error Resume Next
'                        Set n = tvwError.Nodes.Add("Logo", tvwChild, strInputLine1, strInputLine1, "Folder")
'
'                        If strInputLine1 <> "N/A" Then
'                            Set n = tvwError.Nodes.Add(strInputLine1, tvwChild, "E" & strInputLine1, strInputLine1, "Description")
'                        End If
'                    On Error GoTo 0
'                End If
'                sbrBrowser.SimpleText = "Loading data: " & strInputLine1
'            End If
            
            ' Keep the user informed
            j = j + 1
            If j = 10 Then
                DoEvents
                j = 0
            End If
        Loop Until EOF(1)
    Close 1
    sbrBrowser.SimpleText = "Done"
End Sub

Function GetRoot(strNodeText As String) As String
    GetRoot = "Unknown"
    
    'Errors
    If strNodeText Like ("Invalid*") Or strNodeText Like ("Upgrade*") Or strNodeText Like ("The directory*") Or strNodeText Like ("Component '*") Or strNodeText Like ("The shortcut *") Then
        GetRoot = "Error"
    End If
    
    If strNodeText Like ("Class {*") Or strNodeText Like "*is a Font and must be installed*" Or strNodeText Like "*found in AdvtUISequence table. No UI is allowed*" Then
        GetRoot = "Error"
    End If
    
    If strNodeText Like "*not found in Directory table*" Or strNodeText Like "*CostFinalize missing from sequence table:*" Then
        GetRoot = "Error"
    End If
    
    ' Warnings
    If strNodeText Like ("Reg*") Or strNodeText Like ("*String") Or strNodeText Like ("No*") Or strNodeText Like ("Property*") Or strNodeText Like ("Mismatched*") Then
        GetRoot = "Question"
    End If
    If strNodeText Like ("Complete*") Or strNodeText Like ("String*") Or strNodeText Like ("*defined*") Or strNodeText Like ("Component:*") Or strNodeText Like ("WARNING:*") Then
        GetRoot = "Question"
    End If
    
    If strNodeText Like ("The file '*") Or strNodeText Like ("KeyPath for Component*") Then
        GetRoot = "Question"
    End If
    
    ' Logo
    If strNodeText Like ("*shared*") Then
        GetRoot = "Logo"
    End If
    
    If GetRoot = "Unknown" Then
        AddError "The root for " & strNodeText & " is not known."
    End If
End Function

Private Sub mnuOpenDatabase_Click()
    DoOpenDatabase
End Sub

Private Sub mnuOpenTextFile_Click()
    DoOpenTextFile
End Sub

Private Sub mnuTreeView_Click()
    ThisControl.ShowWhatsThis
End Sub

Private Sub mnuValidate_Click()
' #############################################################################
    Dim cubName As String

    Dim Cmd As String
    cubName = App.Path & "\darice.cub"
    Cmd = App.Path & "\msival2.exe " & MSIFilename & " " & cubName & " -l " & App.Path & "\validate.txt"
    Shell Cmd, vbMaximizedFocus
    Debug.Print Cmd
End Sub

Private Sub mnuWise_Click()
    ShowBrowser True
    Browser.Navigate App.Path & "\Help\msi\erro_89f7.htm"
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            DoOpenDatabase
        Case 2
            DoExit
    End Select
End Sub

Private Sub tvwError_Click()
    ClickError
End Sub

Private Sub tvwError_KeyUp(KeyCode As Integer, Shift As Integer)
    ClickError
End Sub

Sub ClickError()
    Dim IceString As String, IceText As String, URL As String
    
    On Error Resume Next
        IceText = tvwError.Nodes(tvwError.SelectedItem.Index).Text & ""
    On Error GoTo 0
    
    If IceText = "" Then
        txtDescription.Text = ""
        Exit Sub
    End If
    
    IceString = Trim(tvwError.Nodes.Item(tvwError.SelectedItem.Index).Tag)
    
    If Left(IceText, 3) <> "ICE" Then
        If IceString = "" Then
            If IceText Like "Contains*" Then
                URL = App.Path & "\Help\msi\over_286r.htm"
                txtDescription.Text = IceText
                ShowBrowser True
                Browser.Navigate2 URL
            ElseIf IceText Like "Shared DLL Reference*" Then
                URL = App.Path & "\Help\msi\tref_09np.htm"
                txtDescription.Text = IceText
                ShowBrowser True
                Browser.Navigate2 URL
            ElseIf IceText = "Shared Files in Application Folder" Then
                URL = App.Path & "\Help\msi\over_5mb2.htm"
                txtDescription.Text = "This error should be ignored.  See the help for more information."
                txtDescription.Text = IceText
                ShowBrowser True
                Browser.Navigate2 URL
            ElseIf IceText Like "No files have been flagged*" Then
                txtDescription.Text = "This error should be ignored."
                ShowBrowser False
            ElseIf IceText Like "Has an empty Key Path*" Then
                txtDescription.Text = "Further information cannot be found, but this error should be rectified."
                ShowBrowser False
            ElseIf IceText = "16 bit file included in package." Then
                txtDescription.Text = "This error should be ignored."
                ShowBrowser False
            Else
                txtDescription.Text = IceText
            End If
        Else
            txtDescription.Text = tvwError.Nodes(tvwError.SelectedItem.Index).Text & ""
            
            If IceString <> "" Then
                DoHelp IceString
            End If
        End If
    Else
        txtDescription.Text = Trim(Left(IceText, 6))
        DoHelp IceString
    End If
End Sub

' #############################################################################
' Extract the path for the given full filename
Public Function GetFilePath(strFilename As String) As String
' #############################################################################
    Dim Position As Integer
    Position = InStrRev(strFilename, "\")
    GetFilePath = Left(strFilename, Position - 1)
End Function

' #############################################################################
' Determine and display the correct URL for the currently selected ICE Number
Sub DoHelp(IceString As String)
' #############################################################################
    Dim Prefix As String, Suffix As String, Content As String, URL As String
    Dim IceNo As Integer
    
    If Left(IceString, 3) = "ICE" Then
        If IsNumeric(Mid(IceString, 4, 5)) Then
            IceNo = CInt(Mid(IceString, 4, 5))
        End If
    Else
        Exit Sub
    End If
    
    Prefix = App.Path & "\Help\msi\ices_1"
    Suffix = ".htm"
    
    Select Case IceNo
        Case 1 To 9
            Content = "xr" & Chr(Asc("l") + (IceNo - 1))
        Case 10 To 19
            Content = "xt" & Chr(Asc("c") + (IceNo) - 10)
        Case 20 To 25
            Content = "xv" & Chr(Asc("4") + (IceNo) - 20)
        Case 26 To 29
            Content = "xv" & Chr(Asc("a") + (IceNo) - 26)
        Case 30 To 33
            Content = "xw" & Chr(Asc("w") + (IceNo) - 30)
        Case 34 To 39
            Content = "xx" & Chr(Asc("0") + (IceNo) - 34)
        Case 40 To 49
            Content = "xy" & Chr(Asc("o") + (IceNo) - 40)
        Case 50 To 59
            Content = "y0" & Chr(Asc("g") + (IceNo) - 50)
        Case 60
            Content = "y28"
        Case 61
            Content = "y29"
        Case 62 To 69
            Content = "y2" & Chr(Asc("a") + (IceNo) - 62)
        Case 70 To 79
            Content = "y4" & Chr(Asc("0") + (IceNo) - 70)
        Case 80 To 88
            Content = "y5" & Chr(Asc("s") + (IceNo) - 80)
        Case 89
            Content = "y61"
        Case 90 To 94
            Content = "y7" & Chr(Asc("k") + (IceNo) - 90)
        Case Else
            Content = ""
            AddError IceString & " is not a known ICE string."
    End Select
    
    If Content <> "" Then
        URL = Prefix & Content & Suffix
    Else
        URL = App.Path & "\notopic_0pk4.htm"
    End If
    
    ShowBrowser True
    Browser.Navigate2 URL

End Sub

Sub ShowBrowser(State As Boolean)
    Browser.Visible = State
    'sbrBrowser.Visible = State
End Sub

Sub AddError(strMessage As String)
    'tabDescription.TabEnabled(1) = True
    'lstErrors.AddItem strMessage, 0
End Sub

Private Sub tvwError_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Set ThisControl = tvwError
        PopupMenu mnuTreeWhatsThis
    End If
End Sub

Private Sub tvwError_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If LCase(Right(Data.Files(1), 3)) = "txt" Then
        Filename = Data.Files(1)
        OpenFile Filename
    End If
End Sub

' #############################################################################
' OpenMSI - Opens the given MSI filename, returning a Database object
Public Function OpenMSI(MSIPath As String, Mode As Integer, ReturnCode As Integer) As WindowsInstaller.Database
' #############################################################################
    Dim oinstaller As WindowsInstaller.Installer
    Set oinstaller = CreateObject("WindowsInstaller.Installer")
    On Error Resume Next
        Set OpenMSI = oinstaller.OpenDatabase(MSIPath, Mode)
        ReturnCode = Err.Number
    On Error GoTo 0
    Debug.Print "OpenMSI Returned:" & vbCrLf & CheckError
    Set oinstaller = Nothing
End Function

' #############################################################################
' CheckError - check for Installer errors
Function CheckError() As String
' #############################################################################
    Dim message, errRec
    Dim oinstaller As WindowsInstaller.Installer
    Set oinstaller = CreateObject("Windowsinstaller.Installer")
    CheckError = 0
    If Err = 0 Then Exit Function
    message = Err.Source & " " & Hex(Err) & ": " & Err.Description
    If Not oinstaller Is Nothing Then
        Set errRec = oinstaller.LastErrorRecord
        If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText: CheckError = errRec.IntegerData(1)
    End If
    If CheckError = 2268 Then
        Err.Clear
        Exit Function
    End If
    CheckError = message
End Function
