VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile Save Wizard"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   7815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optMode 
      BackColor       =   &H80000009&
      Caption         =   "Replay"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H80000009&
      Caption         =   "Copy"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   15
      Top             =   1080
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkADAware 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Redirected Folder(s)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtA2 
      Height          =   285
      Left            =   2040
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtA1 
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Frame fmeResults 
      Caption         =   "Results"
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   7575
      Begin VB.ListBox lstResults 
         Height          =   1230
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdTwo 
      Caption         =   "Next"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdOne 
      Caption         =   "Previous"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "Capture"
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblQ2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblQ1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblNavigation 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   7575
   End
   Begin VB.Label lblDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   7335
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FIND_FOLDER = "PreSOEData"
Private Const FOR_READING = 1
Private Const FOR_WRITING = 2
Private Const FOR_APPENDING = 8

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_RETURNFSANCESTORS = 8
Private Const BIF_EDITBOX = 16
Private Const BIF_NEWDIALOGSTYLE = 64
Private Const BIF_BROWSEINCLUDEFILES = 16384

Private Const S_OK = &H0                ' Success
Private Const S_FALSE = &H1             ' The Folder is valid, but does not exist
Private Const E_INVALIDARG = &H80070057 ' Invalid CSIDL Value

Private Const SHGFP_TYPE_CURRENT = 0
Private Const SHGFP_TYPE_DEFAULT = 1
Private Const MAX_PATH = 260

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Enum CSIDL
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEO = &HE
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMPUTERSNEARME = &H3D
    CSIDL_FLAG_PER_USER_INIT = &H800
    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_MASK = &HFF00
End Enum

Dim StepNo As Integer
Dim DefaultDestination As String, Username As String, bADAware As Boolean, bRegIsNT As Boolean
Dim fso As Scripting.FileSystemObject

'Private Sub ProfileCapture()
'    'Author:         Andrew Rowlands (Empired Ltd)
'    'Date Created:   07/08/02
'    'Date Modified   29/08/02 - Dave Robinson (Empired)
'    '                Default values added for backup server and username
'    '
'    'Copyright info: This script is licensed to Rio Tinto Exploration for the duration of the Rio tinto XP desktop SOE project.
'    '            This script may not be copied or modified in part or as a whole without the express permission on its author.
'    'Modifications:  Requests for modification or functionality enhancements should be sent to the author.
'    'Modem of Trust: Users of this script are trusted to abide by the conditions of use outlined above. Use of this script for
'    '                purposes other than intended are not permitted. This script has been left unencrypted (in a trusting state)
'    '                for educational and cross-platform purposes only. Discovered breaches of this copywright will force the Author
'    '                to encrypt future scripts.
'    'Author Contact: The Author can be contacted via email on: andy@perthweb.net.au or mobile: +61 412 833400
'
'
'    'Drive Types = 0 - Unknown, 1 - Removable, 2 - Fixed, 3 - Network, 4 - CDROM, 5 - RAM DISK
'
'    Dim fso, objFile, colFiles, objDir, colDir, bFolderExists, objShell, strPreSoeData, strPSDPath, strSearchDrives
'    Dim strSearchDrive, i, bPSDFound, strPSDDrive, strDataDest, colDrives, objDrive, strDrive, strBakFolder, strTempFoldername
'    Dim strFileLine, strOldOSTLoc, strNewOSTLoc, strOldPSTLoc, strNewPSTLoc, strOldPABLoc, strNewPABLoc, bRegIsNT
'    Dim objProcLog, intButton1, intButton2, bADAware, bContinue, arrFileType, arrExcludeFolders, objRootFolder
'
'    'Set objShell = CreateObject("Scripting.Shell")
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    'Set Environ = objShell.Environment("Process")
'
'    'On Error Resume Next
'    MsgBox "Performing a User Data Capture may take an extended amount of time." & vbCrLf & _
'               "Please ensure all applications (eg. Outlook) are closed before continuing.", , "Profile Capture Tool"
'    intButton1 = MsgBox("Does this client have folder redirection enabled?" & vbCrLf & vbCrLf & _
'                       "If the client DOES NOT have 'Redirected Folders', or you are not sure," & vbCrLf & _
'                       "click 'No'." & vbCrLf & vbCrLf & "If the client DOES have 'Redirected Folders', click 'Yes'.", vbYesNo, "Client Info")
'    Select Case intButton1
'        Case vbYes
'            bADAware = True
'        Case vbNo
'            bADAware = False
'        Case Else
'            End
'    End Select
'
'    DefaultDestination = "\\expperfile02\SOEBackup"
'    DefaultDestination = "C:\Temp\SOE"
'
'    strDataDest = InputBox("Enter a destination storage point for storing Users backed up data.", "Backup" & _
'            " Storage Point", DefaultDestination)
'    If strDataDest = "" Or strDataDest = "Enter Path Here" Then
'        End
'    ElseIf fso.FolderExists(strDataDest) = False Then
'        MsgBox "Destination not found. Process aborted."
'        End
'    End If
'
'    If Right(strDataDest, 1) = "\" Then
'        strDataDest = Left(strDataDest, Len(strDataDest) - 1)
'    End If
'
'    Username = Environ$("Username")
'
'    strBakFolder = InputBox("Whose Data and Settings are you backing up?" & vbCrLf & "(Enter the Username)", _
'                    "User Details", Username)
'
'    If strBakFolder = "" Or strBakFolder = "Enter Username Here" Then
'        MsgBox "You must enter a Username." & vbCrLf & "This name Is required To make a unique " & _
'                "backup folder."
'        End
'    Else
'        If fso.FolderExists(DefaultDestination) Then
'            strTempFoldername = DefaultDestination
'            Do While fso.FolderExists(strTempFoldername)
'                i = i + 1
'                strTempFoldername = strBakFolder & Trim(Str(i))
'            Loop
'            MsgBox "A folder with that name already exists. A number has been added to the folder" & _
'                        " name to make it unique."
'            MkDir (strDataDest & "\" & strTempFoldername)
'            i = 0
'            strBakFolder = strTempFoldername
'        Else
'            MkDir (DefaultDestination)
'        End If
'    End If
'
'    intButton2 = MsgBox("You have chosen to proceed with the following parameters:" & vbCrLf & vbCrLf & _
'                    "Backup Location: " & strDataDest & vbCrLf & _
'                    "User Backup Dir: " & strBakFolder & vbCrLf & _
'                    "Client AD Aware: " & bADAware & vbCrLf & vbCrLf & "Are these correct?", vbYesNo, "Summary")
'    Select Case intButton2
'        Case vbYes
'            bContinue = True
'        Case vbNo
'            bContinue = False
'        Case Else
'            End
'    End Select
'    If bContinue = False Then
'        fso.DeleteFolder DefaultDestination, True
'        End
'    End If
'
'    ' TODO: change to remove use of fso
'    Set objProcLog = fso.OpenTextFile(DefaultDestination & "\DataCapture.log", FOR_WRITING, True)
'    LogMessage "Log Start " & Date & " " & Time
'    objProcLog.WriteBlankLines 1
'
'    LogMessage "Starting Sub: CopyBackupDir"
'    CopyBackupDir strDataDest & strBakFolder 'Custom Public Sub
'    LogMessage "Starting Sub: GetOutlookSettings"
'    GetOutlookSettings 'Custom Private Sub
'    LogMessage "Starting Sub: BackupUserOLKFiles"
'    BackupUserOLKFiles 'Custom Private Sub
'            'If bRegIsNT = True And bADAware = False Then
'            '   logmessage "Starting Sub: CopyAppData"
'            '   CopyAppData 'Custom Private Sub (For Windows NT OS's only)
'            'End If
'    LogMessage "Starting Sub: CopyDesktopData"
'    CopyDesktopData 'Custom Private Sub
'    LogMessage "Starting Sub: CopyFavorites"
'    CopyFavorites 'Custom Private Sub
'    If bADAware = False Then
'        LogMessage "Starting Sub: CopyMyDocs"
'        CopyMyDocs 'Custom Private Sub
'    End If
'    If bRegIsNT = True Then
'        LogMessage "Starting Sub: CopyTemplates"
'        CopyTemplates 'Custom Private Sub
'    End If
'    LogMessage "Starting Sub: CopyNetHood"
'    CopyNetHood 'Custom Private Sub
'    If bRegIsNT Then
'        LogMessage "Starting Sub: CopyPBK"
'        CopyPBK
'    End If
'    LogMessage "Starting Sub: FileSearch"
'    FileSearch
'    If Err.Number <> 0 Then
'        MsgBox "(" & Err.Number & ") " & Err.Description
'        MsgBox "Backup process completed with errors. Please refer to log file for details.", , "Backup Complete"
'        objProcLog.WriteBlankLines 1
'        LogMessage "Backup process completed with the following return error: (" & Err.Number & ") " & Err.Description
'    Else
'        MsgBox "Backup process completed successfully.", , "Backup Complete"
'        LogMessage "Backup process completed successfully."
'        LogMessage "End Log." & Time
'
'    End If
'End Sub

'Private Sub ProfileReplay()
'    'Author:         Andrew Rowlands (Empired Ltd)
'    'Date Created:   07/08/02
'    'Date Modified   02/09/02 - Dave Robinson (Empired)
'    '                Default values added for backup server and username
'    '
'    'Copyright info: This script is licensed to Rio Tinto Exploration for the duration of the Rio tinto XP desktop SOE project.
'    '            This script may not be copied or modified in part or as a whole without the express permission on its author.
'    'Modifications:  Requests for modification or functionality enhancements should be sent to the author.
'    'Modem of Trust: Users of this script are trusted to abide by the conditions of use outlined above. Use of this script for
'    '                purposes other than intended are not permitted. This script has been left unencrypted (in a trusting state)
'    '                for educational and cross-platform purposes only. Discovered breaches of this copywright will force the Author
'    '                to encrypt future scripts.
'    'Author Contact: The Author can be contacted via email on: andy@perthweb.net.au or mobile: +61 412 833400
'
'    Dim strDataSrc, fso, objShell, strBakFolder, strSrcPath, iButton, bOLKFilesReplay, objLogFile, strThis
'    Dim Environ, DefaultDestination, Username
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set objShell = CreateObject("Shell")
'
'    Set Environ = objShell.Environment("Process")
'
'    strThis = ScriptFullName
'    If fso.FileExists(strThis & "\..\Reg.exe") = False Then
'        MsgBox "Reg.exe Not Found!. This script requires Reg.exe."
'        End
'    End If
'
'    MsgBox "This script will replay a users profile to this computer. This process may take" & _
'           " a long time." & vbCrLf & "On Success, a 'Replay Completed.'" & _
'           " dialog box will be displayed.", , "Profile Replay Tool"
'
'    DefaultDestination = "\\expperfile02\SOEBackup"
'
'    strDataSrc = InputBox("Please enter the path to the Captured User profile(s) location." & vbCrLf & vbCrLf & _
'                         "For example, if the Users profile was backed up to a share called 'profileBak$', " & _
'                         "you would enter the path: <\\Server\profileBak$>.", "Captured Profiles Store", DefaultDestination)
'
'    If Right(strDataSrc, 1) = "\" Then
'        strDataSrc = Left(strDataSrc, Len(strDataSrc) - 1)
'    End If
'
'    Username = Environ("Username")
'
'    If fso.FolderExists(strDataSrc) Then
'        strBakFolder = InputBox("Please enter the name of the folder used to contain the " & _
'                        "Users profile backup." & vbCrLf & vbCrLf & "For example, if the folder " & _
'                        "used to strore the users profile backup was called 'Andrew.Rowlands' then " & _
'                        "you would enter the following: <Andrew.Rowlands>.", "Enter Username Here", Username)
'        If fso.FolderExists(strDataSrc & "\" & strBakFolder) = False Then
'            MsgBox "The folder name entered could not be found. Please ensure you typed it correctly or that the folder " & _
'                   "exists.", , "User Profile Not Found!"
'            End
'        End If
'    Else
'        MsgBox "The path entered could not be found. Please ensure you typed it correctly or that the server is " & _
'                "avalable.", , "Profiles Store Not Found!"
'        End
'    End If
'
'    iButton = MsgBox("You have entered the following script parameters:" & vbCrLf & vbCrLf & _
'                "Captured Profiles Store: " & strDataSrc & vbCrLf & "Profile Backup Folder Name: " & strBakFolder & _
'                vbCrLf & vbCrLf & "Are These Correct? Click 'Yes' to continue, 'No' to abort.", 36, "Summary")
'    If iButton = 7 Then
'        End
'    End If
'
'    strSrcPath = strDataSrc & "\" & strBakFolder
'    Set objLogFile = fso.OpenTextFile(strSrcPath & "\DataReplay.log", 2, True)
'    LogMessage "Log Start " & Date & " " & Time()
'    objLogFile.WriteBlankLines 1
'    LogMessage "Profile Directory Source Location = " & strSrcPath
'    objLogFile.WriteBlankLines 1
'
'    LogMessage "Starting Sub: ReplayPreSOEData"
'    ReplayPreSOEData
'    LogMessage "Starting Sub: ReplayMyDocs"
'    ReplayMyDocs
'    LogMessage "Starting Sub: ReplayFavorites"
'    ReplayFavorites
'    LogMessage "Starting Sub: ReplayDesktop"
'    ReplayDesktop
'    LogMessage "Starting Sub: ReplayNetHood"
'    ReplayNetHood
'    LogMessage "Starting Sub: ReplayTemplates"
'    ReplayTemplates
'    LogMessage "Starting Sub: ReplayOLKFiles"
'    ReplayOLKFiles
'    LogMessage "Starting Sub: ReplayOLKFavFiles"
'    ReplayOLKFavFiles
'            'logmessage "Starting Sub: ReplayAppData"
'            'ReplayAppData
'    LogMessage "Starting Sub: UpdateOLKReg"
'    UpdateOLKReg
'    LogMessage "End Log. " & Time
'    objLogFile.Close
'    MsgBox "Profile Replay Completed!. To ensure success, please check the Replay log located at:" & vbCrLf & strSrcPath & "\", , "Profile Replay Tool"
'End Sub
'________________________________________________________________________________________________________________________
'Private Sub ReplayPreSOEData()
'    Dim strMyDocsPath
'
'    On Error Resume Next
'    If fso.FolderExists(strSrcPath & "\" & "PreSOEData") Then
'        LogMessage "Found Source 'PreSOEData' Folder."
'        strMyDocsPath = objShell.SpecialFolders("MyDocuments")
'        If fso.FolderExists(strMyDocsPath) Then
'            LogMessage "Found Destination 'My Documents' Folder at: " & strMyDocsPath
'            fso.CopyFolder strSrcPath & "\" & "PreSOEData", strMyDocsPath
'        Else
'            MsgBox "Could not find the local 'My Documents' Folder. This process will be skipped", 5, "Error!"
'            LogMessage "Could not find the local 'My Documents' Folder. Profile Replay Terminated..."
'            LogMessage "End Log. " & Time
'            End
'        End If
'    Else
'        MsgBox "Could not replay PreSOEData as no PreSOEData data exists.", 3, "PreSOEData Not Found", 64
'        LogMessage "*** Could not replay PreSOEData as no PreSOEData Exists."
'    End If
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: ReplayPreSOEData. This process will terminate", , "Error!"
'        LogMessage "An Error occurred in Sub: ReplayPreSOEData. Profile Replay Terminated..."
'        LogMessage "End Log. " & Time
'        End
'    End If
'End Sub
    
'________________________________________________________________________________________________________________________
'Private Sub ReplayMyDocs()
'    Dim strMyDocsPath, objMDBakFolder, colMDBakSubFolders, objFolder, objFile
'
'    On Error Resume Next
'    If fso.FolderExists(strSrcPath & "\" & "My Documents") Then
'        strMyDocsPath = objShell.SpecialFolders("MyDocuments")
'        LogMessage "Found Source 'My Documents' Folder"
'        If fso.FolderExists(strMyDocsPath) Then
'            LogMessage "Found Destination 'My Documents' Folder at: " & strMyDocsPath
'            Set objMDBakFolder = fso.GetFolder(strSrcPath & "\" & "My Documents")
'            For Each objFolder In objMDBakFolder.SubFolders
'                If fso.FolderExists(strMyDocsPath & "\" & objFolder.Name) Then
'                    fso.DeleteFolder strMyDocsPath & "\" & objFolder.Name, True
'                End If
'                fso.CopyFolder objFolder.Path, strMyDocsPath & "\", True
'            Next
'            For Each objFile In objMDBakFolder.Files
'                If fso.FileExists(strMyDocsPath & "\" & objFile.Name) Then
'                    fso.DeleteFile strMyDocsPath & "\" & objFile.Name, True
'                End If
'                fso.CopyFile objFile.Path, strMyDocsPath & "\", True
'            Next
'        Else
'            MsgBox "Could not find the local 'My Documents' Folder. This process will now terminate", 5, "Error!"
'            LogMessage "Could not find the local 'My Documents' Folder. Profile Replay Terminated..."
'            LogMessage "End Log. " & Time
'            End
'        End If
'    Else
'        MsgBox "Could not replay 'My Documents' as no 'My Documents' data exists.", 3, "My Documents Not Found", 64
'        LogMessage "*** Could not replay 'My Documents' as no 'My Documents' data Exists."
'    End If
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: ReplayMyDocs.", , "Error!"
'        LogMessage "An Error occurred in Sub: ReplayMyDocs...Continuing if possible."
'        LogMessage "" & Err.Number & " " & Err.Description
'    '   logmessage "End Log. " & Time
'    '   End
'    End If
'End Sub
''________________________________________________________________________________________________________________________
'Private Sub ReplayFavorites()
'    Dim strFavoritesPath, objFavBakFolder, objFolder, objFile
'
'    On Error Resume Next
'    If fso.FolderExists(strSrcPath & "\" & "Favorites") Then
'        strFavoritesPath = objShell.SpecialFolders("Favorites")
'        LogMessage "Found Source 'Favorites' Folder"
'        If fso.FolderExists(strFavoritesPath) Then
'            LogMessage "Found Destination 'Favorites' Folder at: " & strFavoritesPath
'            Set objFavBakFolder = fso.GetFolder(strSrcPath & "\" & "Favorites")
'            For Each objFolder In objFavBakFolder.SubFolders
'                If fso.FolderExists(strFavoritesPath & "\" & objFolder.Name) Then
'                    fso.DeleteFolder strFavoritesPath & "\" & objFolder.Name
'                End If
'                fso.CopyFolder objFolder.Path, strFavoritesPath & "\", True
'            Next
'            For Each objFile In objFavBakFolder.Files
'                If fso.FileExists(strFavoritesPath & "\" & objFile.Name) Then
'                    fso.DeleteFile strFavoritesPath & "\" & objFile.Name, True
'                End If
'                fso.CopyFile objFile.Path, strFavoritesPath & "\", True
'            Next
'        Else
'            MsgBox "Could not find the local 'Favorites' Folder. This process will be skipped", 5, "Error!"
'            LogMessage "Could not find the local 'Favorites' Folder. Profile Replay Terminated..."
'            LogMessage "End Log. " & Time
'            End
'        End If
'    Else
'        MsgBox "Could not replay 'Favorites' as no 'Favorites' data exists.", 3, "Favorites Not Found", 64
'        LogMessage "*** Could not replay 'Favorites' as no 'Favorites' data Exists."
'    End If
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: ReplayFavorites.", , "Error!"
'        LogMessage "An Error occurred in Sub: ReplayFavorites...Continuing if possible."
'        LogMessage "" & Err.Number & " " & Err.Description
'    '   logmessage "End Log. " & Time
'    '   End
'    End If
'End Sub
''________________________________________________________________________________________________________________________
'Private Sub ReplayDesktop()
'    Dim strDesktopPath, objDTBakFolder, objFolder, objFile
'
'    On Error Resume Next
'    If fso.FolderExists(strSrcPath & "\Desktop") Then
'        strDesktopPath = objShell.SpecialFolders("Desktop")
'        LogMessage "Found Source 'Desktop' Folder"
'        If fso.FolderExists(strDesktopPath) Then
'            LogMessage "Found Destination 'Desktop' Folder at: " & strDesktopPath
'            Set objDTBakFolder = fso.GetFolder(strSrcPath & "\Desktop")
'            For Each objFolder In objDTBakFolder.SubFolders
'                If fso.FolderExists(strDesktopPath & "\" & objFolder.Name) Then
'                    fso.DeleteFolder strDesktopPath & "\" & objFolder.Name
'                End If
'                fso.CopyFolder objFolder.Path, strDesktopPath & "\", True
'            Next
'            For Each objFile In objDTBakFolder.Files
'                If fso.FileExists(strDesktopPath & "\" & objFile.Name) Then
'                    fso.DeleteFile strDesktopPath & "\" & objFile.Name, True
'                End If
'                fso.CopyFile objFile.Path, strDesktopPath & "\", True
'            Next
'        Else
'            MsgBox "Could not find the local 'Desktop' Folder. This process will be skipped", 5, "Error!"
'            LogMessage "Could not find the local 'Desktop' Folder. Profile Replay Terminated..."
'            LogMessage "End Log. " & Time
'            End
'        End If
'    Else
'        MsgBox "Could not replay 'Desktop' as no 'Desktop' data exists.", 3, "Desktop Not Found", 64
'        LogMessage "*** Could not replay 'Desktop' as no 'Desktop' data Exists."
'    End If
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: ReplayDesktop", , "Error!"
'        LogMessage "An Error occurred in Sub: ReplayDesktop...Continuing if possible."
'        LogMessage "" & Err.Number & " " & Err.Description
'    '   logmessage "End Log. " & Time
'    '   End
'    End If
'End Sub
''________________________________________________________________________________________________________________________
'Private Sub ReplayNetHood()
'    Dim strNetHoodPath, objNHBakFolder, objFolder, objFile
'
'    On Error Resume Next
'    If fso.FolderExists(strSrcPath & "\NetHood") Then
'        strNetHoodPath = objShell.SpecialFolders("NetHood")
'        LogMessage "Found Source 'NetHood' Folder"
'        If fso.FolderExists(strNetHoodPath) Then
'            LogMessage "Found Destination 'NetHood' Folder at: " & strNetHoodPath
'            Set objNHBakFolder = fso.GetFolder(strSrcPath & "\NetHood")
'            For Each objFolder In objNHBakFolder.SubFolders
'                If fso.FolderExists(strNetHoodPath & "\" & objFolder.Name) Then
'                    fso.DeleteFolder strNetHoodPath & "\" & objFolder.Name
'                End If
'                fso.CopyFolder objFolder.Path, strNetHoodPath & "\", True
'            Next
'            For Each objFile In objNHBakFolder.Files
'                If fso.FileExists(strNetHoodPath & "\" & objFile.Name) = False Then
'                    fso.CopyFile objFile.Path, strNetHoodPath & "\", True
'                End If
'
'            Next
'        Else
'            MsgBox "Could not find the local 'NetHood' Folder. This process will be skipped", 5, "Warning!"
'            LogMessage "Could not find the local 'NetHood' Folder. Profile Replay Terminated..."
'            LogMessage "End Log. " & Time
'            End
'        End If
'    Else
'        MsgBox "Could not replay 'NetHood' as no 'NetHood' data exists.", 3, "NetHood Not Found", 64
'        LogMessage "*** Could not replay 'NetHood' as no 'NetHood' data Exists."
'    End If
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: ReplayNetHood.", , "Error!"
'        LogMessage "An Error occurred in Sub: ReplayNetHood...Continuing if possible."
'        LogMessage "" & Err.Number & " " & Err.Description
'    '   logmessage "End Log. " & Time
'    '   End
'    End If
'End Sub
''________________________________________________________________________________________________________________________
'Private Sub ReplayTemplates()
'    Dim strTemplatesPath, objTemplBakFolder, objFolder, objFile, blah
'    On Error Resume Next
'    If fso.FolderExists(strSrcPath & "\Templates") Then
'        strTemplatesPath = objShell.SpecialFolders("Templates")
'        LogMessage "Found Source 'Templates' Folder"
'        If fso.FolderExists(strTemplatesPath) Then
'            LogMessage "Found Destination 'Templates' Folder at: " & strTemplatesPath
'            Set objTemplBakFolder = fso.GetFolder(strSrcPath & "\" & "Templates")
'            For Each objFolder In objTemplBakFolder.SubFolders
'                If fso.FolderExists(strTemplatesPath & "\" & objFolder.Name) Then
'                    fso.DeleteFolder strTemplatesPath & "\" & objFolder.Name
'                End If
'                fso.CopyFolder objFolder.Path, strTemplatesPath & "\", True
'            Next
'            For Each objFile In objTemplBakFolder.Files
'                If fso.FileExists(strTemplatesPath & "\" & objFile.Name) Then
'                    fso.DeleteFile strTemplatesPath & "\" & objFile.Name, True
'                End If
'                fso.CopyFile objFile.Path, strTemplatesPath & "\", True
'            Next
'        Else
'            MsgBox "Could not find the local 'Templates' Folder. This process will be skipped", 5, "Warning!"
'            LogMessage "Could not find the local 'Templates' Folder. Profile Replay Terminated..."
'            LogMessage "End Log. " & Time
'            End
'        End If
'    Else
'        MsgBox "Could not replay 'Templates' as no 'Templates' data exists.", 3, "Templates Not Found", 64
'        LogMessage "*** Could not replay 'Templates' as no 'Templates' data Exists."
'
'    End If
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: ReplayTemplates.", , "Error!"
'        LogMessage "An Error occurred in Sub: ReplayTemplates...Continuing if possible."
'        LogMessage "" & Err.Number & " " & Err.Description
'    '   logmessage "End Log. " & Time
'    '   End
'    End If
'End Sub
''________________________________________________________________________________________________________________________
'Private Sub ReplayOLKFiles()
'    Dim strProfilePath, strAppDataPath
'
'    On Error Resume Next
'    If fso.FolderExists(strSrcPath & "\OutlookFiles") Then
'        strProfilePath = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
'        LogMessage "Found Source 'OutlookFiles' Folder"
'        If fso.FolderExists(strProfilePath & "\Local Settings" & "\Application Data" & "\Microsoft" & "\Outlook") Then
'            LogMessage "Found Destination 'Outlook' Folder at: " & strProfilePath & "\Local Settings" & "\Application Data" & "\Microsoft" & "\Outlook"
'            fso.CopyFile strSrcPath & "\OutlookFiles" & "\*.*", strProfilePath & "\Local Settings" & "\Application Data" & "\Microsoft" & "\Outlook\", True
'        Else
'            MsgBox "Outlook Folder does not exist in Profile." & vbCrLf & _
'            "Creating Folder: " & strProfilePath & "\Local Settings" & "\Application Data" & "\Microsoft" & "\Outlook", 2, "Information"
'            MkDir (strProfilePath & "\Local Settings" & "\Application Data" & "\Microsoft" & "\Outlook")
'            fso.CopyFile strSrcPath & "\OutlookFiles" & "\*.*", strProfilePath & "\Local Settings" & "\Application Data" & "\Microsoft" & "\Outlook\", True
'            LogMessage "Could not find the local 'Outlook' Folder...This folder has been created."
'        End If
'    End If
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: ReplayOLKFiles.", , "Error!"
'        LogMessage "An Error occurred in Sub: ReplayOLKFiles...Continuing if possible."
'        LogMessage "" & Err.Number & " " & Err.Description
'    '   logmessage "End Log. " & Time
'    '   End
'    End If
'    bOLKFilesReplay = True
'End Sub
'
''_______________________________________________________________________________________________________________________
'Private Sub ReplayOLKFavFiles()
'    Dim strAppDataPath
'
'    On Error Resume Next
'    If fso.FolderExists(strSrcPath & "\Fav") Then
'        strAppDataPath = objShell.ExpandEnvironmentStrings("%APPDATA%")
'        If fso.FolderExists(strAppDataPath & "\Microsoft") = False Then
'            LogMessage "Creating Folder: " & strAppDataPath & "\Microsoft"
'            MkDir (strAppDataPathPath & "\Microsoft")
'            If fso.FolderExists(strAppDataPath & "\Microsoft\Outlook") = False Then
'                LogMessage "Creating Folder: " & strAppDataPath & "\Microsoft\Outlook"
'                MkDir (strAppDataPath & "\Microsoft\Outlook")
'            End If
'        ElseIf fso.FolderExists(strAppDataPath & "\Microsoft\Outlook") = False Then
'                LogMessage "Creating Folder: " & strAppDataPath & "\Microsoft\Outlook"
'                MkDir (strAppDataPath & "\Microsoft\Outlook")
'        End If
'
'        If Err.Number <> 0 Then
'            MsgBox Err.Number & " " & Err.Description
'            MsgBox "An Error occurred in Sub: ReplayOLKFavFiles.", , "Error!"
'            LogMessage "An Error occurred in Sub: ReplayOLKFavFiles...Continuing if possible."
'            LogMessage "" & Err.Number & " " & Err.Description
'    '       logmessage "End Log. " & Time
'    '       End
'        End If
'
'        If fso.FolderExists(strAppDataPath & "\Microsoft\Outlook") Then
'            LogMessage "Found Folder: " & strAppDataPath & "\Microsoft\Outlook"
'            fso.CopyFile strSrcPath & "\Fav\*.*", strAppDataPath & "\Microsoft\Outlook\", True
'            LogMessage "Copying Outlook Fav Files..."
'        End If
'    End If
'    If Err.Number <> 0 Then
'        Err.Clear
'    End If
'End Sub
''_______________________________________________________________________________________________________________________
'Private Sub ReplayAppData()
'    Dim strAppDataPath, objFolder, objFile, objADBakFolder
'
'    On Error Resume Next
'    If fso.FolderExists(strSrcPath & "\AppData") Then
'        strAppDataPath = objShell.ExpandEnvironmentStrings("%APPDATA%")
'        LogMessage "Found Source 'AppData' Folder"
'        If fso.FolderExists(strAppDataPath) Then
'            LogMessage "Found Destination 'Application Data' Folder at: " & strAppDataPath
'            Set objADBakFolder = fso.GetFolder(strSrcPath & "\AppData")
'            For Each objFolder In objADBakFolder.SubFolders
'                If fso.FolderExists(strAppDataPath & "\" & objFolder.Name) Then
'                    fso.DeleteFolder strAppDataPath & "\" & objFolder.Name
'                End If
'                fso.CopyFolder objFolder.Path, strAppDataPath & "\"
'            Next
'            For Each objFile In objADBakFolder.Files
'                If fso.FileExists(strAppDataPath & "\" & objFile.Name) Then
'                    fso.DeleteFile strAppDataPath & "\" & objFile.Name, True
'                End If
'                fso.CopyFiles objFile.Path, strAppDataPath & "\"
'            Next
'        Else
'            MsgBox "An Error occurred in Sub: ReplayAppData. This process will terminate" & vbCrLf & vbCrLf & "Problem: " & _
'            "'Application Data' folder does not exist on this client!", , "Error!"
'            LogMessage "Could not find the local 'Application Data' Folder. Profile Replay Terminated..."
'            LogMessage "End Log. " & Time
'            End
'        End If
'    Else
'        MsgBox "A Backup copy of the 'Application Data' folder does not exist." & vbCrLf & _
'            "This may be because the 'old client' had redirected folders (ie App Data and My Docs) or because the" & vbCrLf & _
'            "'old client' was Windows 9x." & vbCrLf & vbCrLf & "This process will be skipped.", 5, "Information"
'        LogMessage "*** Could not replay 'Application Data' as no 'AppData' data Exists."
'        LogMessage "*** This may be because the 'old' client was Windows 9x."
'        LogMessage "*** 'Application Data' Replay Skipped..."
'    End If
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: ReplayAppData.", , "Error!"
'        LogMessage "An Error occurred in Sub: ReplayTemplates...Continuing if possible."
'        LogMessage "" & Err.Number & " " & Err.Description
'        'logmessage "End Log. " & Time
'        'End
'    End If
'End Sub
''________________________________________________________________________________________________________________________
'Private Sub UpdateOLKReg()
'    Dim objFileStreamIn, strPutLine, objFileStreamOut, strProperty, strValuePath, iEquals
'    Dim objOLKBakFolder, objFile, i, bContinue, strGetLine, strProfilePath, strOLKFilePath, colOLKBakFiles
'
'    Const FORWRITING = 2
'    Const FORREADING = 1
'    Const FORAPPEND = 8
'
'    On Error Resume Next
'    If fso.FileExists(strSrcPath & "\OLKBak.reg") = False Then
'        MsgBox "OLKBak.reg File not found." & vbCrLf & _
'        "Updating And applying Outlook registry settings cannot continue."
'        LogMessage "OLKBak.reg File not found..."
'        LogMessage "Updating And applying Outlook registry settings Aborted."
'        Exit Sub
'    End If
'
'    Set objFileStreamOut = fso.OpenTextFile(strSrcPath & "\OLKNew.reg", FORWRITING, True, vbTrue)
'    Set objFileStreamIn = fso.OpenTextFile(strSrcPath & "\OLKBak.reg", FORREADING, False, vbTrue)
'    strProfilePath = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
'    strProfilePath = Replace(strProfilePath, "\", "\\")
'    strOLKFilePath = """" & strProfilePath & "\\Local Settings\\Application Data\\Microsoft\\Outlook"
'    LogMessage "Outlook Registry file path info will be modified to: " & strProfilePath & "\\Local Settings\\Application Data\\Microsoft\\Outlook"
'
'    If fso.FolderExists(strSrcPath & "\OutlookFiles") Then
'        Set objOLKBakFolder = fso.GetFolder(strSrcPath & "\OutlookFiles")
'        Set colOLKBakFiles = objOLKBakFolder.Files
'        LogMessage "Found Source 'OutlookFiles' Folder"
'
'        If colOLKBakFiles.Count < 1 Then
'            LogMessage "No Outlook files of type: [.ost] [.pst] [.pab] were found in Data Source 'OutlookFiles'"
'            bContinue = MsgBox("No Outlook files of type: [.ost] [.pst] [.pab] were found in this Users backup profile." & vbCrLf & _
'            "Do you want to replay this Users Outlook registry settings anyway? (Recommended in most circumstances)" & vbCrLf & vbCrLf & _
'            "Possible reasons why no files were found:" & vbCrLf & _
'            "                    - No .ost, .pab, .pst files existed on the 'old' client (Most likely)." & vbCrLf & _
'            "                    - The Users profile backup has been modified (Outlook files moved or removed)." & vbCrLf & _
'            "                    - The 'ProfileCapture' tool was unable to backup the Users .ost, .pst, .pab files (Unlikely).", , "Outlook Files Not Found!", 52)
'            If bContinue = 7 Then
'                LogMessage "Outlook Registry Update Aborted (See previous log entry for details)."
'                Exit Sub
'            ElseIf bContinue = 6 Then
'            LogMessage "Outlook Registry Update Aborted. Continuing with Registry data import..."
'                Do While objFileStreamIn.AtEndOfStream = False
'                    strGetLine = objFileStreamIn.ReadLine
'                    strPutLine = strGetLine
'                    objFileStreamOut.WriteLine strPutLine
'                Loop
'                objFileStreamIn.Close
'                objFileStreamOut.Close
'                LogMessage "Starting Sub: ApplyOLKReg"
'                ApplyOLKReg
'            End If
'            Exit Sub
'        End If
'
'        i = 1
'        strGetLine = objFileStreamIn.ReadLine
'        Do While objFileStreamIn.AtEndOfStream = False
'            For Each objFile In colOLKBakFiles
'                If LCase(Right(strGetLine, Len(objFile.Name) + 1)) = LCase(objFile.Name) & """" Then
'                    iEquals = InStr(strGetLine, "=")
'                    strProperty = Left(strGetLine, iEquals)
'                    strGetLine = strProperty & strOLKFilePath & "\\" & objFile.Name & """"
'                End If
'            Next
'            strPutLine = strGetLine
'            objFileStreamOut.WriteLine strPutLine
'            strGetLine = objFileStreamIn.ReadLine
'        Loop
'        objFileStreamIn.Close
'        objFileStreamOut.Close
'    Else
'        MsgBox "OutlookFiles directory does not exist!"
'        LogMessage "Could not find Source 'OutlookFiles' Folder"
'        LogMessage "Outlook Registry Update Aborted. Continuing with Registry data import..."
'    End If
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: UpdateOLKReg.", , "Error!"
'        LogMessage "An Error occurred In Sub: UpdateOLKReg...Continuing if Possible."
'        LogMessage "" & Err.Number & " " & Err.Description
'    '   End
'    End If
'    LogMessage "Starting Sub: ApplyOLKReg"
'    ApplyOLKReg
'End Sub
'
''________________________________________________________________________________________________________________________
'Private Sub ApplyOLKReg()
'    On Error Resume Next
'
'    If fso.FileExists(strSrcPath & "\OLKNew.reg") Then
'        Shell "Regedit.exe /e """ & strSrcPath & "\OLKOld.reg""" & " " & """HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles""", 0, True
'        Shell "Reg.exe DELETE ""HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"" /f", 0, True
'        Shell "Regedit.exe /s """ & strSrcPath & "\OLKNew.reg""", 0, True
'    Else
'        MsgBox "Unable to apply Outlook registry settings. The file OLKNew.reg does not exist!"
'        LogMessage "Unable to apply Outlook registry settings. The file OLKNew.reg does not exist..."
'    End If
'
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & " " & Err.Description
'        MsgBox "An Error occurred in Sub: ApplyOLKReg.", , "Error!"
'        LogMessage "An Error occurred in Sub: ApplyOLKReg...Continuing if possible."
'        LogMessage "" & Err.Number & " " & Err.Description
'    '   logmessage "End Log. " & Time
'    '   End
'    End If
'End Sub
                                                      

'________________________________________________________________________________________________________________________
Sub CopyBackupDir(ByVal DestinationPath As String)
    'Search for preSOEData folder in the root of each fixed drive.
    Dim fso, objFile, colDrives, objDrive, strDrive
    Dim objProcLog
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objProcLog = fso.OpenTextFile(DestinationPath & "\DataCapture.log", FOR_APPENDING, True)
    
    On Error Resume Next
    
    LogMessage "Searching for PreSOEData folder in the root of local hard drives."
    Set colDrives = fso.Drives
    For Each objDrive In colDrives
        If objDrive.DriveType = 2 And objDrive.IsReady Then
            strDrive = objDrive.DriveLetter
            If fso.FolderExists(strDrive & ":" & "\" & FIND_FOLDER) Then
                LogMessage "Found PreSOEData Folder"
                fso.CopyFolder strDrive & ":" & "\" & FIND_FOLDER, DestinationPath & "\" & FIND_FOLDER, True
                If Err.Number <> 0 Then
                    LogMessage "Error 2 Occurred in Sub: CopyBackupDir (" & Err.Number & ") " & Err.Description
                    End
                End If
            End If
        End If
    Next
    If Err.Number <> 0 Then
        LogMessage "An Error Occurred in Sub: CopyBackupDir (" & Err.Number & ") " & Err.Description
        End
    End If
End Sub
'________________________________________________________________________________________________________________________
Private Sub GetOutlookSettings()
    Dim strOS
    
    On Error Resume Next
        
    strOS = Environ("OS")
    If UCase(strOS) = UCase("Windows_NT") Then
        bRegIsNT = True
        LogMessage "OS is type NT"
    Else
        bRegIsNT = False
        LogMessage "OS is type 9x"
    End If
    LogMessage "TODO:  Modify GetOutlookSettings to remove dependency on Regedit.exe"
    If bRegIsNT Then
        Shell "Regedit.exe /e """ & DefaultDestination & "\OLKBak.reg""" & " " & """HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"""
        Wait 10
        If Dir(DefaultDestination & "\OLKBak.reg") = "" Then
            LogMessage "No Personal Outlook settings found. Outlook Skipped. Continuing Backup Process...."
        Else
            LogMessage "Windows NT/2000/XP Outlook Settings Captured"
        End If
    Else
        Shell "Regedit.exe /e """ & DefaultDestination & "\OLKTemp.reg""" & " " & """HKEY_CURRENT_USER\Software\Microsoft\Windows Messaging Subsystem\Profiles"""
        Wait 10
        If Dir(DefaultDestination & "\OLKTemp.reg") = "" Then
            LogMessage "No Personal Outlook settings found. Continuing Backup Process...."
        Else
            LogMessage "Windows 9x Outlook Settings Captured"
            LogMessage "Starting Sub: ConvertOLKTemp"
            'ConvertOLKTemp 'Custom Private Sub
        End If
    End If
    If Err.Number <> 0 Then
        LogMessage "An Error Occurred in Sub: GetOutlookSettings (" & Err.Number & ") " & Err.Description
        End
    End If
End Sub
'________________________________________________________________________________________________________________________
'Private Sub ConvertOLKTemp()
'    Dim strFirstLine, objNewFile, strNewFileLine
'
'    Const WIN9X_MESSAGING = "[HKEY_CURRENT_USER\Software\Microsoft\Windows Messaging Subsystem\Profiles"
'    Const WINNT_MESSAGING = "[HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
'
'    Set objFile = fso.OpenTextFile(DefaultDestination & "\OLKTemp.reg", FOR_READING, True)
'    strFirstLine = Trim(objFile.ReadLine)
'    If strFirstLine = "" Then
'        LogMessage "(OLKTemp.reg) file not authentic. Outlook Window 9x --> NT Settings conversion cannot not continue."
'        objFile.Close
'        Exit Sub
'    Else
'        objNewFile = fso.OpentTextFile(DefaultDestination & "\OLKBak.reg", FOR_WRITING, True)
'        objNewFile.WriteLine ("Windows Registry Editor Version 5.00")
'        Do While Not objFile.AtEndOfStream
'            strFileLine = objFile.ReadLine
'            If InStr(strFileLine, WIN9X_MESSAGING) = 1 Then
'                strFileLine = Replace(strFileLine, WIN9X_MESSAGING, WINNT_MESSAGING)
'                strNewFileLine = strFileLine
'                objNewFile.WriteLine (strNewFileLine)
'            Else
'                strNewFileLine = strFileLine
'                objNewFile.WriteLine (strNewFileLine)
'            End If
'        Loop
'    End If
'    objFile.Close
'    objNewFile.Close
'    If fso.FileExists(DefaultDestination & "\OLKBak.reg") And fso.FileExists(DefaultDestination & "\OLKTemp.reg") Then
'        fso.DeleteFile DefaultDestination & "\OLKTemp.reg", True
'    End If
'    LogMessage "OLKTemp.reg --> OLKBak.reg Conversion Complete"
'    objFile = Nothing
'End Sub
'________________________________________________________________________________________________________________________
Private Sub BackupUserOLKFiles()
    Dim intEqualsLoc, strPSTProperty, strProfilePath, objFavFolder, objFavFile, objFile
    Dim strFileLine, strOldOSTLoc, strOldPSTLoc, strOldPABLoc, FilenameToCopy As String
    
    MkDir (DefaultDestination & "\OutlookFiles")
    If Dir(DefaultDestination & "\OLKBak.reg") <> "" Then
        Set objFile = fso.OpenTextFile(DefaultDestination & "\OLKBak.reg", FOR_READING, False, -1)
        'Search for .ost, .pst, .pab file locations
        Do While Not objFile.AtEndOfStream
            strFileLine = objFile.ReadLine
            'Find / backup .ost File
            If InStr(LCase(strFileLine), ".ost") <> 0 Then
                intEqualsLoc = InStr(strFileLine, "=")
                strOldOSTLoc = Right(strFileLine, (Len(strFileLine) - intEqualsLoc))
                strOldOSTLoc = Replace(strOldOSTLoc, "\\", "\")
                strOldOSTLoc = Replace(strOldOSTLoc, """", "")
                If fso.FileExists(strOldOSTLoc) Then
                    LogMessage "Found .ost file at location:" & strOldOSTLoc & " File will NOT be Copied...."
                    'fso.CopyFile strOldOSTLoc,DefaultDestination & "\OutlookFiles\",True
                Else
                    LogMessage "Failed to find " & strOldOSTLoc & " .ost file In the location stored In the Outlook registry file. Continuing...."
                End If
            End If
                
            'Find / backup .pst files
            If InStr(LCase(strFileLine), ".pst") <> 0 Then
                intEqualsLoc = InStr(strFileLine, "=")
                strOldPSTLoc = Right(strFileLine, (Len(strFileLine) - intEqualsLoc))
                strOldPSTLoc = Replace(strOldPSTLoc, "\\", "\")
                strOldPSTLoc = Replace(strOldPSTLoc, """", "")
                If Dir(strOldPSTLoc) <> "" Then
                    LogMessage "Found .pst file at location:" & strOldPSTLoc & " Copying File...."
                    FilenameToCopy = GetFilenamefromFullPath(strOldPSTLoc)
                    FileCopy strOldPSTLoc, DefaultDestination & "\OutlookFiles\" & FilenameToCopy
                Else
                    LogMessage "Failed to find " & strOldPSTLoc & " .pst file In the location stored In the Outlook registry file. Continuing...."
                End If
            End If
            'Find / backup .pab files
            If InStr(LCase(strFileLine), ".pab") <> 0 Then
                intEqualsLoc = InStr(strFileLine, "=")
                strOldPABLoc = Right(strFileLine, (Len(strFileLine) - intEqualsLoc))
                strOldPABLoc = Replace(strOldPABLoc, "\\", "\")
                strOldPABLoc = Replace(strOldPABLoc, """", "")
                If fso.FileExists(strOldPABLoc) Then
                    LogMessage "Found .pab file at location:" & strOldPABLoc & " Copying File...."
                    FilenameToCopy = GetFilenamefromFullPath(strOldPABLoc)
                    FileCopy strOldPABLoc, DefaultDestination & "\OutlookFiles\" & FilenameToCopy
                Else
                    LogMessage "Failed to find " & strOldPABLoc & " .pab file In the location stored In the Outlook registry file. Continuing...."
                End If
            End If
        Loop
        strProfilePath = Environ("APPDATA")
        If Dir(strProfilePath & "\Microsoft\Outlook", vbDirectory) <> "" Then
            MkDir (DefaultDestination & "\Fav")
            Set objFavFolder = fso.GetFolder(strProfilePath & "\Microsoft\Outlook")
            For Each objFavFile In objFavFolder.Files
                LogMessage "Copying Outlook .fav file: " & objFavFile.Name
                FilenameToCopy = GetFilenamefromFullPath(objFavFile.Path)
                FileCopy objFavFile.Path, DefaultDestination & "\Fav\" & FilenameToCopy
            Next
        End If
    Else
         LogMessage "Outlook Registry Settings File Not Found. Skipping Outlook Direct File backup method...."
        Exit Sub
    End If
    objFile.Close
End Sub
'________________________________________________________________________________________________________________________
 Private Sub CopyDesktopData()
    Dim strDesktopPath, objDesktop, colDeskFiles, objDeskFile, colDTSubFolders, objDTSubFolder
    
    On Error Resume Next
    
    MkDir (DefaultDestination & "\Desktop")
    strDesktopPath = GetFolderPath(CSIDL_DESKTOP)
    If strDesktopPath <> "" Then
        Set objDesktop = fso.GetFolder(strDesktopPath)
        Set colDTSubFolders = objDesktop.SubFolders
        For Each objDTSubFolder In colDTSubFolders
            fso.CopyFolder objDTSubFolder.Path, DefaultDestination & "\Desktop\"
        Next
        
        Set colDeskFiles = objDesktop.Files
        For Each objDeskFile In colDeskFiles
            If objDeskFile.Type <> "Shortcut" Then
                fso.CopyFile objDeskFile.Path, DefaultDestination & "\Desktop\"
            End If
        Next
    End If
    If Err.Number <> 0 Then
        LogMessage "An Error Occurred in Sub: CopyDesktopData (" & Err.Number & ") " & Err.Description
        End
    Else
        LogMessage "Desktop Data Copy completed Successfully"
    End If
End Sub
 '_______________________________________________________________________________________________________________________
 Private Sub CopyAppData()
     Dim strAppDataPath, objAppData, colADFolders, objADSubFolder, colADFiles, objADFile
        
     MkDir (DefaultDestination & "\AppData")
     strAppDataPath = GetFolderPath(CSIDL_APPDATA)
     If Dir(strAppDataPath, vbDirectory) <> "" Then
        Set objAppData = fso.GetFolder(strAppDataPath)
        Set colADFolders = objAppData.SubFolders
        For Each objADSubFolder In colADFolders
            fso.CopyFolder objADSubFolder.Path, DefaultDestination & "\AppData\"
        Next
        
        Set colADFiles = objAppData.Files
        For Each objADFile In colADFiles
            If LCase(objADFile.Name) <> "desktop.ini" Then
                fso.CopyFile objADFile.Path, DefaultDestination & "\AppData\"
            End If
        Next
    End If
    If Err.Number <> 0 Then
        LogMessage "An Error Occurred in Sub: CopyAppData (" & Err.Number & ") " & Err.Description
        End
    Else
        LogMessage "Application Data Copy completed Successfully."
    End If
End Sub
'________________________________________________________________________________________________________________________
Private Sub CopyFavorites()
    Dim strFavoritesPath, objFavorite, colFavFiles, objFavorites, colFavSubFolders, objFavSubFolder
    
    On Error Resume Next
    
    MkDir (DefaultDestination & "\Favorites")
    strFavoritesPath = GetFolderPath(CSIDL_FAVORITES)
    If strFavoritesPath <> "" Then
        Set objFavorites = fso.GetFolder(strFavoritesPath)
        Set colFavSubFolders = objFavorites.SubFolders
        For Each objFavSubFolder In colFavSubFolders
            fso.CopyFolder objFavSubFolder.Path, DefaultDestination & "\Favorites\"
        Next
        
        Set colFavFiles = objFavorites.Files
        For Each objFavorite In colFavFiles
            If LCase(objFavorite.Name) <> "desktop.ini" Then
                fso.CopyFile objFavorite.Path, DefaultDestination & "\Favorites\"
            End If
        Next
    End If
    If Err.Number <> 0 Then
        LogMessage "An Error Occurred in Sub: CopyFavorites (" & Err.Number & ") " & Err.Description
        End
    Else
        LogMessage "Favorites Data Copy completed Successfully."
    End If
End Sub
'________________________________________________________________________________________________________________________
Private Sub CopyMyDocs()
    Dim strMyDocsPath, objMyDocs, colMDFiles, objDoc, colMDSubFolders, objMDSubFolder
    
    On Error Resume Next
    
    MkDir (DefaultDestination & "\My Documents")
    strMyDocsPath = GetFolderPath(CSIDL_PERSONAL)
    If strMyDocsPath <> "" Then
        Set objMyDocs = fso.GetFolder(strMyDocsPath)
        Set colMDSubFolders = objMyDocs.SubFolders
        For Each objMDSubFolder In colMDSubFolders
            fso.CopyFolder objMDSubFolder.Path, DefaultDestination & "\My Documents\"
        Next
        
        Set colMDFiles = objMyDocs.Files
        For Each objDoc In colMDFiles
            If LCase(objDoc.Name) <> "desktop.ini" Then
                fso.CopyFile objDoc.Path, DefaultDestination & "\My Documents\"
            End If
        Next
    End If
    If Err.Number <> 0 Then
        LogMessage "An Error Occurred in Sub: CopyMyDocs (" & Err.Number & ") " & Err.Description
        End
    Else
        If strMyDocsPath <> "" Then
            LogMessage "My Documents Data Copy completed Successfully."
        Else
            LogMessage "My Documents Data Copy could not be completed.  The folder could not be determined."
        End If
    End If
End Sub
'________________________________________________________________________________________________________________________
Private Sub CopyTemplates()
    Dim strTemplatesPath, objTemplates, colTemplFiles, objTemplate, colTemplSubFolders, objTemplSubFolder
    
    On Error Resume Next
    
    strTemplatesPath = GetFolderPath(CSIDL_TEMPLATES)
    If strTemplatesPath <> "" Then
        MkDir (DefaultDestination & "\Templates")
        Set objTemplates = fso.GetFolder(strTemplatesPath)
        Set colTemplSubFolders = objTemplates.SubFolders
        For Each objTemplSubFolder In colTemplSubFolders
            fso.CopyFolder objTemplSubFolder.Path, DefaultDestination & "\Templates\"
        Next
        
        Set colTemplFiles = objTemplates.Files
        For Each objTemplate In colTemplFiles
            If LCase(objTemplate.Name) <> "desktop.ini" Then
                fso.CopyFile objTemplate.Path, DefaultDestination & "\Templates\"
            End If
        Next
    End If
    If Err.Number <> 0 Then
        LogMessage "An Error Occurred in Sub: CopyTemplates (" & Err.Number & ") " & Err.Description
        End
    Else
        If strTemplatesPath <> "" Then
            LogMessage "Templates Data Copy completed Successfully."
        Else
            LogMessage "Templates Data Copy could not be completed.  The folder could not be determined."
        End If
    End If
End Sub
'________________________________________________________________________________________________________________________

Private Sub CopyNetHood()
    Dim strNetHoodPath, objNetHood, colNetHoodSubFolders, objNetHoodSubFolder
    
    On Error Resume Next
    
    strNetHoodPath = GetFolderPath(CSIDL_NETHOOD)
    If strNetHoodPath <> "" Then
        MkDir (DefaultDestination & "\NetHood")
        Set objNetHood = fso.GetFolder(strNetHoodPath)
        Set colNetHoodSubFolders = objNetHood.SubFolders
        For Each objNetHoodSubFolder In colNetHoodSubFolders
            fso.CopyFolder objNetHoodSubFolder.Path, DefaultDestination & "\NetHood\"
        Next
        
    End If
    If Err.Number <> 0 Then
        LogMessage "An Error Occurred in Sub: CopyNetHood (" & Err.Number & ") " & Err.Description
    Else
        If strNetHoodPath <> "" Then
            LogMessage "NetHood Data Copy completed Successfully."
        Else
            LogMessage "NetHood Data Copy could not be completed.  The folder could not be determined."
        End If
    End If
End Sub
'________________________________________________________________________________________________________________________
Private Sub CopyPBK()
    Dim strAllUserProfile, strPBKPath
    
    On Error Resume Next
    
    strAllUserProfile = Environ("ALLUSERSPROFILE")
    strPBKPath = strAllUserProfile & "\" & "Application Data\Microsoft\Network\Connections\Pbk"
    If fso.FolderExists(strPBKPath) Then
        fso.CopyFolder strPBKPath, DefaultDestination & "\"
    End If
    If Err.Number <> 0 Then
        LogMessage "An Error Occurred in Sub: CopyPBK (" & Err.Number & ") " & Err.Description
    Else
        If strAllUserProfile <> "" Then
            LogMessage "RAS Phone Book Data Copy completed Successfully."
        Else
            LogMessage "RAS Phone Book Data Copy could not be completed.  The folder could not be determined."
        End If
    End If
End Sub
'________________________________________________________________________________________________________________________
Private Sub FileSearch()
    Dim objInputFile1, strInputFile1, strReadLine, colRootSubFolder, colFiles, objInputFile2, strInputFile2
    Dim strIncludeList, strExcludeFolders, arrFileType, arrExcludeFolders, objDrive, strDrive, objRootFolder
    Dim colDrives ' <---- global
    
    On Error Resume Next
    
    strInputFile1 = "FileSearch.ini"
    strInputFile2 = "ExcludeFolders.ini"
    
    If fso.FileExists(strInputFile1) = False Then
        LogMessage "No FileSearch input file found. Skipping extended file trawl of hard drive(s)."
        Exit Sub
    End If
        
    Set objInputFile1 = fso.OpenTextFile(strInputFile1)
    strReadLine = objInputFile1.ReadLine
    If Trim(strReadLine) <> "Signature=Grumble" Then
        LogMessage "FileSearch Input file found . . .But file not authentic."
        Exit Sub
    End If
        
    Do While strReadLine <> "[FileTypes]" And objInputFile1.AtEndOfStream = False
        strReadLine = objInputFile1.ReadLine
    Loop
    If objInputFile1.AtEndOfStream Then
        LogMessage "No File types section found in FileSearch file. File Search aborted."
        Exit Sub
    End If
    Do While objInputFile1.AtEndOfStream = False
        strReadLine = objInputFile1.ReadLine
        If Trim(strReadLine) <> "" Then
            strIncludeList = strIncludeList & strReadLine & ";"
        End If
    Loop
    objInputFile1.Close

    If Right(strIncludeList, 1) = ";" Then
        strIncludeList = Mid(strIncludeList, 1, Len(strIncludeList) - 1)
    End If
    arrFileType = Split(strIncludeList, ";")
    If UBound(arrFileType) > 0 Then
        MkDir (DefaultDestination & "\LostAndFound")
    Else
        Exit Sub
    End If
    
    '___
    
    If fso.FileExists(strInputFile2) = False Then
        LogMessage "No Folder Exclusion file found. Cannot perform file trawl of hard drive(s)."
        Exit Sub
    End If
    Set objInputFile2 = fso.OpenTextFile(strInputFile2)
    strReadLine = objInputFile2.ReadLine
    If Trim(strReadLine) <> "Signature=Grumble" Then
        LogMessage "Folder Exclusion file found . . .But file not authentic."
        Exit Sub
    End If
    Do While strReadLine <> "[ExcludeFolders]" And objInputFile2.AtEndOfStream = False
        strReadLine = objInputFile2.ReadLine
    Loop
    Do While objInputFile2.AtEndOfStream = False
        strReadLine = objInputFile2.ReadLine
        If Trim(strReadLine) <> "" Then
            strExcludeFolders = strExcludeFolders & strReadLine & ";"
        End If
    Loop
    objInputFile2.Close
    
    If Right(strExcludeFolders, 1) = ";" Then
        strExcludeFolders = Mid(strExcludeFolders, 1, Len(strExcludeFolders) - 1)
    End If
    arrExcludeFolders = Split(strExcludeFolders, ";")
            
    For Each objDrive In colDrives
        If objDrive.DriveType = 2 And objDrive.IsReady Then
            strDrive = objDrive.DriveLetter
            Set objRootFolder = fso.GetFolder(strDrive & ":\")
            'Set colFiles = objRootFolder.Files
            'Search (objRootFolder)
        End If
    Next
End Sub
'_________________________________________________________________________________________________________________________
'Private Sub Search(objRootFolder)
'    Dim objFile, objSubfolder, bIncludeFile, bDupeName, strFileNewName, intIncr, index, bExcludeFolder
'    index = 0
'    bExcludeFolder = False
'    Do While index <= UBound(arrExcludeFolders)
'        If InStr(LCase(objRootFolder.Path), "\" & LCase(arrExcludeFolders(index)) & "\") <> 0 Or LCase(Right(objRootFolder.Path, Len(arrExcludeFolders(index)))) = LCase(arrExcludeFolders(index)) Then
'            bExcludeFolder = True
'        End If
'        index = index + 1
'    Loop
'    If bExcludeFolder = False Then
'        For Each objFile In objRootFolder.Files
'            i = 0
'            bIncludeFile = False
'            Do While i <= UBound(arrFileType) And bIncludeFile = False
'                If Right(LCase(objFile.Name), 4) = LCase(arrFileType(i)) Then
'                    bIncludeFile = True
'                End If
'                i = i + 1
'            Loop
'            If bIncludeFile = True Then
'                intIncr = 1
'                bDupeName = False
'                strFileNewName = objFile.Name
'                Do While fso.FileExists(DefaultDestination & "\LostAndFound\" & strFileNewName) = True
'                    bDupeName = True
'                    intIncr = intIncr + 1
'                    strFileNewName = objFile.Name & intIncr
'                Loop
'                If bDupeName = True Then
'                    fso.CopyFile objFile.Path, DefaultDestination & "\LostAndFound\" & strFileNewName
'                    LogMessage "Copying file: " & objFile.Path & " --Renamed--> " & strFileNewName
'                Else
'                    fso.CopyFile objFile.Path, DefaultDestination & "\LostAndFound\"
'                    LogMessage "Copying file: " & objFile.Path
'                End If
'            End If
'            intIncr = 1
'        Next
'    End If
'
'    For Each objSubfolder In objRootFolder.SubFolders
'        Search objSubfolder
'    Next
'End Sub
'________________________________________________________________________________________________________________________

Private Sub chkADAware_Click()
    If chkADAware.Value = vbChecked Then
        bADAware = True
    Else
        bADAware = False
    End If
End Sub

Private Sub cmdBrowse_Click()
    DefaultDestination = GetFolderAPI("Select destination server and share")
    txtA1.Text = DefaultDestination
End Sub

'Private Sub cmdCapture_Click()
'    ProfileCapture
'End Sub
'
Private Sub cmdEnd_Click()
    End
End Sub

Private Sub cmdOne_Click()
    StepNo = StepNo - 1
    UpdateStep
End Sub

Private Sub cmdTwo_Click()
    StepNo = StepNo + 1
    UpdateStep
End Sub

Private Sub Form_Load()
    Set fso = CreateObject("Scripting.FileSystemObject")
    StepNo = 1
    cmdOne.Enabled = False
    UpdateStep
    'frmMain.txtA1.Text = "C:\Temp\SOE"
    frmMain.txtA1.Text = App.Path
End Sub

Sub UpdateStep()
    Dim i As Integer, strTempFoldername As String, d As String
    
    If optMode(0).Value = True Then ' Copy
        Select Case StepNo
            Case 1
                lblTitle.Caption = "Welcome"
                lblDescription.Caption = "This application will capture or replay the current user's settings and documents."
                lblNavigation.Caption = "Select 'Next' to start the wizard, or Exit to end."
                lblQ1.Visible = False
                txtA1.Visible = False
                lblQ2.Visible = False
                txtA2.Visible = False
                cmdBrowse.Visible = False
                cmdOne.Enabled = False
                chkADAware.Visible = False
                optMode(0).Visible = False
                optMode(1).Visible = False
            Case 2
                lblTitle.Caption = "Select mode"
                lblDescription.Caption = "Select Copy to save the user's settings, or Replay to retrieve the user's settings from a previously captured session."
                lblQ1.Caption = ""
                lblQ2.Caption = ""
                lblQ1.Visible = False
                txtA1.Visible = False
                lblQ2.Visible = False
                txtA2.Visible = False
                cmdBrowse.Visible = False
                cmdOne.Enabled = False
                chkADAware.Visible = False
                optMode(0).Visible = True
                optMode(1).Visible = True
                lblNavigation.Caption = "Select 'Next' when ready, or Exit to end."
                LogMessage "Determining mode."
            Case 3
                lblTitle.Caption = "Enter details"
                lblDescription.Caption = "Enter the destination server and share where the data is to be saved, and the current username."
                lblQ1.Caption = "Destination"
                lblQ2.Caption = "Username"
                If txtA2.Text = "" Then
                    txtA2.Text = Environ("Username")
                End If
                lblQ1.Visible = True
                txtA1.Visible = True
                lblQ2.Visible = True
                txtA2.Visible = True
                cmdBrowse.Visible = True
                cmdOne.Enabled = True
                chkADAware.Visible = True
                optMode(0).Visible = False
                optMode(1).Visible = False
                lblNavigation.Caption = "Select 'Next' when ready, or Exit to end."
                LogMessage "Collecting data"
            Case 4
                lblTitle.Caption = "Ready"
                lblDescription.Caption = "Click Next to start the profile backup."
                lblQ1.Caption = ""
                lblQ2.Caption = ""
                lblQ1.Visible = False
                txtA1.Visible = False
                lblQ2.Visible = False
                txtA2.Visible = False
                cmdBrowse.Visible = False
                chkADAware.Visible = False
                optMode(0).Visible = False
                optMode(1).Visible = False
                lblNavigation.Caption = "Select 'Next' when ready, or Exit to end."
                
                DefaultDestination = txtA1.Text
                Username = txtA2.Text
                LogMessage "Ready"
            Case 5
                strTempFoldername = DefaultDestination & "\" & Username
                d = Dir(strTempFoldername, vbDirectory)
                Do While d <> ""
                    i = i + 1
                    strTempFoldername = DefaultDestination & "\" & Username & Trim(Str(i))
                    d = Dir(strTempFoldername, vbDirectory)
                Loop
                If DefaultDestination & "\" & Username <> strTempFoldername Then
                    LogMessage DefaultDestination & "\" & Username & " already exists.  " & strTempFoldername & " will be used instead."
                End If
                MkDir strTempFoldername
                DefaultDestination = strTempFoldername
                LogMessage "Starting CopyBackupDir (" & DefaultDestination & ")"
                CopyBackupDir DefaultDestination
                LogMessage "Starting GetOutlookSettings"
                GetOutlookSettings
                LogMessage "Starting BackupUserOLKFiles"
                BackupUserOLKFiles
                LogMessage "Starting CopyDesktopData"
                CopyDesktopData
                LogMessage "Starting CopyFavorites"
                CopyFavorites
                If Not bADAware Then
                    LogMessage "Starting CopyMyDocs"
                    CopyMyDocs
                End If
                If bRegIsNT Then
                    LogMessage "Starting CopyTemplates"
                    CopyTemplates
                End If
                LogMessage "Starting CopyNetHood"
                CopyNetHood
                If bRegIsNT Then
                    LogMessage "Starting CopyPBK"
                    CopyPBK
                End If
                LogMessage "Starting FileSearch"
                FileSearch
                LogMessage "Complete."
                StepNo = StepNo + 1
                UpdateStep
            Case 6
                lblTitle.Caption = "Complete."
                lblDescription.Caption = "You have completed the profile backup process."
                lblQ1.Caption = ""
                lblQ2.Caption = ""
                If txtA2.Text = "" Then
                    txtA2.Text = ""
                End If
                lblQ1.Visible = False
                txtA1.Visible = False
                lblQ2.Visible = False
                txtA2.Visible = False
                cmdBrowse.Visible = False
                chkADAware.Visible = False
                lblNavigation.Caption = "Click Exit to end."
                
                LogMessage "Application finished."
        End Select
    Else
        lblTitle.Caption = "Sorry, this is not yet supported."
        lblDescription.Caption = "You have completed the profile backup process."
        lblQ1.Caption = ""
        txtA1.Text = ""
        lblQ2.Caption = ""
        lblQ1.Visible = False
        txtA1.Visible = False
        lblQ2.Visible = False
        txtA2.Visible = False
        cmdBrowse.Visible = False
        chkADAware.Visible = False
        optMode(0).Visible = False
        optMode(1).Visible = False
        lblNavigation.Caption = "Click Exit to end."
        LogMessage "Application finished."
        cmdOne.Enabled = False
        cmdTwo.Enabled = False
    End If
End Sub

Sub LogMessage(strMessage As String)
    frmMain.lstResults.AddItem FormatDateTime(Now, vbShortTime) & ": " & strMessage, 0
    frmMain.Refresh
End Sub

Private Function GetFolderAPI(ByVal strTitle As String, Optional flags As Long = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_NEWDIALOGSTYLE + BIF_EDITBOX) As String
    'Opens a Treeview control that displays the directories in a computer
    
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    szTitle = strTitle
    With tBrowseInfo
        .hwndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = flags
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        GetFolderAPI = sBuffer
    End If
End Function

Public Function GetFilenamefromFullPath(ByVal strFullPath As String) As String
    Dim i As Integer
    i = InStrRev(strFullPath, "\")
    GetFilenamefromFullPath = Right(strFullPath, Len(strFullPath) - i)
End Function

Function GetFolderPath(lFOLDER_ENUM As Long) As String
    Dim sPath As String
    Dim RetVal As Long
    
    ' Fill our string buffer
    sPath = String(MAX_PATH, 0)
    
    RetVal = SHGetFolderPath(0, lFOLDER_ENUM, 0, SHGFP_TYPE_CURRENT, sPath)
    
    Select Case RetVal
        Case S_OK
            ' We retrieved the folder successfully
            
            ' All C strings are null terminated
            ' So we need to return the string upto the first null character
            sPath = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
            GetFolderPath = sPath
        Case S_FALSE
            ' The CSIDL in nFolder is valid, but the folder does not exist.
            ' Use CSIDL_FLAG_CREATE to have it created automatically
            LogMessage "The folder does not exist."
        Case E_INVALIDARG
            ' nFolder is invalid
            LogMessage "An invalid folder ID was specified (" & lFOLDER_ENUM & ")."
        
    End Select

End Function

Sub Wait(dblSeconds As Double)
    Dim d As Date
    d = DateAdd("s", dblSeconds, Now)
    Do
        DoEvents
    Loop Until Now > d
End Sub
