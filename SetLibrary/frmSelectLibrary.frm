VERSION 5.00
Begin VB.Form frmSelectLibrary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select DRIMS Library"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmSelectLibrary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optLibrary 
      Caption         =   "CDC"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.OptionButton optLibrary 
      Caption         =   "DRIMS"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.OptionButton optLibrary 
      Caption         =   "Sunrise"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.OptionButton optLibrary 
      Caption         =   "Vincent Enfield"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton optLibrary 
      Caption         =   "Onshore Projects"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton optLibrary 
      Caption         =   "NONE (Disabled)"
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdUsage 
      Caption         =   "Usage..."
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      Caption         =   "Please select the DRIMS Library you wish to change to..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmSelectLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strProgramFilesDir As String ':( Missing Scope
Dim strProgDrive As String ':( Missing Scope
Dim PCDocsPath As String ':( Missing Scope
Dim bDebugMode As Boolean ':( Missing Scope

'---------------------------------------------------------------------------------------
' Procedure : cmdSet_Click
' DateTime  : 01-04-2003 10:43
' Author    : Dave Robinson
' Purpose   : Determine which Library is selected.  Kick off SetLibrary with appropriate option.
'
'  V    Date        Author          History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Private Sub cmdSet_Click()

  Dim intLoop As Integer

    For intLoop = 0 To 5
        If optLibrary(intLoop).Value = True Then ':( Remove Pleonasm
            If bDebugMode Then MsgBox "About to enter SetLibrary with : " & intLoop ':( Expand Structure
            SetLibrary intLoop
            Exit For '>---> Next
        End If
    Next intLoop

End Sub

Private Sub cmdUsage_Click()

  Dim strHelp As String

    strHelp = ""
    strHelp = strHelp & "SwapLibrary.exe <option> , where <option> is as follows:" & vbCrLf
    strHelp = strHelp & "                cdc : Set CDC as the default library" & vbCrLf
    strHelp = strHelp & "                drims : Set DRIMS as the default library" & vbCrLf
    strHelp = strHelp & "                sunrise : Set Sunrise as the default library" & vbCrLf
    strHelp = strHelp & "                vincentenfield : Set Vincent Enfield as the default library" & vbCrLf
    strHelp = strHelp & "                onshoreprojects : Set Onshore Projects as the default library" & vbCrLf
    strHelp = strHelp & "                disabled : Disable Docs Open Integration" & vbCrLf
    strHelp = strHelp & "                debug : GUI Mode, with debug messages" & vbCrLf

    MsgBox strHelp, vbOKOnly, "Usage"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : DisableDocsIntegration
' DateTime  : 01-04-2003 10:46
' Author    : Dave Robinson
' Purpose   : Turn DOCS Integration OFF
'
'  V    Date        Author          History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Private Sub DisableDocsIntegration(Optional bCDC As Boolean = False)

  Dim strUserDestinationDir As String
  Dim strWindowsDestinationDir As String
  Dim strWindowsSystemDestinationDir As String
  Dim oRegistry As REGTool5.Registry
  Dim strAcrobatInstallPath As String
  Dim strHint As String
  Dim isTerminalServer As Boolean

    On Error GoTo IntegrationOffError

    strHint = "Creating Registry object"
    Set oRegistry = CreateObject("REGTool5.Registry")

    strUserDestinationDir = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    If Right$(strUserDestinationDir, 1) = "\" Then
      Else 'NOT RIGHT$(STRUSERDESTINATIONDIR,...
        strUserDestinationDir = strUserDestinationDir & "\"
    End If

    strWindowsDestinationDir = strUserDestinationDir & "Windows\"
    strWindowsSystemDestinationDir = strWindowsDestinationDir & "System\"

    ' Is this a terminal Server session?
    If Not isTerminalServer Then
        isTerminalServer = (Environ$("CLIENTNAME") <> "" And LCase$(Environ$("CLIENTNAME")) <> "console") ' True = YES:  Terminal Server Session
    End If

    If bDebugMode Then MsgBox "isTerminalServer: " & isTerminalServer ':( Expand Structure

    If isTerminalServer Then
        strProgramFilesDir = "D:\Program Files\"
      Else 'ISTERMINALSERVER = FALSE/0
        strProgramFilesDir = "C:\Program Files\"
    End If

    If bDebugMode Then MsgBox "strProgramFilesDir: " & strProgramFilesDir ':( Expand Structure

    ' Get Acrobat install path
    strHint = "Retrieving (Default) from HKEY_CURRENT_USER\Software\Adobe\Acrobat Reader\5.0\InstallPath"
    oRegistry.GetKeyValue REGToolRootTypes.HKEY_CURRENT_USER, "Software\Adobe\Acrobat Reader\5.0\InstallPath", "", strAcrobatInstallPath

    ' Look for Office Directory, create as necessary
    If Dir(strUserDestinationDir & "Microsoft Office", vbDirectory) = "" Then
        strHint = "Creating " & strUserDestinationDir & "Microsoft Office directory"
        MkDir strUserDestinationDir & "Microsoft Office"
    End If

    ' Look for XLStart Directory, create as necessary
    If Dir(strUserDestinationDir & "Microsoft Office\XLStart", vbDirectory) = "" Then
        strHint = "Creating " & strUserDestinationDir & "Microsoft Office\XLStart directory"
        MkDir strUserDestinationDir & "Microsoft Office\XLStart"
    End If

    ' Look for Startup Directory, create as necessary
    If Dir(strUserDestinationDir & "Microsoft Office\Startup", vbDirectory) = "" Then
        strHint = "Creating " & strUserDestinationDir & "Microsoft Office\Startup directory"
        MkDir strUserDestinationDir & "Microsoft Office\Startup"
    End If

    ' Remove Excel/Word/Acrobat Startup file(s)
    If Dir(strUserDestinationDir & "Microsoft Office\Startup\WordInXP.dot") <> "" Then
        strHint = "Deleting " & strProgramFilesDir & "Microsoft Office\Startup\WordInXP.dot"
        Kill strUserDestinationDir & "Microsoft Office\Startup\WordInXP.dot"
    End If

    If Dir(strUserDestinationDir & "Microsoft Office\XLStart\DocsxlXP.xla") <> "" Then
        strHint = "Deleting " & strUserDestinationDir & "Microsoft Office\XLStart\DocsxlXP.xla"
        Kill strUserDestinationDir & "Microsoft Office\XLStart\DocsxlXP.xla"
    End If

    If Dir(strAcrobatInstallPath & "plug_ins\Acrodx32.api") <> "" Then
        strHint = "Deleting " & strAcrobatInstallPath & "plug_ins\Acrodx32.api"
        Kill strAcrobatInstallPath & "plug_ins\Acrodx32.api"
    End If

    If bCDC Then
        If Dir(PCDocsPath & "PCDOCSIni\pcdocscdc.ini") <> "" Then
            ' copy pcdocscdc.ini to DRIMS Directory as pcdocs.ini
            If Dir(PCDocsPath & "pcdocs.ini") <> "" Then
                strHint = "Deleting " & PCDocsPath & "pcdocs.ini"
                Kill PCDocsPath & "pcdocs.ini"
            End If
            strHint = "Copying " & PCDocsPath & "PCDOCSIni\pcdocscdc.ini to " & PCDocsPath & "pcdocs.ini"
            FileCopy PCDocsPath & "PCDOCSIni\pcdocscdc.ini", PCDocsPath & "pcdocs.ini"
        End If
      Else 'BCDC = FALSE/0
        If Dir(PCDocsPath & "PCDOCSIni\pcdocs_not_user.ini") <> "" Then
            ' copy pcdocs_not_user.ini to DRIMS Directory as pcdocs.ini
            If Dir(PCDocsPath & "pcdocs.ini") <> "" Then
                strHint = "Deleting " & PCDocsPath & "pcdocs.ini"
                Kill PCDocsPath & "pcdocs.ini"
            End If
            strHint = "Copying " & PCDocsPath & "PCDOCSIni\pcdocs_not_user.ini to " & PCDocsPath & "pcdocs.ini"
            FileCopy PCDocsPath & "PCDOCSIni\pcdocs_not_user.ini", PCDocsPath & "pcdocs.ini"
        End If
    End If

    If Dir(strUserDestinationDir & "Windows\odma32.dll") <> "" Then
        strHint = "Deleting " & strUserDestinationDir & "Windows\odma32.dll"
        Kill strUserDestinationDir & "Windows\odma32.dll"
    End If
    Set oRegistry = Nothing

    Exit Sub

IntegrationOffError:
    MsgBox "Unexpected error " & Err.Number & " (" & Err.Description & ") in DisableDocsIntegration." & vbCrLf & vbCrLf & "Hint:" & vbCrLf & strHint
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnableDocsIntegration
' DateTime  : 01-04-2003 10:46
' Author    : Dave Robinson
' Purpose   : Turn DOCS Integration ON
'
'  V    Date        Author          History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Private Sub EnableDocsIntegration(strIniFilename As String)

  Dim strUserDestinationDir As String
  Dim strWindowsDestinationDir As String
  Dim strWindowsSystemDestinationDir As String
  Dim strAcrobatInstallPath As String
  Dim oRegistry As REGTool5.Registry
  Dim strHint As String

    On Error GoTo IntegrationOnError

    strHint = "Creating Registry object"
    Set oRegistry = CreateObject("REGTool5.Registry")

    strUserDestinationDir = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    If Right$(strUserDestinationDir, 1) = "\" Then
      Else 'NOT RIGHT$(STRUSERDESTINATIONDIR,...
        strUserDestinationDir = strUserDestinationDir & "\"
    End If

    If bDebugMode Then MsgBox "strUserDestinationDir: " & strUserDestinationDir ':( Expand Structure

    If Dir(strUserDestinationDir & "Windows", vbDirectory) = "" Then
        strHint = "Creating " & strUserDestinationDir & "Windows directory"
        MkDir strUserDestinationDir & "Windows"
    End If

    strWindowsDestinationDir = strUserDestinationDir & "Windows\"

    If bDebugMode Then MsgBox "strWindowsDestinationDir: " & strWindowsDestinationDir ':( Expand Structure

    strWindowsSystemDestinationDir = strWindowsDestinationDir & "System\"

    If bDebugMode Then MsgBox "strWindowsSystemDestinationDir: " & strWindowsSystemDestinationDir ':( Expand Structure

    ' Get Acrobat install path
    strHint = "Retrieving (Default) from HKEY_CURRENT_USER\Software\Adobe\Acrobat Reader\5.0\InstallPath"
    oRegistry.GetKeyValue REGToolRootTypes.HKEY_CURRENT_USER, "Software\Adobe\Acrobat Reader\5.0\InstallPath", "", strAcrobatInstallPath

    If bDebugMode Then MsgBox "strAcrobatInstallPath: " & strAcrobatInstallPath ':( Expand Structure

    ' Copy ini files to H:\Windows path
    If Dir(strWindowsDestinationDir & "pcdocs.ini") <> "" Then
        strHint = "Deleting " & PCDocsPath & "pcdocs.ini"
        Kill strWindowsDestinationDir & "pcdocs.ini"
    End If
    strHint = "Copying " & PCDocsPath & "PCDOCSIni\" & strIniFilename & " to " & strWindowsDestinationDir & "pcdocs.ini"
    FileCopy PCDocsPath & "PCDOCSIni\" & strIniFilename, strWindowsDestinationDir & "pcdocs.ini"

    ' Look for Office Directory, create as necessary
    If Dir(strUserDestinationDir & "Microsoft Office", vbDirectory) = "" Then
        strHint = "Creating " & strUserDestinationDir & "Microsoft Office directory"
        MkDir strUserDestinationDir & "Microsoft Office"
    End If

    ' Look for XLStart Directory, create as necessary
    If Dir(strUserDestinationDir & "Microsoft Office\XLStart", vbDirectory) = "" Then
        strHint = "Creating " & strUserDestinationDir & "Microsoft Office\XLStart directory"
        MkDir strUserDestinationDir & "Microsoft Office\XLStart"
    End If

    ' Look for Startup Directory, create as necessary
    If Dir(strUserDestinationDir & "Microsoft Office\Startup", vbDirectory) = "" Then
        strHint = "Creating " & strUserDestinationDir & "Microsoft Office\Startup directory"
        MkDir strUserDestinationDir & "Microsoft Office\Startup"
    End If

    If Dir(strUserDestinationDir & "Microsoft Office\Startup\WordInXP.dot") = "" Then
        strHint = "Copying " & PCDocsPath & "Integration\WordInXP.dot to " & strUserDestinationDir & "Microsoft Office\Startup\WordInXP.dot"
        FileCopy PCDocsPath & "Integration\WordInXP.dot", strUserDestinationDir & "Microsoft Office\Startup\WordInXP.dot"
    End If

    If Dir(strUserDestinationDir & "Microsoft Office\XLStart\DocsxlXP.xla") = "" Then
        strHint = "Copying " & PCDocsPath & "Integration\DocsxlXP.xla to " & strUserDestinationDir & "Microsoft Office\XLStart\DocsxlXP.xla"
        FileCopy PCDocsPath & "Integration\DocsxlXP.xla", strUserDestinationDir & "Microsoft Office\XLStart\DocsxlXP.xla"
    End If

    If strAcrobatInstallPath <> "" Then
        If Dir(strAcrobatInstallPath & "plug_ins\Acrodx32.api") = "" Then
            strHint = "Copying " & PCDocsPath & "Integration\Acrodx32.api to " & strAcrobatInstallPath & "plug_ins\Acrodx32.api"
            FileCopy PCDocsPath & "Integration\Acrodx32.api", strAcrobatInstallPath & "plug_ins\Acrodx32.api"
        End If
    End If

    If Dir(strUserDestinationDir & "Windows\odma32.dll") = "" Then
        strHint = "Copying " & PCDocsPath & "Integration\odma32.dll to " & strUserDestinationDir & "Windows\odma32.dll"
        FileCopy PCDocsPath & "Integration\odma32.dll", strUserDestinationDir & "Windows\odma32.dll"
    End If

    'Modify HKCU registry key
    strHint = "Updating STARTUP-PATH under key HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\10\Word\Options to " & strUserDestinationDir & "Office\WordStart"
    oRegistry.UpdateKey REGToolRootTypes.HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Office\10.0\Word\Options", "STARTUP-PATH", strUserDestinationDir & "Microsoft Office\Startup"
    oRegistry.UpdateKey REGToolRootTypes.HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Office\10.0\Excel\Options", "AltStartup", strUserDestinationDir & "Microsoft Office\XLStart"

    Set oRegistry = Nothing

Exit Sub

IntegrationOnError:
    MsgBox "Unexpected error " & Err.Number & " (" & Err.Description & ") in EnableDocsIntegration." & vbCrLf & vbCrLf & "Hint:" & vbCrLf & strHint
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 01-04-2003 10:49
' Author    : Dave Robinson
' Purpose   : Take argument from command line to select appropriate Library for swapping, or bring up dialog box
'             Arguments accepted are:
'             'cdc'
'             'drims'
'             'sunrise'
'             'vincentenfield'
'             'onshoreprojects'
'             'disabled'
'
'  V    Date        Author          History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'
'---------------------------------------------------------------------------------------
Private Sub Form_Load()

  Dim intMode As Integer
  Dim strDRIMSGroupName(10) As String
  Dim strInputData As String
  Dim oUser As IADsUser
  Dim oGroup As IADsGroup
  Dim bIsDRIMSUser As Boolean
  Dim strArgument As String
  Dim strArguments() As String
  Dim intLoop As Integer

    strArguments() = Split(Command$, " ")

    For intLoop = 0 To UBound(strArguments())
        strArgument = LCase$(strArguments(intLoop))
        If strArgument = "debug" Then
            bDebugMode = True
        End If

        Select Case strArgument
          Case "cdc"
            SetLibrary 0
            Exit Sub '>---> Bottom
          Case "drims"
            SetLibrary 1
            Exit Sub '>---> Bottom
          Case "sunrise"
            SetLibrary 2
            Exit Sub '>---> Bottom
          Case "vincentenfield"
            SetLibrary 3
            Exit Sub '>---> Bottom
          Case "onshoreprojects"
            SetLibrary 4
            Exit Sub '>---> Bottom
          Case "disabled"
            SetLibrary 5
            Exit Sub '>---> Bottom
        End Select
    Next ':( Repeat For-Variable: INTLOOP

End Sub

'---------------------------------------------------------------------------------------
' Procedure : NoAccess
' DateTime  : 01-04-2003 10:47
' Author    : Dave Robinson
' Purpose   : Inform the user they will not be able to use DOCS if they continue
'
'  V    Date        Author          History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Private Function NoAccess(strLibraryName) As Integer

    NoAccess = MsgBox("You do not have access to this library." & vbCrLf & strLibraryName & " will not work correctly." & vbCrLf & vbCrLf & "Are you sure?", vbCritical + vbYesNo, "Swap Libraries Warning")

End Function

'---------------------------------------------------------------------------------------
' Procedure : optLibrary_Click
' DateTime  : 01-04-2003 10:47
' Author    : Dave Robinson
' Purpose   : Enable command button
'
'  V    Date        Author          History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Private Sub optLibrary_Click(Index As Integer)

    cmdSet.Enabled = True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetLibrary
' DateTime  : 01-04-2003 10:47
' Author    : Dave Robinson
' Purpose   : Determine library to set.  Initiate setting the appropriate library.
'
'  V    Date        Author          History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Private Sub SetLibrary(intLibraryNo As Integer)

  Dim isTerminalServer As Boolean
  Dim strIniFilename As String
  Dim strUsername As String
  Dim strDomainName As String
  Dim oUser As IADsUser
  Dim oGroup As IADsGroup
  Dim isCDCUser As Boolean
  Dim isDRIMSUser As Boolean
  Dim isSunriseUser As Boolean
  Dim isVincentEnfieldUser As Boolean
  Dim isOnshoreProjectsUser As Boolean
  Dim isDocumentControllerUser As Boolean
  Dim intReply As Integer
  Dim strDRIMSGroupName(10) As String
  Dim strInputData As String
  Dim strTempCommandLine As String

    On Error GoTo SetLibrary_Error

    strUsername = Environ$("USERNAME")
    strDomainName = Environ$("USERDOMAIN")

    ' Is this a terminal Server session?
    isTerminalServer = (Environ$("CLIENTNAME") <> "" And LCase$(Environ$("CLIENTNAME")) <> "console") ' True = YES:  Terminal Server Session

    If bDebugMode Then MsgBox "isTerminalServer: " & isTerminalServer ':( Expand Structure

    If isTerminalServer Then
        strProgramFilesDir = "D:\Program Files\"
      Else 'ISTERMINALSERVER = FALSE/0
        strProgramFilesDir = "C:\Program Files\"
    End If

    If bDebugMode Then MsgBox "strProgramFilesDir: " & strProgramFilesDir ':( Expand Structure

    strProgDrive = Left$(strProgramFilesDir, 3)  ' Get drive that program files are installed onto.  Format = 'x:\'

    PCDocsPath = strProgramFilesDir & "DRIMS\Progs\"

    strTempCommandLine = Replace(LCase$(Command$), "debug", "")

    Select Case strTempCommandLine
      Case "drims"
        strDRIMSGroupName(1) = "DRIMS_DOCS_USERS"
      Case "cdc"
        strDRIMSGroupName(2) = "DOCS_CONTROLLERS"
      Case "onshoreprojects"
        strDRIMSGroupName(3) = "OSP_USERS"
      Case "vincentenfield"
        strDRIMSGroupName(5) = "VEL_DOCS_USERS"
      Case "sunrise"
        strDRIMSGroupName(4) = "SUNRISE_DOCS_USERS"
    End Select

    Set oUser = GetObject("WinNT://" & strDomainName & "/" & strUsername & ",User")
    For Each oGroup In oUser.Groups
        Select Case LCase$(oGroup.Name)
          Case strDRIMSGroupName(2)
            isCDCUser = True
          Case strDRIMSGroupName(1)
            isDRIMSUser = True
          Case strDRIMSGroupName(4)
            isSunriseUser = True
          Case strDRIMSGroupName(5)
            isVincentEnfieldUser = True
          Case strDRIMSGroupName(3)
            isOnshoreProjectsUser = True
        End Select
    Next ':( Repeat For-Variable: OGROUP

    If bDebugMode Then MsgBox "intLibraryNo: " & intLibraryNo ':( Expand Structure

    Select Case intLibraryNo
      Case 0
        If Not isCDCUser Then
            intReply = NoAccess("CDC")
            If intReply = vbNo Then End ':( Expand Structure
        End If

        strIniFilename = "pcdocscdc.ini"
        If bDebugMode Then MsgBox "strIniFilename: " & strIniFilename ':( Expand Structure
        DisableDocsIntegration True                                 ' Turn DOCS ON (CDC)
        SuccessfulEnd "CDC"
      Case 1
        If Not isDRIMSUser Then
            intReply = NoAccess("DOCS")
            If intReply = vbNo Then End ':( Expand Structure
        End If

        strIniFilename = "pcdocsdrims.ini"
        If bDebugMode Then MsgBox "strIniFilename: " & strIniFilename ':( Expand Structure
        EnableDocsIntegration strIniFilename   ' Turn DOCS ON
        SuccessfulEnd "DRIMS"
      Case 2
        If Not isSunriseUser Then
            intReply = NoAccess("DOCS")
            If intReply = vbNo Then End ':( Expand Structure
        End If

        strIniFilename = "pcdocssunprod.ini"
        EnableDocsIntegration strIniFilename  ' Turn DOCS ON
        SuccessfulEnd "Sunrise"
      Case 3
        If Not isVincentEnfieldUser Then
            intReply = NoAccess("DOCS")
            If intReply = vbNo Then End ':( Expand Structure
        End If
        strIniFilename = "pcdocsveprod.ini"
        EnableDocsIntegration strIniFilename  ' Turn DOCS ON
        SuccessfulEnd "Vincent Enfield"
      Case 4
        If Not isOnshoreProjectsUser Then
            intReply = NoAccess("DOCS")
            If intReply = vbNo Then End ':( Expand Structure
        End If

        strIniFilename = "pcdocsopdmain.ini"
        EnableDocsIntegration strIniFilename  ' Turn DOCS ON
        SuccessfulEnd "Onshore Projects"
      Case 5
        strIniFilename = ""
        DisableDocsIntegration
        SuccessfulEnd ""
    End Select

    ' Clean up
    Set oUser = Nothing
    Set oGroup = Nothing
    End

    On Error GoTo 0 ':( Dead Code

Exit Sub

SetLibrary_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure SetLibrary of Form frmSelectLibrary", Error
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SuccessfulEnd
' DateTime  : 01-04-2003 10:47
' Author    : Dave Robinson
' Purpose   : Successful completion.  End nicely
'
'  V    Date        Author          History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Private Sub SuccessfulEnd(strLibraryName)

    If strLibraryName <> "" Then
        MsgBox "Successfully changed to the " & strLibraryName & " library.", vbOKOnly + vbInformation, "Success"
      Else 'NOT STRLIBRARYNAME...
        MsgBox "Successfully disabled DOCS.", vbOKOnly + vbInformation, "Success"
    End If

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jul-28 12:25) 7 + 584 = 591 Lines
