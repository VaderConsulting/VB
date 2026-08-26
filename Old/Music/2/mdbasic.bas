Attribute VB_Name = "mdMain"
Option Explicit
Option Compare Text
'
'Global Variables
Public gProgramDir                                      ' Directory where program is installed
Public gWindowsDir                                      ' Windows directory
Public gUserName                                        ' Name of logged in user
Public gUserPassword                                    ' Password of logged in user
Public gUserLevel                                       ' Level of logged in user
'Public gNewUserLevel
Public gUserLock                                        ' If 1 indicates current user is locked out
Public gRegistered                                      ' Product is registered if 1
Public gVolume                                          ' Volume label of C: Drive
Public gSerial                                          ' Volume serial Number of C: Drive
Public gCompany                                         ' Registered user's Company Name
Public gPlace                                           ' Registered user's City / Suburb / Town
Public gName                                            ' Registered user's First Name
Public gSurname                                         ' Registered user's Surname
Public gAddress                                         ' Registered user's Address
Public gState                                           ' Registered user's State
Public gPostcode                                        ' Registered user's Postcode
Public gPhone                                           ' Registered user's Phone Number
Public gFax                                             ' Registered user's Fax Number
Public gFooter                                          ' Footer printed at bottom of page
Public gLogo                                            ' Which logo? 0 = CBB, 1 = Keys, 2 = User
Public gLogoPath                                        ' Path for the User defined program logo
Public gDatabasePath
Public gSongIndex
Public gTime                                            ' Used when restricting time with TimeCripple
'
'Global Program Constants
Public Const gProgramName = "Music Database"            ' Program Name
Public Const gDate = "November 1996 - June 1997"        ' Development dates
Public Const gINIFile = "MUSIC"                         ' Name of .INI file eg GENERIC.INI.  Use ALL-CAPS!
Public Const gAuthor = "D. Robinson"                    ' Program Author.  Please make mention of me.
Public Const gEmail = "dave.robinson@usa.net"           ' Authors EMAIL Address
Public Const gHome = "Homeless"                         ' Authors home town
Public Const gAuthorChristian = "Dave"                  ' Authors Christian name
Public Const gAuthorSurname = "Robinson"                ' Authors Surname
Public Const gBusiness = "Chips, Bits and Bytes"        ' Authors business name
Public Const gAuthorAddress = " "                       ' Authors Address
Public Const gAuthorPhone = "0417 927 828"              ' Authors Phone number
Public Const gAuthorFax = " "                           ' Authors Fax number
Public Const gAuthorState = "Australia"                 ' Authors State
Public Const gAuthorPostcode = ""                       ' Authors Postcode
Public Const gPassword = 0                              ' if 1 then password protection enabled
Public Const gRememberName = 1                          ' remember name option in login window
Public Const gKey$ = "PASSWORD"                         ' key used to decode/encode passwords
Public Const gTimeCripple = 1                           ' cripple after a period of time if 1 **NOT IMPLEMEMTED**
Public Const gDemo = 0                                  ' product is demo if 1 **NOT IMPLEMEMTED**
Public Const gVGAWidth = 9720                           ' Form width for 640x480 screens
Public Const gVGAHeight = 7320                          ' Form height for 640x480 screens
Public Const gUpgrade = 20                              ' Price to upgrade from a previous version A$
Public Const gFullPrice = 40                            ' Full product price A$
'
'
'Other Global Constants
'
Global Const gstrNULL$ = ""                             'Empty string
Global Const gstrSEP_DIR$ = "\"                         'Directory separator character
Global Const gstrSEP_DIRALT$ = "/"                      'Alternate directory separator character
Global Const gstrSEP_EXT$ = "."                         'Filename extension separator character
Global Const gstrCOLON$ = ":"
Global Const gstrSwitchPrefix1 = "-"
Global Const gstrSwitchPrefix2 = "/"
Global Const gstrCOMMA$ = ","
Global Const gstrDECIMAL$ = "."
Global Const gstrINI_PROTOCOL = "Protocol"

Global Const gintMAX_SIZE% = 255                        'Maximum buffer size
Global Const gintMIN_BUTTONWIDTH% = 1200
Global Const gsngBUTTON_BORDER! = 1.4

Global Const intDRIVE_REMOVABLE% = 2                    'Constants for GetDriveType
Global Const intDRIVE_FIXED% = 3
Global Const intDRIVE_REMOTE% = 4

Global Const gintNOVERINFO% = 32767                     'flag indicating no version info

'File names
Global Const gstrFILE_SETUP$ = "SETUP.LST"              'Name of setup information file

'Share type macros for files
Global Const mstrPRIVATEFILE = ""
Global Const mstrSHAREDFILE = "$(Shared)"

'INI File keys
#If Win16 Then
Global Const gstrINI_BTRIEVE$ = "Btrieve"
#End If
Global Const gstrINI_SETUP$ = "Setup"
Global Const gstrINI_APPNAME$ = "Title"
Global Const gstrINI_APPDIR$ = "DefaultDir"
Global Const gstrINI_APPEXE$ = "AppExe"
Global Const gstrINI_APPPATH$ = "AppPath"
Global Const gstrINI_FORCEUSEDEFDEST = "ForceUseDefDir"

'Setup information file macros
Global Const gstrAPPDEST$ = "$(AppPath)"
Global Const gstrWINDEST$ = "$(WinPath)"
Global Const gstrWINSYSDEST$ = "$(WinSysPath)"
Global Const gstrWINSYSDESTSYSFILE$ = "$(WinSysPathSysFile)"
Global Const gstrPROGRAMFILES$ = "$(ProgramFiles)"
Global Const gstrCOMMONFILES$ = "$(CommonFiles)"
Global Const gstrCOMMONFILESSYS$ = "$(CommonFilesSys)"
Global Const gstrDAODEST$ = "$(MSDAOPath)"

'Mouse Pointer Constants
Global Const gintMOUSE_DEFAULT% = 0
Global Const gintMOUSE_HOURGLASS% = 11

'MsgError() Constants
Global Const MSGERR_ERROR = 1
Global Const MSGERR_WARNING = 2

'MsgBox Constants
Global Const MB_OK = 0                                  'OK button only
Global Const MB_OKCANCEL = 1                            'OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2                    'Abort, Retry, Ignore buttons
Global Const MB_YESNO = 4                               'Yes and No buttons
Global Const MB_RETRYCANCEL = 5                         'Retry and Cancel buttons
Global Const MB_ICONSTOP = 16                           'Critical message
Global Const MB_ICONQUESTION = 32                       'Warning query
Global Const MB_ICONEXCLAMATION = 48                    'Warning message
Global Const MB_ICONINFORMATION = 64                    'Information message
Global Const MB_DEFBUTTON1 = 0                          'First button is default
Global Const MB_DEFBUTTON2 = 256                        'Second button is default
Global Const MB_DEFBUTTON3 = 512                        'Third button is default

'MsgBox return values
Global Const IDOK = 1                                   'OK button pressed
Global Const IDCANCEL = 2                               'Cancel button pressed
Global Const IDABORT = 3                                'Abort button pressed
Global Const IDRETRY = 4                                'Retry button pressed
Global Const IDIGNORE = 5                               'Ignore button pressed
Global Const IDYES = 6                                  'Yes button pressed
Global Const IDNO = 7                                   'No button pressed

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
'
'Type Definitions
'
Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    nReserved1 As Integer
    nReserved2 As Integer
    szPathName As String * 256
End Type

Type VERINFO                                            'Version FIXEDFILEINFO
    strPad1 As Long                                     'Pad out struct version
    strPad2 As Long                                     'Pad out struct signature
    nMSLo As Integer                                    'Low word of ver # MS DWord
    nMSHi As Integer                                    'High word of ver # MS DWord
    nLSLo As Integer                                    'Low word of ver # LS DWord
    nLSHi As Integer                                    'High word of ver # LS DWord
    strPad3(1 To 36) As Byte                            'Pad out rest of VERINFO struct (36 bytes)
End Type

Type PROTOCOL
    strName As String
    strFriendlyName As String
End Type

Global Const OF_EXIST& = &H4000&
Global Const OF_SEARCH& = &H400&
Global Const HFILE_ERROR% = -1

'
'Global Variables
'
Global LF$                                              'single line break
Global LS$                                              'double line break

'List of available protocols
Global gProtocol() As PROTOCOL
Global gcProtocols As Integer


#If Win16 Then
    Declare Function GetWinPlatform Lib "STKIT416.DLL" () As Long
    Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Declare Function GetFreeSystemResources% Lib "User" (ByVal free%)
#End If

Declare Function sndPlaySound Lib "MMSYSTEM.DLL" (ByVal lpszSoundName$, ByVal wFlags%) As Integer

Public Function fctPassCheck(Name, Password)
    Dim pass As String      '          = stored password
    Dim level As Integer    '          = stored level
                            ' name     = typed name
                            ' password = typed password
    gUserLevel = 1
    fctPassCheck = 0
    frmLogin.lblConfirm.Visible = 0
    frmLogin.txtConfirm.Visible = 0
    frmLogin.cmdOK.Visible = 1
    frmLogin.cmdConfirm.Visible = 0
    frmLogin.txtName.Locked = 0
    frmLogin.txtPassword.Locked = 0
    fctPassCheck = 0
    '
    '   codes for fctPassCheck:
    '   0 = NO - password incorrect
    '   1 = OK - password correct
    '   2 = OK - new user
    '   3 = NO - no password supplied
    '   4 = NO - account locked
    '   5 = NO - no user name supplied
    '
    If Name <> "" Then
        If Password <> "" Then                                  ' password entered
            If Name = "DAVE" And Password = "JASECODY" Then     ' SUPER USER Back door for accessing
                gUserName = "DAVE"                              ' program when ADMIN Password is forgotten
                gUserPassword = Password
                fctPassCheck = 1
                gUserLock = 0
                gUserLevel = 99
            Else
                gUserName = Name
                gUserPassword = fctGetPassword(gUserName)
                gUserLock = fctGetLock(gUserName)
                gUserLevel = fctGetLevel(gUserName)
                If Password = gUserPassword And gUserLock = 0 Then
                    fctPassCheck = 1                            ' existing user, correct password
                End If
                If Password = gUserPassword And gUserLock = 1 Then
                    fctPassCheck = 4                            ' existing user, correct password, locked out
                End If
                If gUserPassword = "Default" Then               ' new user
                    fctPassCheck = 2
                End If
            End If
        Else                                                    ' no password typed
            fctPassCheck = 3
        End If
    Else
    fctPassCheck = 5
    End If
End Function
Public Function fctSaveSettings(UserName, UserPassword, UserLevel, UserLocks)
    Dim varPassword As String
    varPassword = fctEncode(UserPassword)
    SaveSetting gINIFile, "Passwords", UserName, varPassword
    SaveSetting gINIFile, "Levels", UserName, UserLevel
    SaveSetting gINIFile, "Locks", UserName, UserLocks
    If gRememberName = 1 Then                                   'save this user as the last person to log in
        SaveSetting gINIFile, "Login", "Last", UserName
    End If
End Function
Public Sub subAdd()
    frmAdmin.Show
    frmStart.Hide
End Sub
Public Sub subConfirm()
    Dim msg As String
    Dim varStyle As Integer
    Dim response As Integer
    Dim varPassword As String
    Dim varResult As Integer
    If frmLogin.txtPassword = frmLogin.txtConfirm Then
        gUserLevel = 1
        gUserName = UCase(frmLogin.txtName)
        gUserPassword = UCase$(frmLogin.txtPassword)
        varResult = fctSaveSettings(gUserName, gUserPassword, 1, 0)
        SubSaveLastName
        frmLicense.Show
        frmLogin.Hide
    Else
        msg = "Your password confirmation is incorrect"
        varStyle = vbOKOnly + vbCritical + vbDefaultButton1
        response = MsgBox(msg, varStyle, "Error")
        frmLogin.txtConfirm = ""
        frmLogin.txtConfirm.SetFocus
    End If
End Sub
Public Function fctCrypt(gKey$, varPassword)
    Dim varCounter As Integer
    Dim varLoop As Integer
    Dim var As Integer
    varCounter = 1
    For varLoop = 1 To Len(varPassword)
        var = Asc(Mid$(gKey$, varCounter, 1))
        varCounter = varCounter + 1
        If varCounter > Len(gKey$) Then varCounter = 1
        Mid$(varPassword, varLoop, 1) = Chr$(Asc(Mid$(varPassword, varLoop, 1)) Xor var)
    Next
    fctCrypt = varPassword
End Function
Public Function fctDecode(varIn)
    'To read it back in,
    Dim varLoop As Integer
    Dim J As String
    Dim varPassword As String
    Dim varHex As String
    varHex = varIn
    varHex = Mid$(varHex, 3, Val(Left$(varHex, 2)))
    varPassword = ""
    For varLoop = 1 To Len(varHex) Step 2
        J = Mid$(varHex, varLoop, 2)
        varPassword = varPassword + Chr$(Val("&H" + J))
    Next
    fctDecode = fctCrypt(gKey$, varPassword)
End Function
Public Sub subDelete()
    frmAdmin.Show
    frmStart.Hide
End Sub
Public Function fctEncode(varIn)
    Dim varLoop As Integer
    Dim J As String
    Dim varHex As String
    Dim varPassword As String
    varPassword = fctCrypt(gKey$, varIn)
    'When writing an encrypted password to a sequential access file like the
    'INI files, you need to convert the resultant encrypted file to hex data.
    'This is because you can end up with an encrypted password that contains
    'characters which cannot be properly read using sequential access.  So,
    'before saving your encrypted password, use this routine:
    varHex = ""
    For varLoop = 1 To Len(varIn)
        J = Hex(Asc(Mid$(varIn, varLoop, 1)))
        If Len(J) = 1 Then J = "0" + J
        varHex = varHex + J
    Next
    'This will create a string like "0EF31105" or some such.  Save that to
    'the INI file.
    'Store the LENGTH of the password string as 2 bytes and concatenate
    varHex = Format$(Len(varHex), "00") + varHex
    fctEncode = varHex
End Function
Public Sub subEnd()
    Dim msg As String
    Dim varStyle As Integer
    Dim response As Integer
    msg = "Do you really want to exit"
    varStyle = vbYesNo + vbQuestion + vbDefaultButton2
    response = MsgBox(msg, varStyle, "Quit")
    If response = IDYES Then
        End
    End If
End Sub
Public Sub subLogin()
    Dim varPassword As String
    Dim varDir As String
    Dim varRemember As Long
    frmLogin.chkRemember.Enabled = gRememberName
    varDir = Dir(gWindowsDir & "\" & gINIFile & ".INI")
    If gRegistered = 0 Then
        gCompany = gBusiness
        gPlace = gHome
        gSurname = gAuthorSurname
        gAddress = gAuthorAddress
        gPhone = gAuthorPhone
        gFax = gAuthorFax
        gState = gAuthorState
        gPostcode = gAuthorPostcode
    End If
    'Create a default .INI file
    If varDir = "" Then
        Open gWindowsDir & "\" & gINIFile & ".INI" For Output As #1
        varPassword = fctEncode("PASSWORD")
        Print #1, "[Passwords]"
        Print #1, "ADMIN="; varPassword
        Print #1, ""
        Print #1, "[Levels]"
        Print #1, "ADMIN=98"
        Print #1, ""
        Print #1, "[Locks]"
        Print #1, "ADMIN=0"
        Print #1, ""
        Close #1
        MsgBox ("Default " & gINIFile & ".INI file created." & Chr$(10) & "Please refer to your documentation for a User Name and Password.")
    End If
    If gPassword = 0 Then
        frmStart.mnuLogout.Enabled = 0
        frmStart.mnuModifyUser.Enabled = 0
        frmStart.mnuShow.Enabled = 0
        frmLicense.Show
        frmLogin.Hide
    Else                                                            ' the Remember Name option in the code has been enabled
        varRemember = GetSetting(gINIFile, "Setup", "Remember", 0)
        If varRemember = 0 Then
            frmLogin.chkRemember.Value = varRemember
            frmLogin.txtName.SetFocus
        Else                                                        ' the user has opted to remember names
            frmLogin.chkRemember.Value = 1
            frmLogin.txtName = GetSetting(gINIFile, "Login", "Last", "ADMIN")
            frmLogin.txtPassword.SetFocus
        End If
        If frmLogin.chkRemember = 0 Then frmLogin.txtName = ""
        frmLogin.txtName.Locked = 0
        frmLogin.txtPassword.Locked = 0
        frmLogin.txtPassword = ""
        frmLogin.cmdConfirm.Visible = 0
        frmLogin.cmdOK.Visible = 1
        frmLogin.txtConfirm = ""
        frmLogin.lblConfirm.Visible = 0
        frmLogin.txtConfirm.Visible = 0
        frmLogin.cmdOK.Default = 1
    End If
End Sub
Public Sub subModify()
    frmAdmin.Show
    frmStart.Hide
End Sub
Public Sub subRefresh()
    Dim varNames As Variant
    Dim varTop As Integer
    Dim varLoop As Integer
    
    frmAdmin.lstUsers.Clear
    frmAdmin.lstUsers.Enabled = 1
    frmAdmin.txtPassword = ""
    frmAdmin.txtPassword.PasswordChar = "*"
    frmAdmin.lblPassword = "Current password:"
    frmAdmin.cmdUpdate.Enabled = 0
    frmAdmin.cmdUpdate.Visible = 1
    frmAdmin.cmdRefresh.Visible = 0
    frmAdmin.txtNewPassword = ""
    frmAdmin.txtLevel = ""
    frmAdmin.chkLocked = 0
    
    If gUserLevel = 99 Then frmAdmin.txtPassword.PasswordChar = "" ' make passwords visible for superuser
    varNames = GetAllSettings(gINIFile, "Passwords")
    varTop = UBound(varNames, 1)
    For varLoop = LBound(varNames, 1) To UBound(varNames, 1)
        frmAdmin.lstUsers.AddItem varNames(varLoop, 0)
    Next varLoop
End Sub
Public Function fctGetPassword(UserName)
    Dim Result As String
    Dim varPassword As String
    Result = GetSetting(gINIFile, "Passwords", UserName, "Default")
    If Result = "Default" Then
    Else
        Result = fctDecode(Result)
    End If
    fctGetPassword = Result
End Function
Public Function fctGetLevel(UserName)
    Dim Result As String
    Result = GetSetting(gINIFile, "Levels", UserName, 1)
    fctGetLevel = Result
End Function
Public Function fctGetLock(UserName)
    Dim Result As String
    Result = GetSetting(gINIFile, "Locks", UserName, 0)
    fctGetLock = Result
End Function
Public Sub SubSaveLastName()
    If frmLogin.chkRemember <> 0 Then
        SaveSetting gINIFile, "Setup", "Remember", 1
    Else
        SaveSetting gINIFile, "Setup", "Remember", 0
    End If
    SaveSetting gINIFile, "Login", "Last", gUserName
End Sub
Public Sub subAbout()
    frmAbout.Show
End Sub
Function StripTerminator(ByVal strString As String) As String
'-----------------------------------------------------------
' FUNCTION: StripTerminator
'
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.
'
' IN: [strString] - String to remove terminator from
'
' Returns: The value of the string passed in minus any
'          terminating zero.
'-----------------------------------------------------------
'

    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Sub AddDirSep(strPathName As String)
'-----------------------------------------------------------
' SUB: AddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'
' IN/OUT: [strPathName] - path to add separator to
'-----------------------------------------------------------
'

    If Right$(RTrim$(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub

Public Sub Wait(Interval)   ' waits the specified period, in seconds
    Dim Start As Long
    Start = Timer
    While Timer < Start + Interval
        DoEvents            ' process normal windows events
    Wend
End Sub

Public Function fctRND(Low, High) 'returns a random number between low and high inclusive
    Randomize Timer
    fctRND = Int(((Val(High) - Val(Low) + 1) * Rnd) + Val(Low))
End Function

Public Sub subAuthor()
    Dim msg As String
    Dim Result As Integer
    msg = gAuthorChristian & " " & gAuthorSurname & Chr$(13)
    msg = msg & "Email address : " & gEmail & Chr$(13)
    msg = msg & "Phone number : " & gAuthorPhone & Chr$(13)
    msg = msg & "Fax number : " & gAuthorFax & Chr$(13)
    Result = MsgBox(msg, vbInformation, "About me")
End Sub
Public Sub subGetRegoInfo()
    ' ********** GET REGISTRATION INFO **************
    Dim Code As String
    gRegistered = GetSetting(gINIFile, "Register", "Registered", 0)
    Code = GetSetting(gINIFile, "Register", "Company", gBusiness)
    If Code <> gBusiness Then
        gCompany = fctDecode(Code)    ' if it isn't the default, then decode it
    Else
        gCompany = gBusiness
    End If
    Code = GetSetting(gINIFile, "Register", "Place", gHome)
    If Code <> gHome Then
        gPlace = fctDecode(Code)        ' if it isn't the default, then decode it
    Else
        gPlace = gHome
    End If
    Code = GetSetting(gINIFile, "Register", "Name", gAuthorChristian)
    If Code <> gAuthorChristian Then
        gName = fctDecode(Code)      ' if it isn't the default, then decode it
    Else
        gName = gAuthorChristian
    End If
    Code = GetSetting(gINIFile, "Register", "Surname", gAuthorSurname)
    If Code <> gAuthorSurname Then
        gSurname = fctDecode(Code)      ' if it isn't the default, then decode it
    Else
        gSurname = gAuthorSurname
    End If
    Code = GetSetting(gINIFile, "Register", "Address", gAuthorAddress)
    If Code <> gAuthorAddress Then
        gAddress = fctDecode(Code)      ' if it isn't the default, then decode it
    Else
        gAddress = gAuthorAddress
    End If
    Code = GetSetting(gINIFile, "Register", "Phone", gAuthorPhone)
    If Code <> gAuthorPhone Then
        gPhone = fctDecode(Code)
    Else
        gPhone = gAuthorPhone
    End If
    Code = GetSetting(gINIFile, "Register", "Fax", gAuthorFax)
    If Code <> gAuthorFax Then
        gFax = fctDecode(Code)
    Else
        gFax = gAuthorFax
    End If
    Code = GetSetting(gINIFile, "Register", "State", gAuthorState)
    If Code <> gAuthorState Then
        gState = fctDecode(Code)
    Else
        gState = gAuthorState
    End If
    Code = GetSetting(gINIFile, "Register", "Postcode", gAuthorPostcode)
    If Code <> gAuthorPostcode Then
        gPostcode = fctDecode(Code)
    Else
        gPostcode = gAuthorPostcode
    End If
End Sub
Public Function fctGetDriveInfo()
    Dim varLine As Integer
    Dim Volume As String
    Dim Serial As String
    Dim Info(20) As String
    'While gVolume = ""
        Shell "COMMAND.COM /C dir C:\*.BAT >C:\INFO.TXT"
        Wait (2)                    ' Give the drive 2 seconds to write the file
        varLine = 1
        Open "C:\INFO.TXT" For Input As #1
            While Not EOF(1)
                Line Input #1, Info(varLine)
                varLine = varLine + 1
            Wend
        Close #1
        Kill "C:\INFO.TXT"
        Volume = Mid$(Info(2), 23, 11)
        Serial = Mid$(Info(3), 26, 9)
    'Wend
    fctGetDriveInfo = Serial & Volume
End Function
