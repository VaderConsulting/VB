VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migrate Profile"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5400
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLogoff 
      Caption         =   "Logoff"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ListBox lstSavedDrives 
      Height          =   1425
      Left            =   5880
      TabIndex        =   12
      Top             =   2280
      Width           =   2895
   End
   Begin VB.ListBox lstSavedPrinters 
      Height          =   1425
      Left            =   5880
      TabIndex        =   6
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdLoadSettings 
      Caption         =   "Load Settings"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdRenameProfile 
      Caption         =   "Rename Profile"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtProfileDir 
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txtHomePage 
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdChangeDomain 
      Caption         =   "Change Domain"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtDomainName 
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3840
      Width           =   2295
   End
   Begin VB.ListBox lstDrives 
      Height          =   1425
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   3255
   End
   Begin VB.ListBox lstPrinters 
      Height          =   1425
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Save Settings"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton optManual 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton optManual 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.OptionButton optManual 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   375
   End
   Begin VB.OptionButton optManual 
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   18
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label lblProfileDir 
      Caption         =   "Profile Dir"
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblHomePage 
      Caption         =   "Home Page"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblDomainname 
      Caption         =   "Domain:"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
   Begin VB.Image imgStop 
      Height          =   495
      Left            =   120
      Picture         =   "frmMain.frx":068A
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDrives 
      Alignment       =   2  'Center
      Caption         =   "Drives"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label lblPrinters 
      Alignment       =   2  'Center
      Caption         =   "Printers"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      Caption         =   " "
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   4920
      Width           =   5175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChangeDomain_Click()
    'Dim strOU As String
    
    'strOU = "Desktops"
    
    'Shell strAppPath & "joindom.exe /a=APAC\robid /p=Passw0rd /d=APAC /O=" & strOU, vbNormalFocus
    
    'MsgBox "Please restart this computer", vbExclamation + vbOKOnly, "Message"
    
    MsgBox "Add this computer to the " & strDomainName & " Domain.  Restart when done, and load this application again.", vbExclamation + vbOKOnly, "Message"
    If Len(Dir(strAppPath & "joindom.exe")) > 0 Then
        Shell strAppPath & "joindom.exe", vbNormalFocus
    End If
End Sub

Private Sub cmdLoadSettings_Click()
    Dim strData As String
    Dim strHomepage As String
    Dim lResult As Long
    Dim i As Integer, j As Integer, k As Integer
    Dim strPrinters() As String
    Dim strDrives() As String
    Dim bFoundPrinter As Boolean
    Dim bFoundDrive As Boolean
    Dim strServername As String
    Dim strPrintername As String
    Dim strExistingDocumentumUsername As String
    Dim strSP As String
    Dim strOS As String
    Dim strReefEdgePath As String
    Dim strTestURL As String
    Dim strMinIPSecVersion As String
    Dim oFileVersion As New cFileVersion
    Dim strFileVersion As String
    Dim lTemp As Long
    Dim strExistingIPSecversion As String
    Dim strExecuteReefEdgePatch As String
    
    On Error GoTo Er
    
    optManual(2).Value = True
    
    If Len(Dir(strAppPath & Environ$("USERNAME") & "-" & Environ$("COMPUTERNAME") & ".txt")) > 0 Then
        ' Enumerate network printers
        GetPrinters
        
        lstSavedPrinters.Clear
        
        ' Compare lists.  Determine differences.
        Open strAppPath & Environ$("USERNAME") & "-" & Environ$("COMPUTERNAME") & ".txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, strData
                If Left(strData, 7) = "Printer" And Left(strData, 8) <> "Printers" Then
                    If InStr(1, strData, "\\") > 0 Then
                        frmMain.lstSavedPrinters.AddItem Trim(Mid(strData, InStr(1, strData, ":") + 1))
                    End If
                End If
            Loop
        Close 1
        
        For i = 0 To lstSavedPrinters.ListCount
            If Len(Trim(lstSavedPrinters.List(i))) > 0 Then
                bFoundPrinter = False
                
                For j = 0 To lstPrinters.ListCount
                    If lstSavedPrinters.List(i) = lstPrinters.List(j) Then
                        bFoundPrinter = True
                        Exit For
                    End If
                Next j
                
                If Not bFoundPrinter Then
                    k = k + 1
                    ReDim Preserve strPrinters(k)
                    strPrinters(k) = lstSavedPrinters.List(i)
                End If
                
            End If
        Next i
        
        ' Map those printers not connected.
        If k > 0 Then
            For i = 1 To k
                strServername = Left(strPrinters(i), InStr(3, strPrinters(i), "\", vbTextCompare) - 1)
                strPrintername = Mid(strPrinters(i), InStr(3, strPrinters(i), "\", vbTextCompare) + 1)
                AddPrinterConnection strServername & "\" & strPrintername
            Next i
        End If
        
        ' Enumerate Network Drives.
        GetDrives
        
        lstSavedDrives.Clear
        
        ' Compare lists.  Determine differences.
        Open strAppPath & Environ$("USERNAME") & "-" & Environ$("COMPUTERNAME") & ".txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, strData
                If Left(strData, 5) = "Drive" And Left(strData, 6) <> "Drives" Then
                    If InStr(1, strData, "\\") > 0 Then
                        frmMain.lstSavedDrives.AddItem Trim(Mid(strData, InStr(1, strData, ":") + 1))
                    End If
                End If
            Loop
        Close 1
        
        k = 0
        For i = 0 To lstSavedDrives.ListCount
            If Len(Trim(lstSavedDrives.List(i))) > 0 Then
                bFoundDrive = False
                
                For j = 0 To lstDrives.ListCount
                    If lstSavedDrives.List(i) = lstDrives.List(j) Then
                        bFoundDrive = True
                        Exit For
                    End If
                Next j
                
                If Not bFoundDrive Then
                    Debug.Print "Did not find " & lstSavedDrives.List(i)
                    k = k + 1
                    ReDim Preserve strDrives(k)
                    strDrives(k) = lstSavedDrives.List(i)
                End If
                
            End If
        Next i
        
        ' Map drives not connected.
        If k > 0 Then
            For i = 1 To UBound(strDrives())
                lResult = Connect2(Left(strDrives(i), 1) & ":", Mid(strDrives(i), 5))
            Next i
        End If
        
        ' Retrieve original homepage
        Open strAppPath & Environ$("USERNAME") & "-" & Environ$("COMPUTERNAME") & ".txt" For Input As #1
            Do Until EOF(1)
                Line Input #1, strData
                If Left(strData, 9) = "Home page" Then
                    strHomepage = Trim(Mid(strData, InStr(1, strData, ":") + 1))
                    Exit Do
                End If
            Loop
        Close 1
        
        If strHomepage <> "" Then
            ' Set homepage as per original.
            lResult = SaveRegString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page", strHomepage)
        End If
        
        MsgBox "Sequence complete.  Please check user configuration.", vbExclamation + vbOKOnly, "Message"
        iReply = MsgBox("Click OK to test Internet Explorer", vbOKCancel, "Message")
        If iReply = vbOK Then
            ' Retrieve specific URL to test IE with, then open IE with it.
            ShellExecute frmMain.hwnd, "Open", GetIniValue("Setup", "TestURL", "", strAppPath & "setup.ini"), "", strAppPath, 1
        End If
        
        ' Get Documentum username from registry
        strExistingDocumentumUsername = GetRegString(HKEY_CURRENT_USER, "Software\Documentum\Common\LastConnectedDocbase", "User", "")
        
        ' Check for documentum
        If Len(strExistingDocumentumUsername) > 0 Then
            SaveRegString HKEY_CURRENT_USER, "Software\Documentum\Common\LastConnectedDocbase", "User", LCase(Environ$("USERNAME"))
            SaveRegString HKEY_CURRENT_USER, "Software\Documentum\Common\LastConnectedDocbase", "Domain", strDomainName
        End If
        
        Me.cmdLoadSettings.Enabled = False
        Me.cmdLogoff.Enabled = False
        
        ' Get Service pack
        strSP = GetSP
        
        ' Get OS
        strOS = GetOS
        
        'Check for ReefEdge
        If Len(Dir("C:\Program Files\ReefEdge", vbDirectory)) > 0 Then
            ' Get existing file version info
            lTemp = oFileVersion.OpenFile(Environ$("SYSTEMROOT") & "\System32\Drivers\ipsec.sys")
            strFileVersion = oFileVersion.FileVersion
            ' This is the complete version string, eg 5.1.2600.1106 (on XP)
            ' strMinIPSecVersion will only specify the number after the last decimal
            strExistingIPSecversion = Mid(strFileVersion, InStrRev(strFileVersion, ".") + 1) ' This is also the number after the last decimal
            
            Select Case strOS
                Case "Windows NT 4"
                    If strSP <> GetIniValue("ExactServicePackStringforReefEdgePatchExclusionPerOS", "NT4", "", strAppPath & "setup.ini") Then
                        strFileVersion = oFileVersion.OpenFile(Environ$("SYSTEMROOT") & "\System32\Drivers\ipsec.sys")
                        strMinIPSecVersion = GetIniValue("MinimumIPSec.SysVersion", "NT4", "", strAppPath & "setup.ini")
                        If CInt(strExistingIPSecversion) < CInt(strMinIPSecVersion) Then
                            strExecuteReefEdgePatch = LCase(GetIniValue("Setup", "ExecuteReefEdgePatch", "", strAppPath & "setup.ini"))
                            If strExecuteReefEdgePatch = "true" Or strExecuteReefEdgePatch = "1" Or strExecuteReefEdgePatch = "yes" Then
                                ' Update ReefEdge
                                strReefEdgePath = GetIniValue("IPSecPatchLocations", "NT4", "", strAppPath & "setup.ini")
                                Shell strReefEdgePath, vbNormalFocus
                            Else
                                MsgBox "Reefedge Patch not installed.", vbExclamation + vbOKOnly, "Message"
                            End If
                        End If
                    End If
                Case "Windows 2000"
                    If strSP <> GetIniValue("ExactServicePackStringforReefEdgePatchExclusionPerOS", "2000", "", strAppPath & "setup.ini") Then
                        strFileVersion = oFileVersion.OpenFile(Environ$("SYSTEMROOT") & "\System32\Drivers\ipsec.sys")
                        strMinIPSecVersion = GetIniValue("MinimumIPSec.SysVersion", "2000", "", strAppPath & "setup.ini")
                        If CInt(strExistingIPSecversion) < CInt(strMinIPSecVersion) Then
                            strExecuteReefEdgePatch = LCase(GetIniValue("Setup", "ExecuteReefEdgePatch", "", strAppPath & "setup.ini"))
                            If strExecuteReefEdgePatch = "true" Or strExecuteReefEdgePatch = "1" Or strExecuteReefEdgePatch = "yes" Then
                                ' Update ReefEdge
                                strReefEdgePath = GetIniValue("IPSecPatchLocations", "2000", "", strAppPath & "setup.ini")
                                Shell strReefEdgePath, vbNormalFocus
                            Else
                                MsgBox "Reefedge Patch not installed.", vbExclamation + vbOKOnly, "Message"
                            End If
                        End If
                    End If
                Case "Windows XP"
                    If strSP <> GetIniValue("ExactServicePackStringforReefEdgePatchExclusionPerOS", "XP", "", strAppPath & "setup.ini") Then
                        strFileVersion = oFileVersion.OpenFile(Environ$("SYSTEMROOT") & "\System32\Drivers\ipsec.sys")
                        strMinIPSecVersion = GetIniValue("MinimumIPSec.SysVersion", "XP", "", strAppPath & "setup.ini")
                        If CInt(strExistingIPSecversion) < CInt(strMinIPSecVersion) Then
                            strExecuteReefEdgePatch = LCase(GetIniValue("Setup", "ExecuteReefEdgePatch", "", strAppPath & "setup.ini"))
                            If strExecuteReefEdgePatch = "true" Or strExecuteReefEdgePatch = "1" Or strExecuteReefEdgePatch = "yes" Then
                                ' Update ReefEdge
                                strReefEdgePath = GetIniValue("IPSecPatchLocations", "XP", "", strAppPath & "setup.ini")
                                Shell strReefEdgePath, vbNormalFocus
                            Else
                                MsgBox "Reefedge Patch not installed.", vbExclamation + vbOKOnly, "Message"
                            End If
                        End If
                    End If
                Case "Windows 2003"
                    If strSP <> GetIniValue("ExactServicePackStringforReefEdgePatchExclusionPerOS", "2003", "", strAppPath & "setup.ini") Then
                        strFileVersion = oFileVersion.OpenFile(Environ$("SYSTEMROOT") & "\System32\Drivers\ipsec.sys")
                        strMinIPSecVersion = GetIniValue("MinimumIPSec.SysVersion", "2003", "", strAppPath & "setup.ini")
                        If CInt(strExistingIPSecversion) < CInt(strMinIPSecVersion) Then
                            strExecuteReefEdgePatch = LCase(GetIniValue("Setup", "ExecuteReefEdgePatch", "", strAppPath & "setup.ini"))
                            If strExecuteReefEdgePatch = "true" Or strExecuteReefEdgePatch = "1" Or strExecuteReefEdgePatch = "yes" Then
                                ' Update ReefEdge
                                strReefEdgePath = GetIniValue("IPSecPatchLocations", "2003", "", strAppPath & "setup.ini")
                                Shell strReefEdgePath, vbNormalFocus
                            Else
                                MsgBox "Reefedge Patch not installed.", vbExclamation + vbOKOnly, "Message"
                            End If
                        End If
                    End If
            End Select
        End If
    Else
        MsgBox "Sequence incomplete.  User settings not found.", vbCritical + vbOKOnly, "Error"
    End If
    
    Set oFileVersion = Nothing
    
    Exit Sub
Er:
    frmMain.lblInfo = "Unexpected error " & Err.Number & " (" & Err.Description & ")."
    Resume Next
End Sub

Private Sub cmdLogoff_Click()
    ShutdownSystem EWX_LOGOFF
End Sub

Private Sub cmdRenameProfile_Click()
    optManual(1).Value = True
    
    frmProfiles.Show vbModal
    
    MsgBox "Please log on as the user in the " & strDomainName & " Domain, and run this application again.", vbExclamation + vbOKOnly, "Message"
    
    Me.cmdRenameProfile.Enabled = False
    Me.cmdLogoff.Enabled = True
End Sub

Private Sub cmdStart_Click()
    optManual(0).Value = True
    
    GetPrinters
    GetDrives
    GetHomePage
    GetProfileDir
    
    SaveData
    
    MsgBox "Please Perform the following steps:" & vbCrLf & "Logoff this user," & vbCrLf & "Log back on as the same user in the " & strDomainName & " domain." & vbCrLf & "Log off" & vbCrLf & "Log on again as an administrative user." & vbCrLf & "Finally, run this application again.", vbExclamation + vbOKOnly, "Message"
    
    Me.cmdStart.Enabled = False
    Me.cmdLogoff.Enabled = True
    
    Set oFSO = Nothing
End Sub

Private Sub Form_Load()
    If Right(App.Path, 1) <> "\" Then
        strAppPath = App.Path & "\"
    Else
        strAppPath = App.Path
    End If
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    strTempDir = GetIniValue("Setup", "Temp Dir", "", strAppPath & "setup.ini")
    strDomainName = GetIniValue("Setup", "Domain", "", strAppPath & "setup.ini")
    
    ' Have we already saved settings for this computer?
    If Len(Dir(strAppPath & "*-" & Environ$("COMPUTERNAME") & ".txt")) > 0 Then
        ' Yes
        cmdStart.Enabled = False
        
        ' Have we renamed the profile?
        If Right(UCase(Environ$("USERPROFILE")), 4) = strDomainName Or Command$ = "debug" Then
            ' Yes
            cmdLoadSettings.Enabled = True
        Else
            ' No, not yet
            cmdRenameProfile.Enabled = True
        End If
    Else
        ' No
        cmdStart.Enabled = True
    End If
    
    ' Is this computer in the correct domain?
    frmMain.txtDomainName.Text = GetRegString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\WinLogon", "cachePrimaryDomain", "")
    
    ' If not, display the stop icon!
    If UCase(frmMain.txtDomainName.Text) <> strDomainName Then
        If Command$ <> "debug" And UCase(frmMain.txtDomainName.Text) <> "HOME" Then
            frmMain.imgStop.Visible = True
            AddInfo "This computer is not in the correct Domain.", "This computer is not in the correct Domain."
            frmMain.cmdChangeDomain.Enabled = True
            frmMain.cmdStart.Enabled = False
        End If
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set oFSO = Nothing
End Sub

Private Sub optManual_Click(Index As Integer)
    Select Case Index
        Case 0
            cmdStart.Enabled = True
            cmdRenameProfile.Enabled = False
            cmdLoadSettings.Enabled = False
            cmdLogoff.Enabled = False
        Case 1
            cmdStart.Enabled = False
            cmdRenameProfile.Enabled = True
            cmdLoadSettings.Enabled = False
            cmdLogoff.Enabled = False
        Case 2
            cmdStart.Enabled = False
            cmdRenameProfile.Enabled = False
            cmdLoadSettings.Enabled = True
            cmdLogoff.Enabled = False
        Case 3
            cmdStart.Enabled = False
            cmdRenameProfile.Enabled = False
            cmdLoadSettings.Enabled = False
            cmdLogoff.Enabled = True
    End Select
End Sub
