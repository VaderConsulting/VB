Attribute VB_Name = "ModStart"
'Quick Reference...
Global QuickRef As tQuickRef

Type tQuickRef
    CancelOperation As Boolean
    DBPassWord As String
    DBFileName As String
    DBTimeOut As Long
    UserINIFileName As String
    GlobalINIFileName As String
End Type

Public Const HWND_TOPMOST = -1&
Public Const HWND_NOTOPMOST = -2&
Public Const SWP_NOSIZE = &H1&
Public Const SWP_NOMOVE = &H2&
Public Const SWP_NOACTIVATE = &H10&
Public Const SWP_SHOWWINDOW = &H40&

Public Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'INI File Functions...
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Colors...
Global Colors As tColors

Type tColors
    MainTextColor As Long
    MainTitleColor As Long
    MainBarColor As Long
    DesktopTextColor As Long
    DesktopTitleColor As Long
    DesktopBarColor As Long
    StatsTextColor As Long
    StatsTitleColor As Long
    UpdateColors As Boolean
End Type

Global IconToUse As String

Sub Main()
    
    QuickRef.UserINIFileName = App.Path & "\settings.Ini"
    
    FrmMain.Show
End Sub

Function ReadINI(sSection As String, sKeyName As String, sINIFileName As String) As String
    
    On Local Error Resume Next
    
    Dim sRet As String
    
    sRet = String(255, Chr(0))
    
    'Note: INI Filename can point to a local ini file or a remote ini file...
    ReadINI = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))

End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean
    
    On Local Error Resume Next
    
    Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
    
    WriteINI = (Err.Number = 0)

End Function

Sub SetColors(Who As Form)

    'On Local Error Resume Next
    
    If Who.Name = "FrmColors" Then
        FrmColors.Label2.ForeColor = Colors.MainTitleColor
        FrmColors.lblLabelColor.ForeColor = Colors.MainTextColor
        FrmColors.Label1.ForeColor = Colors.MainBarColor
        FrmColors.Label6.ForeColor = Colors.StatsTitleColor
        FrmColors.Label8.ForeColor = Colors.StatsTextColor
        FrmColors.Label11.ForeColor = Colors.DesktopTitleColor
        FrmColors.Label14.ForeColor = Colors.DesktopTextColor
        FrmColors.Label13.ForeColor = Colors.DesktopBarColor
    End If
    
    If Who.Name = "FrmDesktop" Then
        FrmDesktop.Line1.BorderColor = Colors.DesktopBarColor
        FrmDesktop.Line2.BorderColor = Colors.DesktopBarColor
        FrmDesktop.Line3.BorderColor = Colors.DesktopBarColor
        FrmDesktop.Line4.BorderColor = Colors.DesktopBarColor
        FrmDesktop.Line5.BorderColor = Colors.DesktopBarColor
        FrmDesktop.Line6.BorderColor = Colors.DesktopBarColor
        FrmDesktop.Line7.BorderColor = Colors.DesktopBarColor
        FrmDesktop.Line8.BorderColor = Colors.DesktopBarColor
        FrmDesktop.Line9.BorderColor = Colors.DesktopBarColor
        FrmDesktop.lblType.ForeColor = Colors.DesktopTitleColor
        FrmDesktop.Label1.ForeColor = Colors.DesktopTextColor
        FrmDesktop.Label2.ForeColor = Colors.DesktopTextColor
        FrmDesktop.Label3.ForeColor = Colors.DesktopTextColor
        FrmDesktop.Label4.ForeColor = Colors.DesktopTextColor
        FrmDesktop.Label5.ForeColor = Colors.DesktopTextColor
        FrmDesktop.Label6.ForeColor = Colors.DesktopTextColor
        FrmDesktop.Label7.ForeColor = Colors.DesktopTextColor
        FrmDesktop.Label8.ForeColor = Colors.DesktopTextColor
        FrmDesktop.Label9.ForeColor = Colors.DesktopTextColor
        FrmDesktop.Label10.ForeColor = Colors.DesktopTextColor
        FrmDesktop.lblRecv.ForeColor = Colors.DesktopTextColor
        FrmDesktop.lblSent.ForeColor = Colors.DesktopTextColor
    End If
    
    If Who.Name = "FrmMain" Then
        FrmMain.lblCaption.ForeColor = Colors.MainTitleColor
        FrmMain.Line1.BorderColor = Colors.MainBarColor
        FrmMain.Line2.BorderColor = Colors.MainBarColor
        FrmMain.Line3.BorderColor = Colors.MainBarColor
        FrmMain.Line4.BorderColor = Colors.MainBarColor
        FrmMain.Line5.BorderColor = Colors.MainBarColor
        FrmMain.Line6.BorderColor = Colors.MainBarColor
        FrmMain.Line7.BorderColor = Colors.MainBarColor
        FrmMain.Line8.BorderColor = Colors.MainBarColor
        FrmMain.Line9.BorderColor = Colors.MainBarColor
        FrmMain.lblType.ForeColor = Colors.MainTextColor
        FrmMain.Label1.ForeColor = Colors.MainTextColor
        FrmMain.Label2.ForeColor = Colors.MainTextColor
        FrmMain.Label3.ForeColor = Colors.MainTextColor
        FrmMain.Label4.ForeColor = Colors.MainTextColor
        FrmMain.Label5.ForeColor = Colors.MainTextColor
        FrmMain.Label6.ForeColor = Colors.MainTextColor
        FrmMain.Label7.ForeColor = Colors.MainTextColor
        FrmMain.Label8.ForeColor = Colors.MainTextColor
        FrmMain.Label9.ForeColor = Colors.MainTextColor
        FrmMain.Label10.ForeColor = Colors.MainTextColor
        FrmMain.Label11.ForeColor = Colors.MainTextColor
        FrmMain.Label12.ForeColor = Colors.MainTextColor
        FrmMain.lblRecv.ForeColor = Colors.MainTextColor
        FrmMain.lblSent.ForeColor = Colors.MainTextColor
    End If
    
    If Who.Name = "FrmStats" Then
        FrmStats.lblCaption.ForeColor = Colors.StatsTitleColor
        FrmStats.Label1(0).ForeColor = Colors.StatsTextColor
        FrmStats.Label1(1).ForeColor = Colors.StatsTextColor
        FrmStats.Label1(2).ForeColor = Colors.StatsTextColor
        FrmStats.Label1(3).ForeColor = Colors.StatsTextColor
        FrmStats.Label1(4).ForeColor = Colors.StatsTextColor
        FrmStats.ListView1.ForeColor = Colors.StatsTextColor
        FrmStats.ListView2.ForeColor = Colors.StatsTextColor
        FrmStats.ListView3.ForeColor = Colors.StatsTextColor
        FrmStats.ListView4.ForeColor = Colors.StatsTextColor
        FrmStats.ListView5.ForeColor = Colors.StatsTextColor
    End If
End Sub

Sub LoadColors()
    
    'Loads user specific colors for the application...
    
    On Local Error Resume Next
    
    Colors.MainTextColor = Val(ReadINI("colors", "maintextcolors", App.Path & "\settings.ini"))
    Colors.MainTitleColor = Val(ReadINI("colors", "maintitlecolors", App.Path & "\settings.ini"))
    Colors.MainBarColor = Val(ReadINI("colors", "mainbarcolors", App.Path & "\settings.ini"))
    Colors.DesktopTextColor = Val(ReadINI("colors", "desktoptextcolors", App.Path & "\settings.ini"))
    Colors.DesktopTitleColor = Val(ReadINI("colors", "desktoptitlecolors", App.Path & "\settings.ini"))
    Colors.DesktopBarColor = Val(ReadINI("colors", "desktopbarcolors", App.Path & "\settings.ini"))
    Colors.StatsTitleColor = Val(ReadINI("colors", "statstitlecolors", App.Path & "\settings.ini"))
    Colors.StatsTextColor = Val(ReadINI("colors", "statstextcolors", App.Path & "\settings.ini"))

End Sub
