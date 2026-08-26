VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set DRIMS Library"
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblInfo 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 03-04-2003 13:05
' Author    : Dave Robinson
' Purpose   : Start of application
'             Accepts one command line argument - 'ts'.  This will force Terminal Server mode
'
'  V    Date        Author          History
' 1.0   03-04-2003  Dave Robinson   Initial Version
' 1.1   07-04-2003  Dave Robinson   Added modRegistry, and modified SetDRIMS to use DelRegValue
' 1.2   07-04-2003  Dave Robinson   Added code for DRIMS RM
' 1.3   07-04-2003  Dave Robinson   Removed group membership dependency.  Uses Registry instead.  HKCU\Software\Woodside\DocsOpen\Library
' 1.4   11-04-2003  Dave Robinson   Added Visio, Project and Powerpoint registry keys to add/delete in Task3/Task6.
' 1.5   11-04-2003  Dave Robinson   Added Visio, Project and Powerpoint registry keys to add/delete in Disable/Enable Integration.
' 1.6   11-04-2003  Dave Robinson   Removed addition from 1.5, and removed/copied odma32.dll in Disable/Enable Integration.
'---------------------------------------------------------------------------------------
Private Sub Form_Load()

  Dim bInGroup As Boolean
  Dim strGroupname As String
  Dim bSetResult As Boolean
  Dim strLibraryName As String
  Dim isTerminalServer As Boolean

    bReportErrors = True
    bReportWarnings = True
    bReportInformation = True
    bReportToEventLog = True

    bDebugMode = True

    ' Is this a terminal Server session?
    isTerminalServer = (Environ$("CLIENTNAME") <> "" And LCase$(Environ$("CLIENTNAME")) <> "console") ' True = YES:  Terminal Server Session

    If LCase$(InStr(1, Command$, "ts")) <> 0 Then isTerminalServer = True ':( Expand Structure

    LogEvent "isTerminalServer: " & isTerminalServer, DebugInfo

    strUsername = Environ$("USERNAME")
    strComputername = Environ$("COMPUTERNAME")

    If isTerminalServer Then
        frmSelectLibrary.Show vbModal
      Else 'ISTERMINALSERVER = FALSE/0
        Me.Show
        Me.Refresh

        ' Generate an App_Path variable that is the path to this application, similar to App.Path, though it will always end with a backslash.
  Dim strApp_Path As String ':( Move line to top of current Sub
        If Right$(App.Path, 1) <> "\" Then
            strApp_Path = App.Path & "\"
          Else 'NOT RIGHT$(APP.PATH,...
            strApp_Path = App.Path
        End If

        If InStr(1, LCase$(Command$), "debug") <> 0 Then bDebugMode = True ':( Expand Structure

        strLibraryName = GetRegString(HKEY_CURRENT_USER, "Software\Woodside\Docsopen", "Library", "")
        LogEvent "User has " & strLibraryName & " specified as the DocsOpen library", DebugInfo
        Select Case LCase$(Trim$(strLibraryName))
          Case "cdc"
            bSetResult = SetDRIMS(CDC)
          Case "opd"
            bSetResult = SetDRIMS(OPD)
          Case "sunrise"
            bSetResult = SetDRIMS(Sunrise)
          Case "ve"
            bSetResult = SetDRIMS(VincentEnfield)
          Case "drims"
            bSetResult = SetDRIMS(Standard)
          Case "drims rm"
            bSetResult = SetDRIMS(DRIMSRM)
          Case Else
            bSetResult = SetDRIMS(Disabled)
        End Select

        If bSetResult Then
            LogEvent "Settings DocsOpen library to " & strLibraryName & " was successful", Information
          Else 'BSETRESULT = FALSE/0
            LogEvent "Setting DocsOpen library to " & strLibraryName & " failed.", Warning
            MsgBox "Setting DocsOpen library to " & strLibraryName & " failed.", vbExclamation, "Warning"
        End If

    End If

    EndApp

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jul-28 12:26) 1 + 175 = 176 Lines
