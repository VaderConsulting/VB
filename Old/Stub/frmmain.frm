VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Installer Stub"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5580
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Stub.exe
'
' Used to ensure that the pre-requisites are installed for applications(s) before attempting to run them.
' This app should only use the standard VB run time libraries, and 0 additional OCX's.
'
' Implementation:
' Place Stub.ini and Stub.exe in the same directory, and execute Stub.exe.
' Of course stub.ini will at this stage already be modified to list the appropriate files.
' Stub.exe will copy all required files, and register them as necessary.

' Started 28 July 2001
' D. Robinson


'    Example Stub.ini file:

'    [Applications]
'    Names=Logon,App2,App3
'
'    [Logon]
'    Require=c:\file1.dll,c:\file2.ini,c:\file3.exe
'    Source=c:\
'
'    [App2]
'    Require=c:\file4.exe
'    Source=d:\
'
'    [App3]
'    Require=c:\file5.ocx
'    Source=e:\

' *******************************************TO DO:******************************************************
' Put file sizes and dates into Stub.ini to ensure the required file(s) are the correct versions.
' *******************************************************************************************************

    Option Explicit
    
    ' .ini file manipulation API call
    Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Sub Form_Load()
    Dim strPath As String
    Dim intFileNumber As Integer
    Dim lpReturnString As String
    Dim intValid As Integer, strApps() As String, intLoop As Integer, intLoop2 As Integer
    Dim strRequire() As String, strSourceDir As String
    
    ' Ensure Path ends in '\'
    strPath = App.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    If Dir(strPath & "Stub.ini", vbNormal) = "" Then EndProgram
    ' Proceed... Stub.ini found in same dir as Stub.exe
    
    lpReturnString = Space$(128)
    intValid = GetINIValue("Applications", "Names", "", lpReturnString, Len(lpReturnString), strPath & "Stub.ini")
    '* The return value (assigned to intValid) represents the number of characters read into lpReturnString.
    If intValid = 0 Then EndProgram
    '* Discard the trailing spaces and null character, and split App names at each comma.
    strApps() = Split(Left$(lpReturnString, intValid), ",")
    
    ' For each of the applications listed, get the list of required files, and the source for these files if not found
    For intLoop = 0 To UBound(strApps())
        lpReturnString = Space$(128)
        intValid = GetINIValue(strApps(intLoop), "Require", "", lpReturnString, Len(lpReturnString), strPath & "Stub.ini")
        
        strRequire() = Split(Left$(lpReturnString, intValid), ",")
        
        lpReturnString = Space$(128)
        intValid = GetINIValue(strApps(intLoop), "Source", "", lpReturnString, Len(lpReturnString), strPath & "Stub.ini")
        strSourceDir = Left$(lpReturnString, intValid)
        
        For intLoop2 = 0 To UBound(strRequire())
            ' Check if the required file requires registering
            If LCase(Right(strRequire(intLoop2), 3)) = "ocx" Or LCase(Right(strRequire(intLoop2), 3)) = "dll" Then ' yes
                Debug.Print strApps(intLoop) & " requires " & strRequire(intLoop2) & ".  Copy from " & strSourceDir & " if missing.  (must be registered)"
            Else ' no
                Debug.Print strApps(intLoop) & " requires " & strRequire(intLoop2) & ".  Copy from " & strSourceDir & " if missing."
            End If
        Next intLoop2
        
        ' Finished with array, discard contents
        Erase strRequire
    Next intLoop
End Sub

Function GetINIValue(lpAppName As String, lpKeyName As String, lpDefault As String, lpReturnString As String, intSize As Integer, lpFileName As String) As Integer
    GetINIValue = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, lpReturnString, intSize, lpFileName)
End Function
Sub EndProgram()
    ' Close app down gracefully
    Unload Me
    End
End Sub
