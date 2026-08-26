VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove Text"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Remove all lines from the given file that contain the given text
' Wildcards may be used.
' For example,
' remtext.exe "blah" "c:\temp\blah.txt"
' remove all lines that are blah

' remtext.exe "*blah" "c:\temp\blah.txt"
' remove all lines that end in blah

' remtext.exe "blah*" "c:\temp\blah.txt"
' remove all lines that start with blah

' remtext.exe "*blah*" "c:\temp\blah.txt"
' remove all lines that contain blah

' remtext.exe "blah?" "c:\temp\blah.txt"
' remove all lines that contain blah and another character - eg blahs will be removed.

' Usage:  remtext.exe "stringtosearchfor" "filename"
Private Sub Form_Load()
    Dim strSearch As String
    Dim strFilename As String
    Dim strInputText As String
    Dim strOutputText As String
    Dim strCommandline As String
    Dim iTemp As Integer
    Dim sTemp As String
    
    strCommandline = Trim(Command$)
    iTemp = InStr(2, strCommandline, Chr(34), vbTextCompare)
    
    strSearch = Mid(strCommandline, 2, iTemp - 2)
    
    strFilename = Replace(Trim(Mid(strCommandline, iTemp + 1)), Chr(34), "", 1, -1, vbTextCompare)
    
    Open strFilename For Input As #1
        Do Until EOF(1)
            Line Input #1, strInputText
            If strInputText Like strSearch Then
            Else
                strOutputText = strOutputText & strInputText & vbCrLf
            End If
        Loop
    Close 1
    
    If Len(strOutputText) > 2 Then
        strOutputText = Left(strOutputText, Len(strOutputText) - 1)
        Open strFilename For Output As #1
            Print #1, strOutputText
        Close 1
    End If
    
    End
End Sub
