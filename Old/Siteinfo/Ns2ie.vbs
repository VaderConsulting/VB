' *************************************************************
' NS2IE.vbs
' Netscape Bookmark to Internet Explorer Favorite Converter
' (c) 2002 D. Robinson
' *************************************************************

Option Explicit
Dim Fav_Folder
Dim WshShell
Dim WshSysEnv
Dim fso
Dim fldr
Dim f1
Dim InFile
Dim OutFile
Dim TextStream
Dim ModFav_Folder
Dim Netscape_Favourites_File
Dim Netscape_Folder
Dim ForReading
Dim ForWriting
Dim ForAppending
Dim LPos
Dim RPos
Dim URL
Dim Name

ForReading =   1
ForWriting =   2
ForAppending = 8

'
' Netscape 4 uses bookmark.htm
' Netscape 6 uses bookmarks.html
'
'
' NOTE: Netscape 6 does not use a normal (ie VB) CRLF for its EOL character,
' so this script will NOT work in its current state
'
Netscape_Favourites_File = "bookmark.htm"
Netscape_Folder = "H:\"

Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("process")

' Get location of favourites: your choice of 2 methods
Fav_Folder = WshShell.SpecialFolders("Favorites")
If Fav_Folder = "" Then
    Fav_Folder = WshSysEnv("USERPROFILE") & "\Favorites"
End If

' The filesystem object requires \'s to be represented as \\
Fav_Folder = replace(Fav_Folder,"\","\\")

Set fso = CreateObject("Scripting.FileSystemObject")
'Set fldr = fso.GetFolder("Fav_Folder")

' Favourites are .url files, so create a .url file for each Netscape Bookmark
Set InFile = fso.OpenTextFile(Netscape_Folder & Netscape_Favourites_File, ForReading)

Do
    TextStream = InFile.ReadLine
Loop Until (Instr(1,TextStream,"Personal Bookmarks") > 0) Or InFile.AtEndOfStream

Do
    TextStream = InFile.ReadLine
    If Instr(1,TextStream,"<DT>") > 0 Then
	LPos = Instr(1,LCase(TextStream),"href=") + 6
        RPos = Instr(LPos,LCase(TextStream),chr(34))
        URL = Mid(TextStream,LPos,RPos - LPos)
        URL = Replace(URL,"\","/") ' Cannot have \, so use / instead
        
        LPos = Instr(1,LCase(TextStream),"last_modified=") + 15
        LPos = Instr(LPos,LCase(TextStream),">") + 1
        RPos = Instr(LPos,LCase(TextStream),"<")
        Name = Mid(TextStream,LPos,RPos - LPos)
        'WScript.echo Name & " = " & URL
    End If

    If Name <> "" Then
        Set OutFile = fso.CreateTextFile(Fav_Folder & "\" & Name & ".url", True)
        OutFile.Write ("[DEFAULT]") & Chr(13) & chr(10)
        OutFile.Write ("BASEURL=" & URL) & Chr(13) & chr(10)
        OutFile.Write ("[InternetShortcut]") & Chr(13) & chr(10)
        OutFile.Write ("URL=" & URL) & Chr(13) & chr(10)
        OutFile.Close
    End If
Loop Until InFile.AtEndOfStream
