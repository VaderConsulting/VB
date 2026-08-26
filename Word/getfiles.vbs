Set FSO = CreateObject("Scripting.FileSystemObject")

ShowFiles FSO.GetFolder("C:\Temp"),""

Sub ShowFiles(Folder, Padding)
    For Each File in Folder.Files
        wscript.echo padding & File.name
    Next
    For Each Subfolder in Folder.SubFolders
        Wscript.Echo padding & Subfolder.Path
        ShowFiles Subfolder, "    "
    Next
    Padding = Padding + "    "
End Sub