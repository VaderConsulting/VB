Attribute VB_Name = "Module1"
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName  As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Type SHITEMID
    cb As Long
    abID As Byte
End Type

Type ITEMIDLIST
    mkid As SHITEMID
End Type

