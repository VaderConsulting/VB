Attribute VB_Name = "fileState"
Public Container    As IADsContainer
Public FileToOpen   As String
Public TestMode     As Boolean

Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH = 260
Public Const SHGFI_DISPLAYNAME = &H200

Public Const FO_MOVE = &H1
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FOF_SILENT = &H4
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_NOERRORUI = &H400
Public Const FOF_ALLOWUNDO = &H40

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Public Type SHFILEOPSTRUCT
    hWnd               As Long
    wFunc              As Long
    pFrom              As String
    pTo                As String
    fFlags             As Integer
    fAborted           As Boolean
    hNameMaps          As Long
    sProgress          As String
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime   As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime  As FILETIME
   nFileSizeHigh    As Long
   nFileSizeLow     As Long
   dwReserved0      As Long
   dwReserved1      As Long
   cFileName        As String * MAX_PATH
   cAlternate       As String * 14
End Type

Public Type SECURITY_ATTRIBUTES
    nLength             As Long
    pSecurityDescriptor As Long
    bInheritHandle      As Long
End Type

Public Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Public Function FileExists(sFile As String) As Boolean
    
    Dim shfi As SHFILEINFO
    
    If SHGetFileInfo(ByVal sFile, 0&, shfi, Len(shfi), SHGFI_DISPLAYNAME) Then
        FileExists = True
        Else
        FileExists = False
    End If
End Function

Public Function FileOperation(ByVal absPath As String, ByVal absDest As String, fOperation As Long) As Long
    
    Dim FOF_FLAGS As Long
    Dim SHFileOp As SHFILEOPSTRUCT
    
    FOF_FLAGS = 0&
    FOF_FLAGS = FOF_FLAGS Or FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOERRORUI Or FOF_ALLOWUNDO
    absPath = absPath & Chr$(0) & Chr$(0)
    With SHFileOp
        .wFunc = fOperation
        .pFrom = absPath
        .pTo = absDest
        .fFlags = FOF_FLAGS
    End With
    SHFileOperation SHFileOp
End Function

Public Function ReplicateFiles(ByVal absPath As String, ByVal absDest As String, ByVal Override As Boolean, ByVal deleteFiles As Boolean) As Long

    Dim FindData As WIN32_FIND_DATA
    Dim SA          As SECURITY_ATTRIBUTES
    Dim getFile     As Long
    Dim nextFile    As Long
    Dim firstFile   As String
    Dim fPath       As String
    
    fPath = Left$(absPath, InStrRev(absPath, "\"))
    getFile = FindFirstFile(absPath, FindData)
    If getFile Then
        Do
            firstFile = Left$(FindData.cFileName, InStr(FindData.cFileName, Chr$(0)))
            If Not (GetAttr(fPath & firstFile) And vbDirectory) = vbDirectory Then
                If Not deleteFiles Then
                    Call CopyFile(fPath & firstFile, absDest & firstFile, Override)
                ElseIf deleteFiles Then
                    If FileExists(absDest & firstFile) Then Call DeleteFile(absDest & firstFile)
                End If
            End If
            nextFile = FindNextFile(getFile, FindData)
        Loop Until nextFile = 0
    End If
    Call FindClose(getFile)
End Function

Public Function ReplicateTree(ByVal absPath As String, ByVal absDest As String, ByVal Override As Boolean) As Long
                             
    Dim queryFolder     As String
    Dim subFolderArray  As New Collection
    Dim subFolder       As Variant

    queryFolder = Dir(absPath & "*.*", vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory Or vbArchive)
    Do Until queryFolder = ""
        If queryFolder = "." Or queryFolder = ".." Then
        ElseIf (GetAttr(absPath & queryFolder) And vbDirectory) = vbDirectory Then
            If Not FileExists(absDest & queryFolder) Then MkDir absDest & queryFolder
            subFolderArray.Add queryFolder & "\"
        Else
            Call CopyFile(absPath & queryFolder, absDest & queryFolder, Override)
        End If
        queryFolder = Dir
    Loop
    
    For Each subFolder In subFolderArray
        Call ReplicateTree(absPath & CStr(subFolder), absDest & CStr(subFolder), Override)
    Next
End Function

Public Function PurgeFolder(ByVal absPath As String, ByVal absDest As String, ByVal inclSubfolders As Boolean) As Long
                             
    Dim queryDestFolder     As String
    Dim subDestFolderArray  As New Collection
    Dim subFolder           As Variant
                                 
    queryDestFolder = Dir(absDest & "*.*", vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory Or vbArchive)
    Do Until queryDestFolder = ""
        If queryDestFolder = "." Or queryDestFolder = ".." Then
            ElseIf (GetAttr(absDest & queryDestFolder) And vbDirectory) = vbDirectory Then
                If inclSubfolders Then
                    If Not FileExists(absPath & queryDestFolder) Then
                        Call FileOperation(absDest & queryDestFolder, "", FO_DELETE)
                    Else
                        subDestFolderArray.Add queryDestFolder & "\"
                    End If
                End If
            ElseIf Not FileExists(absPath & queryDestFolder) Then Call FileOperation(absDest & queryDestFolder, "", FO_DELETE)
        End If
        queryDestFolder = Dir
    Loop
    If Not inclSubfolders Then Exit Function
    For Each subDestFolder In subDestFolderArray
        Call PurgeFolder(absPath & CStr(subDestFolder), absDest & CStr(subDestFolder), inclSubfolders)
    Next
End Function

Public Function OpenFile(FileToOpen As String)

    Dim s           As Integer
    Dim I           As Integer
    Dim J           As Integer
    Dim fStream     As String
    Dim fldStream   As String
    Dim frmChk      As String
    Dim hStream     As String
    Dim skip1       As String
    Dim flChk       As String
    Dim shStream    As String
    Dim shareAdd    As Boolean
    
    On Error Resume Next
    With EReplicator
        .FLocation.Text = ""
        .FileList.Clear
        .DestFolder.Text = ""
        .List1.Clear
        .ManualAdd.Text = ""
        .ComChar.Text = ""
        .SConn.Clear
        .FConn.Clear
        .HCount.Caption = "0"
    End With
    
    Open FileToOpen For Input As #1
    Line Input #1, skip1
    If Not skip1 = "[Files]" Then
        MsgBox "This is not a valid session file, or file is corrupted!", vbCritical, "Error:"
        Close #1
        EReplicator.Caption = "Replicator 1-2-3 (Untitled.sss)"
        Exit Function
    End If
    
    Do While Not EOF(1)
        Line Input #1, fStream
        If fStream = "[Share]" Then Exit Do
        If fStream = "[Folder]" Then
            MsgBox "File Format is " & "Invalid.", vbCritical, "Error:"
            Close #1
            Exit Function
        End If
        If FileExists(Trim(fStream)) Then EReplicator.FileList.AddItem Trim(fStream)
    Loop
    
    Line Input #1, shStream
    EReplicator.RShare.Text = Trim(shStream)
    shareAdd = True
    s = 0
    Do
        If EReplicator.RShare.List(s) = Trim(shStream) Then shareAdd = False
        s = s + 1
    Loop Until s > EReplicator.RShare.ListCount - 1
    If shareAdd Then EReplicator.RShare.AddItem (shStream)
    
    Line Input #1, flChk
    If Not flChk = "[Folder]" Then
        MsgBox "File Format is " & "Invalid.", vbCritical, "Error:"
        Close #1
        Exit Function
    End If
    Line Input #1, fldStream
    EReplicator.DestFolder.Text = Trim(fldStream)
    Line Input #1, frmChk
    If Not frmChk = "[Hosts]" Then
        MsgBox "File Format is " & "Invalid.", vbCritical, "Error:"
        Close #1
        Exit Function
    End If
    Do While Not EOF(1)
       Line Input #1, hStream
       EReplicator.List1.AddItem Trim(hStream)
    Loop
    Close #1
    
    I = 0
    Do
        If EReplicator.List1.List(I) = "" Then
            EReplicator.List1.RemoveItem (I)
        Else
            I = I + 1
        End If
    Loop Until I > EReplicator.List1.ListCount - 1
    
    J = 0
    Do
        If EReplicator.FileList.List(J) = "" Then
            EReplicator.FileList.RemoveItem (J)
        Else
            J = J + 1
        End If
    Loop Until J > EReplicator.FileList.ListCount - 1
    If FileListStatus Then EReplicator.chkCopyTree.Value = 0
    EReplicator.HCount.Caption = EReplicator.List1.ListCount
End Function

Public Function SaveFile(FileToSave As String)

    Dim I       As Integer
    Dim J       As Integer
    
    On Error Resume Next
    Open FileToSave For Output As #1
    Print #1, "[Files]"
    I = 0
    Do
        Print #1, EReplicator.FileList.List(I)
        I = I + 1
    Loop Until I > EReplicator.FileList.ListCount - 1
    Print #1, "[Share]"
    Print #1, EReplicator.RShare.Text
    Print #1, "[Folder]"
    Print #1, EReplicator.DestFolder.Text
    Print #1, "[Hosts]"
    J = 0
    Do
        Print #1, EReplicator.List1.List(J)
        J = J + 1
    Loop Until J > EReplicator.List1.ListCount - 1
    Close #1
End Function

Public Function FileListStatus() As Boolean
    If EReplicator.FileList.ListCount <> 1 Or Not Right(Trim(EReplicator.FileList.List(0)), 3) = "*.*" Then FileListStatus = True
End Function

Public Function MakeTree(FolderIn, HostShare)
    Dim NextFolder As String
    Dim SA As SECURITY_ATTRIBUTES
    Dim FolderArray As Variant
    
    FolderArray = Split(FolderIn, "\", -1, vbTextCompare)
    NextFolder = Left(HostShare, Len(HostShare) - 1)
    For Each subFolder In FolderArray
        NextFolder = NextFolder & "\" & subFolder
        If Not FileExists(NextFolder) Then
            createSuccess = CreateDirectory(NextFolder, SA)
        End If
    Next
    If createSuccess = 0 Then
        MakeTree = False
    Else
        MakeTree = True
    End If
End Function
