Attribute VB_Name = "Messages"
Public NewLogFile As String
Public TestModeLog  As String

Public Const msgEmptyFields = "You need to specify at list 1 file to replicate, " & _
                              "and destination host name."
                              
Public Const msgCopyTree = "You are about to perform very critical operation on your network!" & vbCr & _
                            "Please make sure you have selected right files, hosts, and destination," & vbCr & _
                            "as well as correct Synchronization mode." & vbCr & _
                            "Click Cancel to Stop." & vbCr & _
                            "Click OK to continue."

Public Const msgDelete = "You are about to delete files listed in a  ''Files to Replicate''  window " & _
                         "on the remote hosts!" & _
                         vbCr & "Click  ''Cancel''  to Stop." & vbCr & "Click  ''OK'' to " & _
                         "continue." & vbCr & vbCr & "Note: You can not Delete folders, or read-only files with this " & _
                         "function. " & vbCr & "Use  ''Purge Folder''  or  ''Mirror Folder''  with " & _
                         " ''Incl Subfolders''  checked. See Help for more " & _
                         "information."
                         
Public Const msgEmptyDest = "Destination Folder is not specified! Files might be deleted " & _
                            "at the root of the share." & vbCr & "Click Cancel to stop!"
                            
Public Const msgCopyTreeError = "''Replicate tree'' will function if only ONE entry exists in " & _
                                "''Files to Replicate'' window." & vbCr & "To specify a tree use  ''*.*''  " & _
                                "wildcard, or use  ''Add Tree''." & vbCr & vbCr & "Example:   " & _
                                "''C:\test\Scripts\*.*''"

Public Function CreateLog(ByVal lFile As String, ByVal rMode As String)
    cName = Environ$("COMPUTERNAME")
    Open lFile For Output As #1
    Print #1, "                                     " & rMode & "  Log  on   " & "''" & "\\" & cName & "''"
    Print #1, "============================================================="
    Close #1
End Function

Public Function OpenReplicatorLog()
    Open NewLogFile For Append As #1
    
    Print #1, vbCrLf & "=================="
    Print #1, Now
    Print #1, "==================" & vbCrLf
    Print #1, "Replication of the following file(s):" & vbCrLf
    
    J = 0
    Do
        Print #1, Trim(EReplicator.FileList.List(J))
        J = J + 1
    Loop Until J > EReplicator.FileList.ListCount - 1
    
    Print #1, vbCrLf & "to  ''" & EReplicator.RShare.Text & EReplicator.DestFolder.Text & "''   failed on the following hosts:" & vbCrLf
End Function

Public Function OpenTestModeLog()
    Open TestModeLog For Append As #7
    Print #7, vbCrLf & "=================="
    Print #7, Now
    Print #7, "==================" & vbCrLf
    Print #7, "For the list of host access errors see Replicator Log."
    Print #7, "The Following ''File Operations'' would take place:" & vbCrLf
End Function

Public Function ShowLogs(ByVal LogFile As String)
    On Error GoTo ErrorHandler
    Open LogFile For Input As #1
    File_Length = LOF(1)
    Read_Buffer = Input(File_Length, #1)
    EReplicator.LogPath.Caption = LogFile
    EReplicator.Log.Text = Read_Buffer
    Close #1
    Exit Function
ErrorHandler:
    LogError = MsgBox("Can not open log file " & LogFile, vbCritical, "Error")
    Err.Clear
    Close #1
End Function

Public Function TestReplicateFiles(ByVal absPath As String, ByVal absDest As String, ByVal Override As Boolean, ByVal deleteFiles As Boolean) As Long
    
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
                    Print #7, "Copy File From:  ''" & fPath & Left$(firstFile, Len(firstFile) - 1) & "'' To: ''" & absDest & Left$(firstFile, Len(firstFile) - 1) & "''"
                ElseIf deleteFiles Then
                    If FileExists(absDest & firstFile) Then Print #7, "Delete File: ''" & absDest & Left$(firstFile, Len(firstFile) - 1) & "''"
                End If
            End If
            nextFile = FindNextFile(getFile, FindData)
        Loop Until nextFile = 0
    End If
    Call FindClose(getFile)
End Function

Public Function TestReplicateTree(ByVal absPath As String, ByVal absDest As String, ByVal Override As Boolean) As Long
                             
    Dim queryFolder     As String
    Dim subFolderArray  As New Collection
    Dim subFolder       As Variant
    
    queryFolder = Dir(absPath & "*.*", vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory Or vbArchive)
    Do Until queryFolder = ""
        If queryFolder = "." Or queryFolder = ".." Then
        ElseIf (GetAttr(absPath & queryFolder) And vbDirectory) = vbDirectory Then
            If Not FileExists(absDest & queryFolder) Then Print #7, "Create Folder: ''" & absDest & queryFolder & "''"
                subFolderArray.Add queryFolder & "\"
            Else
            Print #7, "Copy File From:  ''" & absPath & queryFolder & "'' To: ''" & absDest & queryFolder & "''"
        End If
        queryFolder = Dir
    Loop
    
    For Each subFolder In subFolderArray
    Call TestReplicateTree(absPath & CStr(subFolder), absDest & CStr(subFolder), Override)
    Next
End Function

Public Function TestPurgeFolder(ByVal absPath As String, ByVal absDest As String, ByVal inclSubfolders As Boolean) As Long
                             
    Dim queryDestFolder     As String
    Dim subDestFolderArray  As New Collection
    Dim subFolder           As Variant
    
    queryDestFolder = Dir(absDest & "*.*", vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory Or vbArchive)
    Do Until queryDestFolder = ""
        If queryDestFolder = "." Or queryDestFolder = ".." Then
        ElseIf (GetAttr(absDest & queryDestFolder) And vbDirectory) = vbDirectory Then
            If inclSubfolders Then
                If Not FileExists(absPath & queryDestFolder) Then
                    Print #7, "Delete Folder:   ''" & absDest & queryDestFolder & "''  -  Including Subfolders, if Any."
                Else
                    subDestFolderArray.Add queryDestFolder & "\"
                End If
            End If
            ElseIf Not FileExists(absPath & queryDestFolder) Then Print #7, "Delete File:   ''" & absDest & queryDestFolder & "''"
        End If
        queryDestFolder = Dir
    Loop
    If Not inclSubfolders Then Exit Function
    For Each subDestFolder In subDestFolderArray
        Call TestPurgeFolder(absPath & CStr(subDestFolder), absDest & CStr(subDestFolder), inclSubfolders)
    Next
End Function

