Attribute VB_Name = "modMain"
Option Explicit

Public Sub CreateThumbnails(strSourceDir As String, strDestinationDir As String, Optional intType As FileType = JPG, _
               Optional MaxWidth As Integer = 50, Optional MaxHeight As Integer = 50, Optional bForceDimensions As Boolean = False, _
               Optional lBackColour As Long = &HFFFFFF)
    Dim colThumbFiles As Collection
    Dim aFilename   As String
    Dim i           As Long
    Dim progX       As Long
    Dim strDir      As String
    
    Set colThumbFiles = New Collection
    If right(strSourceDir, 1) <> "\" Then strSourceDir = strSourceDir & "\"
    strDir = Dir(strSourceDir & "*.jpg")
    Do Until strDir = ""
        i = i + 1
        If LCase(right(strDir, 3)) = "jpg" Then
            colThumbFiles.Add strDir
        End If
        strDir = Dir
    Loop
    
    For i = 1 To colThumbFiles.Count
        ' Create the thumbnail image
        createThumbnail colThumbFiles(i), strSourceDir, strDestinationDir, intType, MaxWidth, MaxHeight, bForceDimensions, lBackColour
        ' Don't lock up the application
        DoEvents
    Next i
End Sub

Public Sub createThumbnail(aFilename As String, strSource As String, strDestination As String, iType As FileType, _
               intMaxWidth As Integer, intMaxHeight As Integer, ForceDims As Boolean, BackColour As Long)
    Dim aThumb          As cThumbnail
    Dim aDestFileName   As String
    Dim aDestPath       As String
    Dim aSourcePath     As String
    
    ' Get the source path
    aSourcePath = Trim$(strSource)
    If right$(aSourcePath, 1) <> "\" Then aSourcePath = aSourcePath & "\"
    aSourcePath = aSourcePath & aFilename
    ' create the dest path
    aDestFileName = left$(aFilename, InStr(aFilename, ".") - 1)
    aDestPath = Trim$(strDestination)
    If right$(aDestPath, 1) <> "\" Then aDestPath = aDestPath & "\"
    aDestPath = aDestPath & "thumb" & aDestFileName
    Select Case iType
        Case BMP
            aDestPath = aDestPath & ".bmp"
        Case JPG
            aDestPath = aDestPath & ".jpg"
    End Select

    ' Create a new thumbnail image
    Set aThumb = New cThumbnail
    aThumb.MaxWidth = CLng(intMaxWidth)
    aThumb.MaxHeight = CLng(intMaxHeight)
    aThumb.BackColor = BackColour
    aThumb.ForcedDimensions = ForceDims
    ' Set our destination path
    aThumb.DestFilePath = aDestPath
    ' Create the thumbnail file!
    aThumb.CreateFromFile aSourcePath

    Set aThumb = Nothing

End Sub

