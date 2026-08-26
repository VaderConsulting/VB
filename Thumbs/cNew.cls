VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Use to create thumbnails."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum FileType
    BMP = 0
    JPG = 1
End Enum

'set this to 0 to disable debug code in this class
#Const DebugMode = 1


#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
Public Sub Thumbnail(aFilename As String, strSource As String, strDestination As String, Optional iType As FileType = JPG, Optional intMaxWidth As Integer = 50, Optional intMaxHeight As Variant = 50, Optional bForceDimensions As Boolean = False, Optional lBackColour As Long = &HFFFFFF)
Attribute Thumbnail.VB_Description = "Use to create a single thumbnail from a single file"
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
    aThumb.BackColor = lBackColour
    aThumb.ForcedDimensions = bForceDimensions
    ' Set our destination path
    aThumb.DestFilePath = aDestPath
    ' Create the thumbnail file!
    aThumb.CreateFromFile aSourcePath

    Set aThumb = Nothing
End Sub

#End If

Public Sub Thumbnails(strSourceDir As String, strDestinationDir As String, Optional intType As FileType = JPG, Optional MaxWidth As Integer = 50, Optional MaxHeight As Integer = 50, Optional bForceDimensions As Boolean = False, Optional lBackColour As Long = &HFFFFFFF)
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
        Thumbnail colThumbFiles(i), strSourceDir, strDestinationDir, intType, MaxWidth, MaxHeight, bForceDimensions, lBackColour
        ' Don't lock up the application
        DoEvents
    Next i
End Sub



Private Sub Class_Initialize()
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub


#If DebugMode Then
    Public Property Get ClassDebugID()
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

