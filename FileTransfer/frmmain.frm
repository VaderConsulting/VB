VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ChunkSize As Long = 100000

Private Sub Form_Load()

TFileCopy "C:\Temp\Central.xml", "C:\Temp\Central2.xml"
'    Dim Length1 As Long
'    Dim Attributes1 As Integer
'    Dim SourceFile As String
'    Dim TempDirectory As String
'    Dim DestinationFile As String
'    Dim Chunk As String
'    Dim ThisChunkSize As Long
'    Dim lp As Long
'    Dim i As Integer
'
'    SourceFile = "C:\Temp\Central.xml"
'    TempDirectory = "C:\Temp\"
'
'    Length1 = FileLen(SourceFile)
'    Attributes1 = GetAttr(SourceFile)
'
'    'Debug.Print "Source length:     " & Length1
'    'Debug.Print "Source attributes: " & Attributes1
'
'    ' Remove existing temp file
'    If Dir(TempDirectory & "chunk.txt") <> "" Then Kill TempDirectory & "chunk.txt"
'    i = 0
'    For lp = 1 To Length1 Step ChunkSize
'        i = i + 1
'        'Debug.Print lp
'        ThisChunkSize = ChunkSize
'        If lp + ChunkSize > Length1 Then
'            ThisChunkSize = (Length1 - lp) + 1
'        End If
'
'        ' open the source file, read in ThisChunkSize bytes from lp position
'        Open SourceFile For Input As #1 Len = ThisChunkSize
'        ' Set the position to read from
'        Seek #1, lp
'        Chunk = Input(ThisChunkSize, 1)
'        ' write chunksize bytes to a file in the tempdirectory called chunk.txt
'        Open TempDirectory & "chunk.txt" For Append As #2 Len = ThisChunkSize
'            Print #2, Left(Chunk, ThisChunkSize)
'        Close 2
'        Close 1
'    Next lp
'
'    End
    
End Sub
