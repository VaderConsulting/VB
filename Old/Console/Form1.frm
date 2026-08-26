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
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim BytesWritten As Long
    Dim CodePage As Long
    Dim ScreenBufferHandle As Long
    CodePage = &H352   ' &H352 = 850
    Dim TextToPrint As String
    
    AllocConsole
    ScreenBufferHandle = CreateConsoleScreenBuffer(GENERIC_WRITE, FILE_SHARE_WRITE + FILE_SHARE_READ, Null, CONSOLE_TEXTMODE_BUFFER, Null)
    Debug.Print "Screen Buffer Handle :" & ScreenBufferHandle
    SetConsoleActiveScreenBuffer (ScreenBufferHandle)
    
    InputHandle = GetStdHandle(STD_INPUT_HANDLE)
    OuputHandle = GetStdHandle(STD_OUTPUT_HANDLE)
    ErrorHandle = GetStdHandle(STD_ERROR_HANDLE)
    Debug.Print "Input Handle:" & InputHandle
    Debug.Print "Output Handle:" & OutputHandle
    Debug.Print "Error Handle:" & ErrorHandle
    ConsoleMode = SetConsoleMode(ScreenBufferHandle, ENABLE_ECHO_INPUT + ENABLE_LINE_INPUT + ENABLE_PROCESSED_INPUT)
    Debug.Print "Console Mode returns: " & ConsoleMode
    
    SetConsoleOutputCP CodePage
    SetConsoleCP CodePage
    
    TextToPrint = "Hello World"
    WritingToConsole = WriteConsole(ScreenBufferHandle, TextToPrint, Len(TextToPrint), BytesWritten, Null)
    If WritingToConsole <> 0 Then
        ConsoleResult = "Success"
    Else
        ConsoleResult = "Failure"
    End If
    Debug.Print "Writing to Console returned: " & WritingToConsole & " (" & ConsoleResult & ")"
    Debug.Print "Bytes Written: " & BytesWritten
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    FreeConsole
End Sub
