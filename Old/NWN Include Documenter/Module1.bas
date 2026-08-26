Attribute VB_Name = "Module1"
Option Explicit

Private Type Comment
    LineIsSingle As Boolean
    Start As String
    Finish As String
End Type

Const NoOfCommentTypes As Integer = 10
Const SINGLELINE As Boolean = True
Const MULTILINE As Boolean = False

Dim Comments(NoOfCommentTypes) As Comment

Const NWN_Dir As String = "c:\temp\NWN"
Const FileType As String = "*.NSS"

Public Sub main()
    Dim i As Integer, j As Integer
    Dim s As String
    Dim Comment() As String
    Dim Parameters() As String
    Dim Filename As String
    Dim LineNo As Integer
    Dim D As String
    Dim CommentBlockStarted As Boolean
    Dim CommentBlockEnded As Boolean
    Dim NWNFileData As String
    Dim FunctionsPerFile As Integer
    Dim FileLineNumber As Integer
    Dim FunctionLineNumber() As Integer
    Dim LeftPos As Integer
    Dim RightPos As Integer
    Dim NWNFileData_Part As String
    Dim FunctionName() As String
    Dim Returns() As String
    Dim Params() As String
    Dim v As Variant
    Dim ParamType() As String
    Dim isFunction As Boolean, DoExit As Boolean
    Dim CommentPart As String
    Dim Prototype As String
    Dim AllParams As String
    
    ' Add comment types we are interested in
    AddComment SINGLELINE, "//"
    AddComment MULTILINE, "/*", "*/"
    
    ' Start processing files
    D = Dir(NWN_Dir & "\" & FileType)
    
    ' Open file for output
    Open "C:\temp\nwn\NWNFunctions2.xml" For Output As #2
    
    ' Write out XML header
    Print #2, "<?xml version=""1.0"" ?> "
    Print #2, "<!-- "
    Print #2, "NeverWinter Nights Functions"
    Print #2, "Dave Robinson 27 July 2002"
    Print #2, "windows_mcp@hotmail.com"
    Print #2, ""
    Print #2, "27/07/02: Created - Exported from Bioware NWSCRIPT.NSS (v1.21)"
    Print #2, "28/07/02: Modified - Added extra function return types, and removed more code remarks"
    Print #2, "29/07/02: Modified - Added prototype"
    Print #2, "--> "
    Print #2, "<nwn_functions>"
    
    ' Loop through each file found
    If D <> "" Then
        Do
            Filename = D
            
            If UCase(Filename) = "NWSCRIPT.NSS" Then
            
                ' Set Defaults
                CommentBlockEnded = True
                FileLineNumber = 0
                FunctionsPerFile = 0
                
                ReDim FunctionLineNumber(1) As Integer
                ReDim Parameters(1) As String
                ReDim FunctionName(1) As String
                ReDim Returns(1) As String
                ReDim Comment(1) As String
                
                'Open each file in turn
                Open NWN_Dir & "\" & D For Input As #1
                    Do
                        Line Input #1, NWNFileData
                        FileLineNumber = FileLineNumber + 1
                        
                        ' Ensure the data returned is not just a blank line
                        If NWNFileData <> "" Then
                            'Remove leading and trailing spaces
                            NWNFileData = Trim(NWNFileData)
                            
                            ' Look for comments
                            For i = 1 To NoOfCommentTypes
                                
                                ' Check if this is a valid comment
                                If Comments(i).Start = "" Then
                                    Exit For ' No comment at this location
                                End If
                                
                                If CommentBlockEnded Then
                                    ' Look for start of comment
                                    If InStr(1, NWNFileData, Comments(i).Start) <> 0 Then
                                        ' Is this one of the common comment block items we don't want?
                                        DoExit = False
                                        If NWNFileData = "::/" Then DoExit = True
                                        If NWNFileData = "//::::" Then DoExit = True
                                        If NWNFileData = "//::" Then DoExit = True
                                        If NWNFileData = "/*" Then DoExit = True
                                        If NWNFileData = "//" Then DoExit = True
                                        If NWNFileData = "/*::" Then DoExit = True
                                        If NWNFileData = "//::///////////////////////////////////////////////" Then DoExit = True
                                        If UCase(Left(NWNFileData, 16)) = "//:: COPYRIGHT (" Then DoExit = True
                                        If UCase(Left(NWNFileData, 15)) = "//:: CREATED BY" Then DoExit = True
                                        If UCase(Left(NWNFileData, 15)) = "//:: CREATED ON" Then DoExit = True
                                        
                                        CommentPart = Replace(NWNFileData, "//", "")
                                        If CommentPart = "::" Then DoExit = True
                                        If CommentPart = "/" Then DoExit = True
                                        If InStr(1, UCase(CommentPart), "PASS BY:") <> 0 Then DoExit = True
                                        If InStr(1, UCase(CommentPart), "LAST UPDATED BY:") <> 0 Then DoExit = True
                                        
                                        ' Remove deliberate remarks
                                        If CommentPart = "************************************************************************************************************************************" Then DoExit = True
                                        
                                        ' Remove commented out code
                                        If Right(CommentPart, 1) = ";" Then DoExit = True
                                        If UCase(Left(NWNFileData, 5)) = "//INT" Then DoExit = True
                                        If UCase(Left(NWNFileData, 6)) = "// INT" Then DoExit = True
                                        
                                        If UCase(Left(NWNFileData, 8)) = "//RETURN" Then DoExit = True
                                        If UCase(Left(NWNFileData, 9)) = "// RETURN" Then DoExit = True
                                        
                                        If UCase(Left(NWNFileData, 3)) = "IF " Then DoExit = True
                                        If UCase(Left(NWNFileData, 5)) = "ELSE " Then DoExit = True
                                        
                                        If CommentPart = "}" Then DoExit = True
                                        If CommentPart = "{" Then DoExit = True
                                        
                                        If Not DoExit Then
                                            ' Remove the actual comment start characters
                                            NWNFileData = Trim(Replace(NWNFileData, Comments(i).Start, ""))
                                            
                                            'Remove instances of "::"
                                            NWNFileData = Trim(Replace(NWNFileData, "::", ""))
                                            
                                            ' Remove instances of "- " at the beginning of comments
                                            'If Left(NWNFileData, 2) = "- " Then NWNFileData = Trim(Right(NWNFileData, Len(NWNFileData) - 2))
                                            ' NOTE removed this, due to Bioware's [over]use of "-"
                                            
                                            ReDim Preserve Comment(FunctionsPerFile + 1)
                                            
                                            Comment(FunctionsPerFile + 1) = Comment(FunctionsPerFile + 1) & NWNFileData & vbCrLf
                                            If Comments(i).LineIsSingle = False Then
                                            
                                            Else
                                                CommentBlockEnded = True
                                            End If
                                        End If
                                    End If
                                Else
                                    ' Look for end of comment
                                    ' NOTE I don't believe this is working.
                                    If InStr(1, NWNFileData, Comments(i).Finish) <> 0 Then
                                        ' Remove the actual comment finish characters
                                        NWNFileData = Trim(Replace(NWNFileData, Comments(i).Finish, ""))
                                        Comment(FunctionsPerFile + 1) = ""
                                        CommentBlockEnded = True
                                    End If
                                End If
                            Next i
                        End If
                        
                        ' Look for a Function Declaraction
                        isFunction = False
                        
                        If UCase(Left(NWNFileData, 5)) = "VOID " Then isFunction = True
                        If UCase(Left(NWNFileData, 4)) = "INT " And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        If UCase(Left(NWNFileData, 7)) = "ACTION " And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        If UCase(Left(NWNFileData, 7)) = "EFFECT " And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        If UCase(Left(NWNFileData, 6)) = "EVENT " And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        If UCase(Left(NWNFileData, 6)) = "FLOAT " And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        If UCase(Left(NWNFileData, 9)) = "LOCATION " And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        If UCase(Left(NWNFileData, 6)) = "OBJECT" And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        If UCase(Left(NWNFileData, 7)) = "STRING " And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        If UCase(Left(NWNFileData, 7)) = "TALENT " And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        If UCase(Left(NWNFileData, 7)) = "VECTOR " And InStr(1, NWNFileData, "=") = 0 Then isFunction = True
                        
                        If isFunction Then
                            
                            FunctionsPerFile = FunctionsPerFile + 1
                            ReDim Preserve FunctionLineNumber(FunctionsPerFile)
                            
                            FunctionLineNumber(FunctionsPerFile) = FileLineNumber
                            
                            ' Retrieve Function name
                            LeftPos = InStr(1, NWNFileData, " ")
                            RightPos = InStr(1, NWNFileData, "(")
                            
                            ' Ensure this is a function, and not (MAYBE) a CONST declaration
                            If RightPos <> 0 And InStr(1, NWNFileData, ")") <> 0 Then
                                ReDim Preserve FunctionName(FunctionsPerFile)
                                ReDim Preserve Parameters(FunctionsPerFile)
                                ReDim Preserve Returns(FunctionsPerFile)
                                ReDim Preserve Comment(FunctionsPerFile)
                                
                                ' Determine Return type
                                Returns(FunctionsPerFile) = Trim(Left(NWNFileData, InStr(1, NWNFileData, " ")))
                                
                                FunctionName(FunctionsPerFile) = Trim(Mid(NWNFileData, LeftPos, RightPos - LeftPos))
                                ' Retrieve parameter list
                                LeftPos = InStr(1, NWNFileData, "(")
                                RightPos = InStr(1, NWNFileData, ")")
                                Parameters(FunctionsPerFile) = Trim(Mid(NWNFileData, LeftPos + 1, RightPos - LeftPos - 1))
                                
                                ' Replace instances of ", " with "," (ie remove the space character)
                                Parameters(FunctionsPerFile) = Replace(Parameters(FunctionsPerFile), ", ", ",")
                                
                                ' Split parameters up
                                Params() = Split(Parameters(FunctionsPerFile), ",")
                                
                                ' Ensure no blank lines exist
                                Comment(FunctionsPerFile) = Replace(Comment(FunctionsPerFile), vbCrLf & vbCrLf, "")
                                
                                ' Ensure nothing to break XML/HTML file
                                Comment(FunctionsPerFile) = Replace(Comment(FunctionsPerFile), "<", "less than")
                                Comment(FunctionsPerFile) = Replace(Comment(FunctionsPerFile), ">", "greater than")
                                Comment(FunctionsPerFile) = Replace(Comment(FunctionsPerFile), "&", "and")
                                
                                Print #2, "<function include=""" & Filename & """>"
                                Print #2, "<name>" & FunctionName(FunctionsPerFile) & "</name>"
                                Print #2, "<call>"
                                Print #2, "<return_type>" & Returns(FunctionsPerFile) & "</return_type>"
                                
                                j = 0
                                AllParams = ""
                                Erase ParamType()
                                For Each v In Params()
                                    j = j + 1
                                    ParamType() = Split(CStr(v), " ")
                                    Print #2, "<param index=""" & j & """ type=""" & ParamType(0) & """>" & ParamType(1) & "</param>"
                                    AllParams = AllParams & ParamType(0) & " " & ParamType(1) & ", "
                                Next
                                
                                If AllParams <> "" Then
                                    ' Remove final ", " from the list of paramaters
                                    AllParams = Left(AllParams, Len(AllParams) - 2)
                                End If
                                
                                Print #2, "</call>"
                                Print #2, "<remarks>" & Filename & ": " & FunctionLineNumber(FunctionsPerFile) & "</remarks>"
                                If Comment(FunctionsPerFile) <> "" Then
                                    ' Remove final vbcrlf from comments
                                    If Right(Comment(FunctionsPerFile), 2) = vbCrLf Then
                                        Comment(FunctionsPerFile) = Left(Comment(FunctionsPerFile), Len(Comment(FunctionsPerFile)) - 2)
                                    End If
                                    ' Ensure XML is closed properly
                                    Comment(FunctionsPerFile) = Replace(Comment(FunctionsPerFile), vbCrLf, "</note>" & vbCrLf & "<note>")
                                    Print #2, "<note>" & Comment(FunctionsPerFile) & "</note>"
                                End If
                                Print #2, "<prototype>" & Returns(FunctionsPerFile) & " " & FunctionName(FunctionsPerFile) & "(" & AllParams & ");</prototype>"
                                Print #2, "</function>"
                            End If
                        End If
                    Loop Until EOF(1)
                Close 1
            End If
            Debug.Print "Finished " & D
            D = Dir
        Loop Until D = ""
    End If
    Print #2, "</nwn_functions>"
    Close 2
End Sub

Public Sub AddComment(LineIsSingle As Boolean, Start As String, Optional Finish As String = "")
    Static CommentNumber As Integer
    
    CommentNumber = CommentNumber + 1
    Comments(CommentNumber).LineIsSingle = LineIsSingle
    Comments(CommentNumber).Start = Start
    Comments(CommentNumber).Finish = Finish
End Sub


