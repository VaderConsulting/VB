http://www.vb-helper.com/HowTo/findandreplace.zip

	Purpose
Search for files of a specific type in a directory that contain a string and replace the string

	Method
Use Dir$ to list the files matching the selected patterns. Read the files into a string, use Replace to make the substitution, and rewrite the file.

    Private Sub cmdReplace_Click()
    Dim from_text As String
    Dim to_text As String
    Dim dir_name As String
    Dim patterns As Variant
    Dim file_name As String
    Dim i As Integer
    Dim results As String

        ' Get the text to find and replace.
        from_text = txtFind.Text
        to_text = txtReplace.Text

        ' Get the directory name.
        dir_name = FileList.Path
        If Right$(dir_name, 1) <> "\" Then dir_name = dir_name & "\"

        ' Get the file patterns.
        patterns = Split(FileList.Pattern, ";")

        results = "Files:"

        ' Repeat for each pattern.
        m_QuitEarly = False
        For i = LBound(patterns) To UBound(patterns)
            ' Add the pattern to the file name.
            file_name = Dir$(dir_name & patterns(i))
            Do While Len(file_name) > 0
                ' Process this file.
                If ReplaceInFile(dir_name & file_name, from_text, to_text) Then
                    results = results & " " & file_name
                End If
                If m_QuitEarly Then Exit For
    
                ' Get the next file.
                file_name = Dir$()
            Loop
        Next i

        MsgBox results
    End Sub

    ' In the file file_name, replace occurrances of from_text
    ' with to_text. Return True if the string appeared and
    ' was replaced, False if the string was not in this file.
    Private Function ReplaceInFile(ByVal file_name As String, ByVal from_text As String, ByVal to_text As String) As Boolean
    Dim fnum As Integer
    Dim file_text As String

        On Error GoTo ReplaceError

        ' Read the file.
        fnum = FreeFile
        Open file_name For Input As fnum
        file_text = Input$(LOF(fnum), #fnum)
        Close #fnum

        ' See if the text appears.
        If InStr(file_text, from_text) > 0 Then
            ' Replace the text.
            file_text = Replace(file_text, from_text, to_text)

            ' Rewrite the file.
            fnum = FreeFile
            Open file_name For Output As fnum
            Print #fnum, file_text;
            Close #fnum

            ReplaceInFile = True
        End If
        Exit Function

    ReplaceError:
        Select Case MsgBox("Error " & Err.Number & _
          " processing file " & file_name & _
          vbCrLf & Err.Description & _
          "Try again?", vbYesNoCancel)
            Case vbYes
                Resume
            Case vbNo
                Exit Function
            Case Else
                m_QuitEarly = True
                Exit Function
        End Select
    End Function

	Disclaimer
This example program is provided "as is" with no warranty of any kind. It is
intended for demonstration purposes only. In particular, it does no error
handling. You can use the example in any form, but please mention
www.vb-helper.com.
