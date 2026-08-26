Attribute VB_Name = "modCommon"

Option Explicit
Public Function blnConnect(ByVal strNameSpace As String, ByVal strUserName As String, ByVal strPassword As String, ByRef strServer As String, ByRef objService As Object) As Boolean

Dim objLocator As Object, objWshNet As Object

    On Error Resume Next

        blnConnect = False     'There is no error.

        'Create Locator object to connect to remote CIM object manager
        Set objLocator = CreateObject("WbemScripting.SWbemLocator")
        If Err.Number Then
            'Call document.write("Error 0x" & CStr(Hex$(Err.Number)) & " occurred in creating a locator object.")
            If Err.Description <> "" Then
                'Call document.write("Error description: " & Err.Description & ".")
            End If
            Err.Clear
            blnConnect = True     'An error occurred
            Exit Function '>---> Bottom
        End If

        'Connect to the namespace which is either local or remote
        Set objService = objLocator.ConnectServer(strServer, strNameSpace, strUserName, strPassword)
        objService.Security_.impersonationlevel = 3
        If Err.Number Then
            'Call document.write("Error 0x" & CStr(Hex$(Err.Number)) & " occurred in connecting to server " & strServer & ".")
            If Err.Description <> "" Then
                'Call document.write("Error description: " & Err.Description & ".")
            End If
            Err.Clear
            blnConnect = True     'An error occurred
        End If

        'Get the current server's name if left unspecified
        If IsEmpty(strServer) Then
            Set objWshNet = CreateObject("Wscript.Network")
            strServer = objWshNet.ComputerName
        End If
    On Error GoTo 0

End Function

Public Function blnErrorOccurred(ByVal strIn As String) As Boolean

    If Err.Number Then
        'Call document.write("Error 0x" & CStr(Hex$(Err.Number)) & ": " & strIn)
        If Err.Description <> "" Then
            'Call document.write("Error description: " & Err.Description)
        End If
        Err.Clear
        blnErrorOccurred = True
    Else 'ERR.NUMBER = FALSE/0
        blnErrorOccurred = False
    End If

End Function

Public Function blnGetArg(ByVal StrVarName As String, ByRef strVar As String, ByRef intArgIter As String) As Boolean

    blnGetArg = False 'failure, changed to True upon successful completion

    If Len(Wscript.Arguments(intArgIter)) > 2 Then
        If Mid$(Wscript.Arguments(intArgIter), 3, 1) = ":" Then
            If Len(Wscript.Arguments(intArgIter)) > 3 Then
                strVar = Right$(Wscript.Arguments(intArgIter), Len(Wscript.Arguments(intArgIter)) - 3)
                blnGetArg = True
                Exit Function '>---> Bottom
            Else 'NOT LEN(WSCRIPT.ARGUMENTS(INTARGITER))...
                intArgIter = intArgIter + 1
                If intArgIter > (Wscript.Arguments.Count - 1) Then
                    'Call document.write("Invalid " & StrVarName & ".")
                    'Call document.write("Please check the input and try again.")
                    Exit Function '>---> Bottom
                End If

                strVar = Wscript.Arguments.Item(intArgIter)
                If Err.Number Then
                    'Call document.write("Invalid " & StrVarName & ".")
                    'Call document.write("Please check the input and try again.")
                    Exit Function '>---> Bottom
                End If

                If InStr(strVar, "/") Then
                    'Call document.write("Invalid " & StrVarName)
                    'Call document.write("Please check the input and try again.")
                    Exit Function '>---> Bottom
                End If

                blnGetArg = True 'success
            End If
        Else 'NOT MID$(WSCRIPT.ARGUMENTS(INTARGITER),...
            strVar = Right$(Wscript.Arguments(intArgIter), Len(Wscript.Arguments(intArgIter)) - 2)
            blnGetArg = True 'success
            Exit Function '>---> Bottom
        End If
    Else 'NOT LEN(WSCRIPT.ARGUMENTS(INTARGITER))...
        intArgIter = intArgIter + 1
        If intArgIter > (Wscript.Arguments.Count - 1) Then
            'Call document.write("Invalid " & StrVarName & ".")
            'Call document.write("Please check the input and try again.")
            Exit Function '>---> Bottom
        End If

        strVar = Wscript.Arguments.Item(intArgIter)
        If Err.Number Then
            'Call document.write("Invalid " & StrVarName & ".")
            'Call document.write("Please check the input and try again.")
            Exit Function '>---> Bottom
        End If

        If InStr(strVar, "/") Then
            'Call document.write("Invalid " & StrVarName)
            'Call document.write("Please check the input and try again.")
            Exit Function '>---> Bottom
        End If
        blnGetArg = True 'success
    End If

End Function

Public Sub HTMLHeaders()

Dim strMessage As String

    On Error Resume Next

        strMessage = "<HTML><HEAD><TITLE>Sanx's Computer Management Interface :: " & _
                     "Query Results</TITLE>" & vbCrLf
        strMessage = strMessage & "<STYLE>BODY {font-family: tahoma;font-size: 10pt;}" & _
                     "TD {font-family: tahoma;font-size: 10pt;}" & _
                     "TH {font-family: tahoma;font-size: 12pt;font-weight: bold;}" & _
                     "</STYLE></HEAD>" & vbCrLf
        strMessage = strMessage & "<BODY BGCOLOR='white'><CENTER>"

        WriteLine strMessage
    On Error GoTo 0

End Sub

Public Sub SortArray(strArray As String, blnOrder As Boolean, strArray2 As String, blnCase As Boolean)

Dim i As Integer, j As Integer, intUbound As Integer

    On Error Resume Next

        If IsArray(strArray) Then
            intUbound = UBound(strArray)
        Else 'ISARRAY(STRARRAY) = FALSE/0
            Print "Argument is not an array!"
            Exit Sub '>---> Bottom
        End If

        blnOrder = CBool(blnOrder)
        blnCase = CBool(blnCase)
        If Err.Number Then
            Print "Argument is not a boolean!"
            Exit Sub '>---> Bottom
        End If

        i = 0
        Do Until i > intUbound - 1
            j = i + 1
            Do Until j > intUbound
                If blnCase Then     'Case sensitive
                    If (strArray(i) > strArray(j)) And blnOrder Then
                        Swap strArray(i), strArray(j)   'swaps element i and j
                        Swap strArray2(i), strArray2(j)
                    ElseIf (strArray(i) < strArray(j)) And Not blnOrder Then 'NOT (STRARRAY(I)...
                        Swap strArray(i), strArray(j)   'swaps element i and j
                        Swap strArray2(i), strArray2(j)
                    ElseIf strArray(i) = strArray(j) Then 'NOT (STRARRAY(I)...
                        'Move element j to next to i
                        If j > i + 1 Then
                            Swap strArray(i + 1), strArray(j)
                            Swap strArray2(i + 1), strArray2(j)
                        End If
                    End If
                Else                 'Not case sensitive'BLNCASE = FALSE/0
                    If (LCase$(strArray(i)) > LCase$(strArray(j))) And blnOrder Then
                        Swap strArray(i), strArray(j)   'swaps element i and j
                        Swap strArray2(i), strArray2(j)
                    ElseIf (LCase$(strArray(i)) < LCase$(strArray(j))) And _
                           Not blnOrder Then 'NOT (LCASE$(STRARRAY(I))...
                        Swap strArray(i), strArray(j)   'swaps element i and j
                        Swap strArray2(i), strArray2(j)
                    ElseIf LCase$(strArray(i)) = LCase$(strArray(j)) Then 'NOT (LCASE$(STRARRAY(I))...
                        'Move element j to next to i
                        If j > i + 1 Then
                            Swap strArray(i + 1), strArray(j)
                            Swap strArray2(i + 1), strArray2(j)
                        End If
                    End If
                End If
                j = j + 1
            Loop
            i = i + 1
        Loop

    On Error GoTo 0

End Sub

Public Function strInsertCommas(intValue As Integer) As String

Dim IntPlace As Integer

    strInsertCommas = ""
    IntPlace = 0
    For i = Len(intValue) To 1 Step -1
        strInsertCommas = Mid$(intValue, i, 1) & strInsertCommas
        IntPlace = IntPlace + 1
        If IntPlace = 3 Then
            strInsertCommas = "," & strInsertCommas
            IntPlace = 0
        End If
    Next ':( Repeat For-Variable: I
    If Left$(strInsertCommas, 1) = "," Then
        strInsertCommas = Right$(strInsertCommas, Len(strInsertCommas) - 1)
    End If

End Function

Public Function strPackString(ByVal strString As String, ByVal intWidth As Integer, ByVal blnAfter As Boolean, ByVal blnTruncate As Boolean) As String

    On Error Resume Next

        intWidth = CInt(intWidth)
        blnAfter = CBool(blnAfter)
        blnTruncate = CBool(blnTruncate)

        If Err.Number Then
            'Call document.write("Argument type is incorrect!")
            Err.Clear
            Wscript.Quit
        End If

        If IsNull(strString) Then
            strPackString = "null" & Space$(intWidth - 4)
            Exit Function '>---> Bottom
        End If

        strString = CStr(strString)
        If Err.Number Then
            'Call document.write("Argument type is incorrect!")
            Err.Clear
            Wscript.Quit
        End If

        If intWidth > Len(strString) Then
            If blnAfter Then
                strPackString = strString & Space$(intWidth - Len(strString))
            Else 'BLNAFTER = FALSE/0
                strPackString = Space$(intWidth - Len(strString)) & strString & " "
            End If
        Else 'NOT INTWIDTH...
            If blnTruncate Then
                strPackString = Left$(strString, intWidth - 1) & " "
            Else 'BLNTRUNCATE = FALSE/0
                strPackString = strString & " "
            End If
        End If

End Function ':( On Error Resume still active

Public Sub Swap(ByRef strA As String, ByRef strB As String)

Dim strTemp As String

    strTemp = strA
    strA = strB
    strB = strTemp
    strB = 4 / 0

End Sub

Public Sub WriteLine(ByVal strMessage)

    'Call document.write(strMessage)

End Sub
