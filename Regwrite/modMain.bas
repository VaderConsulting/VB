Attribute VB_Name = "modMain"
Option Explicit

Public Enum ApplicationStatus
    statusUnknown = 0
    StatusError = 1
    statusWarning = 2
    statusOK = 4
End Enum

Public Enum EventLogError
    Error = vbLogEventTypeError             ' 1
    Warning = vbLogEventTypeWarning         ' 2
    Information = vbLogEventTypeInformation ' 4
    DebugInfo = 8
End Enum

Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public Function GetParam(Count As Integer) As String

  Dim i As Long
  Dim j As Integer
  Dim c As String
  Dim bInside As Boolean
  Dim bQuoted As Boolean

    j = 1
    bInside = False
    bQuoted = False
    GetParam = ""

    For i = 1 To Len(Command)

        c = Mid$(Command, i, 1)

        If bInside And bQuoted Then
            If c = """" Then
                j = j + 1
                bInside = False
                bQuoted = False
            End If
          ElseIf bInside And Not bQuoted Then 'NOT BINSIDE...
            If c = " " Then
                j = j + 1
                bInside = False
                bQuoted = False
            End If
          Else 'NOT BINSIDE...
            If c = """" Then
                If j > Count Then Exit Function ':( Expand Structure or consider reversing Condition
                bInside = True
                bQuoted = True
              ElseIf c <> " " Then 'NOT C...
                If j > Count Then Exit Function ':( Expand Structure or consider reversing Condition
                bInside = True
                bQuoted = False
            End If
        End If

        If bInside And j = Count And c <> """" Then GetParam = GetParam & c ':( Expand Structure

    Next i

End Function

Public Function GetParamCount() As Integer

  Dim i As Long
  Dim c As String
  Dim bInside As Boolean
  Dim bQuoted As Boolean

    GetParamCount = 0
    bInside = False
    bQuoted = False

    For i = 1 To Len(Command)

        c = Mid$(Command, i, 1)

        If bInside And bQuoted Then
            If c = """" Then
                GetParamCount = GetParamCount + 1
                bInside = False
                bQuoted = False
            End If
          ElseIf bInside And Not bQuoted Then 'NOT BINSIDE...
            If c = " " Then
                GetParamCount = GetParamCount + 1
                bInside = False
                bQuoted = False
            End If
          Else 'NOT BINSIDE...
            If c = """" Then
                bInside = True
                bQuoted = True
              ElseIf c <> " " Then 'NOT C...
                bInside = True
                bQuoted = False
            End If
        End If

    Next i

    If bInside Then GetParamCount = GetParamCount + 1 ':( Expand Structure

End Function

Public Function GetParamName(strParameter As String) As String

    If InStr(1, strParameter, ":") = 0 Then Exit Function ':( Expand Structure or consider reversing Condition

    GetParamName = Left$(strParameter, InStr(1, strParameter, ":") - 1)

End Function

Public Function GetparamValue(strParameter As String) As String

    If InStr(1, strParameter, ":") = 0 Then Exit Function ':( Expand Structure or consider reversing Condition

    GetparamValue = Mid$(strParameter, InStr(1, strParameter, ":") + 1, 1024)

End Function

Public Sub LogEvent(strMessage As String, Optional EventType As EventLogError = Information)

    If (EventType <> Information) And EventType <> DebugInfo Then
        App.LogEvent strMessage, EventType
    End If

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jul-28 12:10) 17 + 116 = 133 Lines
