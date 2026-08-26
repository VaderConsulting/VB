Attribute VB_Name = "UserNames"
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long

Public Property Get UserName() As String
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    UserName = Left$(sBuffer, lSize)
    UserName = ClipNull(UserName)
End Property

Public Function ClipNull(InString As String) As String
    Dim intpos As Integer


    If Len(InString) Then
        intpos = InStr(InString, vbNullChar)


        If intpos > 0 Then
            ClipNull = Left(InString, intpos - 1)
        Else
            ClipNull = InString
        End If
    End If
End Function

