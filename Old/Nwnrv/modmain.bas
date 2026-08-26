Attribute VB_Name = "modMain"
Option Explicit

' *********************************************************
' Turn a series of integer bytes into a decimal number
Public Function MakeLongFromStringBytes(strInput As String) As Long
' *********************************************************
' Expects to receive a string such as nB6f
    Dim Portions(8) As String
    Dim Numbers(8) As Long
    Dim Chars(8) As String
    Dim i As Long, intLength As Integer
    Dim HexString As String
    Dim Pos As Integer
    
    If strInput = "" Then Exit Function
    
    intLength = Len(strInput)
    
    For i = 1 To intLength
        ' Now don't forget that the byte order is reversed, so, re-reverse it.
        Pos = (intLength + 1) - i
        Portions(i) = Mid(strInput, Pos, 1)
        Numbers(i) = Asc(Portions(i))
        Chars(i) = Hex(Numbers(i))
        If Len(Chars(i)) = 1 Then Chars(i) = "0" & Chars(i)
    Next i
    
    ' Prepend Hex character
    HexString = "&H"
    
    For i = 1 To intLength
        HexString = HexString & Chars(i)
    Next i
    
    ' Turn this hex number back to decimal
    MakeLongFromStringBytes = CLng(HexString)
End Function
