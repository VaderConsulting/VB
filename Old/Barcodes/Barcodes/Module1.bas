Attribute VB_Name = "Module1"
Function CodeAToByte(Number As String) As String
Select Case Number
Case 0
    CodeAToByte = "0001101"
Case 1
    CodeAToByte = "0011001"
Case 2
    CodeAToByte = "0010011"
Case 3
    CodeAToByte = "0111101"
Case 4
    CodeAToByte = "0100011"
Case 5
    CodeAToByte = "0110001"
Case 6
    CodeAToByte = "0101111"
Case 7
    CodeAToByte = "0111011"
Case 8
    CodeAToByte = "0110111"
Case 9
    CodeAToByte = "0001011"
End Select
End Function
Function CodeBToByte(Number As String) As String
Select Case Number
Case 0
    CodeBToByte = "0100111"
Case 1
    CodeBToByte = "0110011"
Case 2
    CodeBToByte = "0011011"
Case 3
    CodeBToByte = "0100001"
Case 4
    CodeBToByte = "0011101"
Case 5
    CodeBToByte = "0111001"
Case 6
    CodeBToByte = "0000101"
Case 7
    CodeBToByte = "0010001"
Case 8
    CodeBToByte = "0001001"
Case 9
    CodeBToByte = "0010111"
End Select
End Function
Function CodeCToByte(Number As String) As String
Select Case Number
Case 0
    CodeCToByte = "1110010"
Case 1
    CodeCToByte = "1100110"
Case 2
    CodeCToByte = "1101100"
Case 3
    CodeCToByte = "1000010"
Case 4
    CodeCToByte = "1011100"
Case 5
    CodeCToByte = "1001110"
Case 6
    CodeCToByte = "1010000"
Case 7
    CodeCToByte = "1000100"
Case 8
    CodeCToByte = "1001000"
Case 9
    CodeCToByte = "1110100"
End Select
End Function
Function code(ByVal Number As String) As String
Select Case Number
Case 0
    code = "AAAAAA"
Case 1
    code = "AABBAB"
Case 2
    code = "AABBAB"
Case 3
    code = "AABBBA"
Case 4
    code = "ABAABB"
Case 5
    code = "ABBAAB"
Case 6
    code = "ABBBBA"
Case 7
    code = "ABABAB"
Case 8
    code = "ABABBA"
Case 9
    code = "ABBABA"
End Select
End Function
Function PaintCode(frm As Form, fi, se, th)
Dim reihe
Dim z
Dim b
Dim d
frm.Line (0 + 10, 0)-(0 + 10, 25)
frm.Line (2 + 10, 0)-(2 + 10, 25)
reihe = code(fi)
For z = 1 To 6
    If Mid(reihe, z, 1) = "A" Then
        b = CodeAToByte(Mid(se, z, 1))
        For d = 1 To 7
            If Mid(b, d, 1) = 1 Then
                frm.Line ((z - 1) * 7 + d + 3 + 10, 0)-((z - 1) * 7 + d + 3 + 10, 20), &H0
            Else
                frm.Line ((z - 1) * 7 + d + 3 + 10, 0)-((z - 1) * 7 + d + 3 + 10, 20), &HFFFFFF
            End If
        Next
    ElseIf Mid(reihe, z, 1) = "B" Then
        b = CodeBToByte(Mid(se, z, 1))
        For d = 1 To 7
            If Mid(b, d, 1) = 1 Then
                frm.Line ((z - 1) * 7 + d + 3 + 10, 0)-((z - 1) * 7 + d + 3 + 10, 20), &H0
            Else
                frm.Line ((z - 1) * 7 + d + 3 + 10, 0)-((z - 1) * 7 + d + 3 + 10, 20), &HFFFFFF
            End If
        Next
    End If
Next
frm.Line (6 * 7 + 5 + 10, 0)-(6 * 7 + 5 + 10, 25)
frm.Line (6 * 7 + 7 + 10, 0)-(6 * 7 + 7 + 10, 25)
    For z = 1 To 6
        b = CodeCToByte(Mid(th, z, 1))
        For d = 1 To 7
            If Mid(b, d, 1) = 1 Then
                frm.Line ((z - 1) * 7 + d + 50 + 10, 0)-((z - 1) * 7 + d + 50 + 10, 20), &H0
            Else
                frm.Line ((z - 1) * 7 + d + 50 + 10, 0)-((z - 1) * 7 + d + 50 + 10, 20), &HFFFFFF
            End If
        Next
    Next
frm.Line (94 + 10, 0)-(94 + 10, 25)
frm.Line (96 + 10, 0)-(96 + 10, 25)
End Function
Function CheckCode(FullCode As String) As Boolean
Dim a
Dim b
Dim c
b = 1
For a = 1 To 12
    If b = 1 Then
        c = c + Mid(FullCode, a, 1)
        b = 0
    Else
        c = c + (Mid(FullCode, a, 1) * 3)
        b = 1
    End If
Next
If (c + Mid(FullCode, 13, 1)) Mod 10 = 0 Then
    CheckCode = True
Else
    CheckCode = False
End If
End Function
