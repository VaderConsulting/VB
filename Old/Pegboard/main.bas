Attribute VB_Name = "Module1"
    Public gMail(232) As String
    Public gName(232) As String
    Public gPersonalNumber(232) As Long
    Public gLocationText(300) As String
    Public gLocation(300) As Integer  ' where person number x is currently
    Public gBilletNumber(232) As String * 3
    Public gBilletText(232) As String
    Public gRank(232) As String
    Public gIndex As Integer
    Public gTime(232) As String
    Public gOtherInfo(232) As String
    
    
Sub Main()
    Open "C:\locations.txt" For Input As #1
    Line Input #1, no_of_locations
    For a = 1 To Val(no_of_locations)
        Input #1, location_no, locationtext
        gLocation(a) = location_no
        gLocationText(location_no) = locationtext
    Next a
    Close 1
    
    Open "C:\personnel.txt" For Input As #1
    Line Input #1, no_of_personnel
    For a = 1 To Val(no_of_personnel)
        Input #1, person_no, billetnumber, billettext, rank, surname, personalnumber
        gBilletNumber(person_no) = billetnumber
        gBilletText(person_no) = billettext
        gRank(person_no) = rank
        gName(person_no) = surname
        gPersonalNumber(person_no) = personalnumber
    Next a
    Close 1
    
    For a = 1 To 232
        gLocation(a) = 0
    Next a
    
    frmLeave.Show
End Sub
