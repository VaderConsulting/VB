Attribute VB_Name = "modMain"
    Public gStepNo As Integer
    Public gStepMax As Integer
    Public gL(19) As Integer
    Public gT(19) As Integer
    Public gW(19) As Integer
    Public gH(19) As Integer
    Public gMax(19) As Integer
    Public gValue(19) As Integer
    Public gDay(7) As String
    
    Public dbDiamond As Database
    Public rsStock As Recordset
    Public rsWarranty As Recordset
    Public rsSuppliers As Recordset
    Public fldSerial As Field
    Public gTotal As Integer
    Public db As String
    'Public Const Db = "c:\data\access\diamond.mdb"

Sub Main()
    dblocation = GetSetting("Diamond", "Setup", "dblocation", "")
    If dblocation = "" Then
        Do Until db <> ""
            db = InputBox("Enter database location", "Setup info required", "")
        Loop
        SaveSetting "Diamond", "Setup", "dbLocation", db
    Else
        db = dblocation
    End If
    Set dbDiamond = OpenDatabase(db)
    Set rsStock = dbDiamond.OpenRecordset("Stock")
    Set rsWarranty = dbDiamond.OpenRecordset("Warranty")
    Set rsSuppliers = dbDiamond.OpenRecordset("Suppliers")
    gStepNo = 0
    gStepMax = 19
    For a = 0 To gStepMax
        gL(a) = 360
        gL(a) = 1080
        gW(a) = 1440
    Next a
    gT(0) = 0:     gH(0) = 735:                  gMax(0) = 7: gValue(0) = 0
    gT(1) = 720:   gH(1) = 735:                  gMax(1) = 7: gValue(1) = 1
    gT(2) = 1440:  gH(2) = 1215:                 gMax(2) = 7: gValue(2) = 2
    gT(3) = 3360:  gH(3) = 1215:                 gMax(3) = 7: gValue(3) = 3
    gT(4) = 4560:  gH(4) = 1215:                 gMax(4) = 7: gValue(4) = 4
    gT(5) = 5760:  gH(5) = 1215:                 gMax(5) = 7: gValue(5) = 5
    gT(6) = 7200:  gH(6) = 1215:                 gMax(6) = 7: gValue(6) = 6
    gT(7) = 1440:  gH(7) = 1215:  gL(7) = 2280:  gMax(7) = 5: gValue(7) = 3
    gT(8) = 2640:  gH(8) = 735:   gL(8) = 2280:  gMax(8) = 4: gValue(8) = 4
    gT(9) = 7440:  gH(9) = 735:   gL(9) = 2280:  gMax(9) = 6: gValue(9) = 6
    gT(10) = 2640: gH(10) = 735:  gL(10) = 3360: gMax(10) = 5: gValue(10) = 4
    gT(11) = 3600: gH(11) = 735:  gL(11) = 3360: gMax(11) = 7: gValue(11) = 4
    gT(12) = 4560: gH(12) = 1215: gL(12) = 3360: gMax(12) = 7: gValue(12) = 5
    gT(13) = 5760: gH(13) = 1215: gL(13) = 3360: gMax(13) = 7: gValue(13) = 6
    gT(14) = 6960: gH(14) = 735:  gL(14) = 3360: gMax(14) = 7: gValue(14) = 7
    gT(15) = 2640: gH(15) = 735:  gL(15) = 4560: gMax(15) = 5: gValue(15) = 5
    gT(16) = 4800: gH(16) = 735:  gL(16) = 4560: gMax(16) = 5: gValue(16) = 5
    gT(17) = 6000: gH(17) = 735:  gL(17) = 4560: gMax(17) = 7: gValue(17) = 7
    gT(18) = 8400: gH(18) = 735:  gL(18) = 1080: gMax(18) = 7: gValue(18) = 7
    gT(19) = 7440: gH(19) = 735:  gL(19) = -120: gMax(19) = 7: gValue(19) = 7
    
    gDay(1) = "Sunday"
    gDay(2) = "Monday"
    gDay(3) = "Tuesday"
    gDay(4) = "Wednesday"
    gDay(5) = "Thursday"
    gDay(6) = "Friday"
    gDay(7) = "Saturday"
    
    mdiParent.Show
End Sub

Sub StepChange()
    If gStepNo < 0 Then gStepNo = 0
    If gStepNo > gStepMax Then gStepNo = gStepMax
    frmFlow.lblStep = gStepNo
    
    ' Positon Orange indicator shape
    frmFlow.shpHilite.Left = gL(gStepNo)
    frmFlow.shpHilite.Top = gT(gStepNo)
    frmFlow.shpHilite.Width = gW(gStepNo)
    frmFlow.shpHilite.Height = gH(gStepNo)
    frmFlow.shpHilite.Refresh
    
    'Update progress bar
    frmFlow.pbrStatus.Max = gMax(gStepNo)
    frmFlow.pbrStatus.Value = gValue(gStepNo)
    frmFlow.sldStatus.Max = gMax(gStepNo)
    frmFlow.sldStatus.Value = gValue(gStepNo)
End Sub

Sub WarrantyCheck()
    ' check if in warranty
    Dim days As Integer
    days = DateDiff("d", frmData.lblField(5), Now)
    Select Case days
        Case Is < 365
            If days < 8 Then
                frmData.lblStatus = "SALE WAS WITHIN THE LAST 7 DAYS"
                gStepNo = 5
                StepChange
            End If
            If days > 7 Then
                frmData.lblStatus = "ITEM IS IN WARRANTY"
                gStepNo = 9
                StepChange
            End If
        Case Is > 365
            frmData.lblStatus = "WARRANTY HAS LAPSED"
            gStepNo = 11
            StepChange
    End Select
    'lblID = rsStock.Fields(1)
    frmData.lstMatch_Click
    frmData.lblDays = days
End Sub

Sub CheckStock()
    'On Error Resume Next
    rsStock.MoveFirst
    a = 0
    Do Until ID = frmData.lblField(1) And out = ""
        a = a + 1
        ID = rsStock.Fields(1)
        out = rsStock.Fields(5) & ""
        rsStock.MoveNext
    Loop
    If a <> rsStock.RecordCount Then
        frmData.lblStatus = "Found replacement item in stock."
        gStepNo = 18
        StepChange
    Else
        frmData.lblStatus = "Replacement item not in stock."
        gStepNo = 9
        StepChange
    End If
End Sub
