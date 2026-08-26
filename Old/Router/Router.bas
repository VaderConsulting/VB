Attribute VB_Name = "Module1"
Public gDir(160) As String
Public gExtension As String
Public gGroup(10) As String
Public gGroups As String
Public gArchive As String
Public gLog
Public gClass(6) As Integer
Public gClassification As String
Public gSecurity As Integer
Public gType(6)
Public gNo_Of_Classes
Rem 160 directories for files

Sub Main()
    gNo_Of_Classes = 6
    distribution_file = "C:\DISTRIBUTION.TXT"
    gGroups = "C:\GROUPS.TXT"
    gLog = "C:\Router.txt"
    gType(1) = "Unclassified"
    gType(2) = "Restricted"
    gType(3) = "Confidential"
    gType(4) = "Secret"
    gType(5) = "Top Secret"
    gType(6) = "In Confidence"
    If Dir(distribution_file) <> "" Then
        Open distribution_file For Input As #1
        Rem get the number of individuals
        Line Input #1, Number
        Number = Val(Number)
        Rem then get the individuals names
        For a = 0 To Number - 1
            Line Input #1, dist_text
            frmSend.chkPersonnel(a).Caption = dist_text
            frmSend.chkPersonnel(a).Visible = True
            frmSend.lstIndividuals.AddItem dist_text
            Line Input #1, dist_directory
            gDir(a) = dist_directory
            frmSend.chkPersonnel(a).ToolTipText = gDir(a)
        Next a
        Rem get the Archive directory
        Line Input #1, gArchive
        Rem get the default extension
        Line Input #1, gExtension
        Close 1
    Else
        Rem some problem.  Distribution file was not found
        MsgBox distribution_file & " not found"
    End If
    Close 1
    If Dir(gGroups) <> "" Then
        Open gGroups For Input As #1
        Rem get the group names
        For a = 0 To 9
            Line Input #1, group_no
            Line Input #1, group_name
            gGroup(a) = group_name
        Next a
        Rem get the group members
        For a = 1 To 10
            Input #1, group_no
            Input #1, number_of_members
            If number_of_members > 0 Then
                For b = 1 To number_of_members
                    Input #1, member
                    frmSend.lstGroup(group_no - 1).AddItem frmSend.lstIndividuals.List(member)
                    frmSend.lstGroup(group_no - 1).ItemData(b - 1) = member
                Next b
            End If
        Next a
        Close 1
    Else
        Rem some problem.  Groups file was not found
        MsgBox gGroups & " not found"
    End If
    Rem show main window
    frmSend.Show
    Rem show splash screen, and keep the focus on the splash screen
    frmSplash.Show vbModal
End Sub
