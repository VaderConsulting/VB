Attribute VB_Name = "modmain"
Public gDefaultAction As String
Public gOrderPath As String
Public gDbPath As String
Public gWordViewPath As String
Public gAutoMakeID As Boolean
Public gLogOnboard As Boolean
Public gLog As Boolean
Public gClickOrigin As String
Public gLogFile As String
Public gDay(7) As String
Public gColours As Variant
Public gJSTime As Variant
Public gSSTime As Variant
Public gOTime As Variant
Public gJSLeave As Variant
Public gSSLeave As Variant
Public gOLeave As Variant
Public gSearch As String
Public Type User
    ID As Integer
    Rank As String * 8
    Christian As String
    Surname As String
    Number As String
    Billet As String
    Department As String
    BilletDesc As String
    Location As String
    Tout As String
    DOut As String
    TIn As String
    DIn As String
    Message As String
    Photo As String
    Comment As String
End Type
Global User(200) As User

    
Sub Main()
    Randomize Timer
    gDefaultAction = "Properties"
    gAutoMakeID = True
    gLogOnboard = True
    gDay(1) = "Sunday"
    gDay(2) = "Monday"
    gDay(3) = "Tuesday"
    gDay(4) = "Wednesday"
    gDay(5) = "Thursday"
    gDay(6) = "Friday"
    gDay(7) = "Saturday"
    gColours = "0800"
    gJSTime = "0750"
    gSSTime = "0755"
    gOTime = "0755"
    gJSLeave = "1550"
    gSSLeave = "1550"
    gOLeave = "1550"
    'gOrderPath = "E:\Data\Word"
    gOrderPath = "J:\Bulletin\Weekly Orders"
    'gDbPath = "E:\DATA\Projects\PersMan"
    'gDbPath = "D:\WINNT\Profiles\administrator.000\Desktop\Persman"
    gDbPath = "\\ECC\ZIP$\PersMan"
    gLog = True
    'gLogFile = "E:\Data\Projects\PersMan\Persman.log"
    gLogFile = "D:\WINNT\Profiles\administrator.000\Desktop\Persman\Persman.log"
    gWordViewPath = "D:\Program Files\Wordview\Wordview.exe"
    frmSplash.Show
End Sub
