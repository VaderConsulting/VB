VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Service Accounts"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000

Private Sub Form_Load()
    Dim oPDC As IADsContainer
    Dim oUser As IADsUser
    Dim UserExists As Boolean
    Dim Username As String
    Dim Password As String
    Dim Fullname As String
    Dim Flag As Long
    Dim newFlag As Long
    Dim AppPath As String
    Dim ADPath As String
    Dim RetDesc As String
    
    Me.Show
    Me.Refresh
    Screen.MousePointer = vbHourglass
    
    AppPath = App.Path
    If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    
    ADPath = "WinNT://"
    
    ADPath = ADPath & "JUSTICE/EPDCCSDOM02"
    'ADPath = ADPath & "TUSK/DELLP1000"
    ADPath = ADPath & ",computer"
    
    Open AppPath & "sms.csv" For Input As #1
    Open AppPath & "sms.log" For Output As #2
    On Error Resume Next
    Set oPDC = GetObject(ADPath)
    Print #2, "GetObject returns " & Err.Number & " (" & Err.Description & ")"
    
    Do Until EOF(1)
        Input #1, Username
        Input #1, Password
        Input #1, Fullname
        Print #2, "Working with " & Username
        Set oUser = oPDC.Create("user", Username)
        Print #2, "Create User returns " & Err.Number & " (" & Err.Description & ")"
        oUser.SetInfo
        Set oUser = GetObject(ADPath & "/" & Username)
        oUser.SetPassword Password
        oUser.Fullname = Fullname
        oUser.Description = "DO NOT MODIFY.  Systems Management Server Internal Account"
        oUser.Profile = ""
        oUser.HomeDirectory = ""
        oUser.LoginScript = ""
        oUser.SetInfo
        Flag = oUser.Get("UserFlags")
        newFlag = Flag Or ADS_UF_DONT_EXPIRE_PASSWD
        oUser.Put "userFlags", newFlag
        oUser.SetInfo
        Print #2, "Created User " & Username
        DoEvents
        Set oUser = Nothing
    Loop
    Set oPDC = Nothing
    
    Close 1
    Close 2
    Screen.MousePointer = vbDefault
    MsgBox "Complete.  Review " & AppPath & "sms.log for further info"
End Sub

