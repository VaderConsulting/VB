VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Users"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim adoConn As ADODB.Connection
  Dim adoRS As ADODB.Recordset
  Dim SQL As String
  Dim DSN As String
  Dim SQLServer As String
  Dim Database As String
  Dim oPDC As IADsContainer
  Dim oUser As IADsUser
  Dim LocalUserPath As String
  Dim RemoteUserPath As String
  Dim fso As Object
  Dim f As Object
  Dim sec As Object
  Dim sd As Object
  Dim acl As Object
  Dim aceB As Object
  Dim CommandLine As String
  Dim Groups() As String
  Dim Group As IADsGroup
  Dim ReturnVal As Long
  Dim msg As String
  Dim objMail As Object
  Dim UserExists As Boolean
  Dim HomeShareExists As Boolean
  Dim PermError As Boolean
  Dim PermErrorCode As Long
  Dim HomeExists As Boolean
  Dim ProfileExists As Boolean
  
  Set adoConn = CreateObject("ADODB.Connection")
  Set adoRS = CreateObject("ADODB.Recordset")
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set WshShell = CreateObject("WScript.Shell")
    
  Me.Show
  Me.Refresh
  Screen.MousePointer = vbHourglass
  SQLServer = "CBDXAAI"
  Database = "POLICE"
  LocalUserPath = "E:\Users"
  
  DSN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & Database & ";Data Source=" & SQLServer
  SQL = "SELECT * FROM tblCreateUsers WHERE complete = 0"
  
  adoConn.Open DSN
  adoRS.Open SQL, adoConn
  Do Until adoRS.EOF Or adoRS.BOF
    ' Create User
    Set oPDC = GetObject("WinNT://POLICE/CBDXAAA,computer")
    Set oUser = oPDC.Create("user", adoRS("Username"))
    lblInfo = "Creating account for " & adoRS("Username")
    lblInfo.Refresh
    On Error Resume Next
    oUser.SetInfo
    If Err.Number = -2147022672 Then
      UserExists = True
    Else
      UserExists = False
      lblInfo = "Setting info for " & adoRS("Username")
      lblInfo.Refresh
      oUser.SetPassword (adoRS("Password"))
      oUser.FullName = adoRS("Last") & " " & adoRS("First")
      oUser.Description = "Creation Ref:  " & adoRS("CallNo")
      oUser.Profile = adoRS("ProfileDir")
      oUser.HomeDirectory = adoRS("HomeDir")
      oUser.LoginScript = adoRS("Script")
      oUser.Put "PasswordExpired", CLng(1)
      oUser.homedirdrive = "H:"
      oUser.SetInfo
      
      ' Put user into specified groups
      Groups = Split(adoRS("Groups"), "|")
      For a = 0 To UBound(Groups)
        Set Group = GetObject("WinNT://POLICE/CBDXAAA/" & Groups(a))
        lblInfo = "Adding " & adoRS("Username") & " to " & Groups(a)
        lblInfo.Refresh
        Group.Add oUser.ADsPath
        Set Group = Nothing
      Next a
      Group.SetInfo
      Erase Groups ' remove all elements of array 'groups'
      
      ' Create Home Directory
      If adoRS("HomeDir") <> "" Then
        lblInfo = "Creating Home Directory for " & adoRS("Username")
        lblInfo.Refresh
        RemoteUserPath = adoRS("HomeDir")
        RemoteUserPath = Replace(RemoteUserPath, "\\", "")
        p = InStr(1, RemoteUserPath, "\")
        Server = Left(RemoteUserPath, p - 1)
        FolderPath = "\\" & Server & "\E$\Users\" & adoRS("Username")
        On Error Resume Next
        Set f = fso.CreateFolder(FolderPath)
        If Err.Number = 58 Then
          HomeExists = True
        Else
          HomeExists = False
        End If
        On Error GoTo 0
        
        ' Create Share for User's Home Directory
        strSharePath = LocalUserPath & "\" & adoRS("Username")
        strShareDescription = adoRS("Username") & " " & "'s Home Folder"
        Set FServer = GetObject("WinNT://" & Server & "/lanmanserver")
        On Error Resume Next
        Set newShare = FServer.Create("fileshare", adoRS("Username") & "$")
        If Err.Number = 0 Then
          HomeShareExists = False
          newShare.Path = strSharePath
          newShare.Description = adoRS("Username") & " " & "'s Home Folder"
          newShare.MaxUserCount = -1
          newShare.SetInfo
        Else
          If Err.Number = -2147467259 Then 'Share Exists
            HomeShareExists = True
          End If
        End If
        On Error GoTo 0
        Set FServer = Nothing
        Set newShare = Nothing
        lblInfo = "Modifying Home Directory Permissions for " & adoRS("Username")
        lblInfo.Refresh
        
        '--- Add NTFS permissons to the folder ---
        CommandLine = "cmd /c ""echo y > xcacls " & FolderPath & " /g " & adoRS("Username") & ":C"""
        On Error GoTo 0
        Shell CommandLine, vbMinimizedNoFocus
        If ReturnVal <> 0 Then
          ' Error modifying perms
          PermError = True
        Else
          PermError = False
        End If
        PermErrorCode = ReturnVal
      End If
      
      ' Create Profile Directory
      If adoRS("ProfileDir") <> "" Then
        lblInfo = "Creating Profile Directory for " & adoRS("Username")
        lblInfo.Refresh
        RemoteUserPath = adoRS("ProfileDir")
        RemoteUserPath = Replace(RemoteUserPath, "\\", "")
        p = InStr(1, RemoteUserPath, "\")
        Server = Left(RemoteUserPath, p - 1)
        FolderPath = "\\" & Server & "\E$\Profiles\" & adoRS("Username")
        On Error Resume Next
        Set f = fso.CreateFolder(FolderPath)
        If Err.Number = 58 Then
          ProfileExists = True
        Else
          ProfileExists = False
        End If
        On Error GoTo 0
        ' No special perms for profile folder
      End If
    End If
    On Error GoTo 0
    
    ' Send Mail
    Set objMail = CreateObject("CDONTS.Newmail")
    objMail.to = Left(adoRS("RequestUser"), 7) & "@police.wa.gov.au"
    objMail.From = "CBDXAAI@nowhere.com"
    objMail.subject = "Notification from Create User Application"
    msg = ""
    If Not UserExists Then
      msg = msg & adoRS("Username") & " was created on the POLICE domain at " & Now & vbCrLf
    Else
      msg = msg & adoRS("Username") & " could not be created on the POLICE domain at " & Now & " because this username already exists." & vbCrLf
    End If
    msg = msg & "You received this email because you requested this user to be created at " & adoRS("DateTime") & "." & vbCrLf
    If HomeShareExists Then
      msg = msg & "The users' Home share was not created as it already existed." & vbCrLf
    End If
    If HomeExists Then
      msg = msg & "The users' Home drive folder was not created as it already existed." & vbCrLf
    End If
    If ProfileExists Then
      msg = msg & "The users' Profile folder was not created as it already existed." & vbCrLf
    End If
    If PermError Then
      msg = msg & "Permissions on the users' Home share were not set due to error " & PermErrorCode & "." & vbCrLf
    End If
    msg = msg & vbCrLf
    msg = msg & "Your original request supplied the following information:" & vbCrLf
    For Each Field In adoRS.Fields
      msg = msg & Format(Field.Name, "               ") & ": " & Field.Value & vbCrLf
    Next
    
    objMail.body = msg
    objMail.Send
    Set objMail = Nothing
    
    ' Update DB
    SQL = "UPDATE tblCreateUsers SET Complete=1, Password='' where ID = " & adoRS("ID")
    adoConn.Execute SQL
    adoRS.MoveNext
    DoEvents
    Set oUser = Nothing
  Loop
  
  adoRS.Close
  adoConn.Close
  Set adoRS = Nothing
  Set adoConn = Nothing
  Screen.MousePointer = vbDefault
  lblInfo = "Idle"
  lblInfo.Refresh
  Unload Me
  End
End Sub
