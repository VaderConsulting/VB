VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Policy Query Tool"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtValue 
      Height          =   645
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4320
      Width           =   3135
   End
   Begin VB.ListBox lstUsers 
      Height          =   3765
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdGetUser 
      Caption         =   "Query"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblAttribute 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label lblValue 
      Caption         =   "Value"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type UserViewedPolicy
    PolicyName As String
    Retries As Integer
    AttributeNumber As Integer
End Type

Dim UserPolicies() As UserViewedPolicy

'=======================================================================================
' Procedure : cmdGetUser_Click
' DateTime  : 14/10/2003 15:18
' Author    : Dave Robinson
' Returns   :
' Purpose   : Process the command click
'=======================================================================================
'  V    Date            Author                 History
' 1.0   14/10/2003      Dave Robinson          Initial Version
'=======================================================================================
Private Sub cmdGetUser_Click()
    Dim EA() As String
    Dim intLoop As Integer
    
    frmMain.lstUsers.Clear
    frmMain.txtValue = ""
    
    GetDetails Environ$("USERDOMAIN"), txtUsername.Text, EA()
    
    For intLoop = 1 To 15
        If EA(intLoop) <> "" Then
            frmMain.lstUsers.AddItem intLoop & ": " & EA(intLoop)
        End If
    Next intLoop
End Sub

'=======================================================================================
' Procedure : cmdOK_Click
' DateTime  : 14/10/2003 15:18
' Author    : Dave Robinson
' Returns   :
' Purpose   : Process the command button click
'=======================================================================================
'  V    Date            Author                 History
' 1.0   14/10/2003      Dave Robinson          Initial Version
'=======================================================================================
Private Sub cmdOK_Click()
    SetDetails CInt(lblAttribute.Caption), Environ$("USERDOMAIN"), txtUsername.Text, txtValue.Text
End Sub

'=======================================================================================
' Procedure : GetDetails
' DateTime  : 13/10/2003 08:29
' Author    : Dave Robinson
' Returns   : Boolean
' Purpose   : Get the user extensionAttributes into an array (strEA())
'=======================================================================================
'  V    Date            Author                 History
' 1.0   13/10/2003      Dave Robinson          Initial Version
'=======================================================================================
Private Function GetDetails(strDomainname As String, strUsername As String, strEA() As String) As Boolean
  Dim adoConn As ADODB.Connection
  Dim strADRoot As String
  Dim strFilter As String
  Dim adoCommand As ADODB.Command
  Dim adoUsers As ADODB.Recordset
  Dim intRecordCount As Integer
  Dim objUser As IADs
  Dim intLoop As Integer
  
  ReDim strEA(15)

  'On Error Resume Next

    Screen.MousePointer = vbHourglass
    Set adoConn = CreateObject("ADODB.Connection")
    adoConn.Provider = "ADsDSOObject"
    adoConn.Open "Active Directory Provider"

    Set adoCommand = CreateObject("ADODB.Command")
    adoCommand.ActiveConnection = adoConn

    strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & strUsername & "))"

    strADRoot = "LDAP://" & strDomainname

    adoCommand.CommandText = "<" & strADRoot & ">;" & strFilter & _
                             ";name,distinguishedName,objectGuid;subTree"
    Set adoUsers = adoCommand.Execute
    intRecordCount = adoUsers.RecordCount

    If intRecordCount = 0 Then
        GetDetails = False
      Else 'NOT INTRECORDCOUNT...
        strDN = adoUsers.Fields("distinguishedName")
        Set objUser = GetObject("LDAP://" & strDN)
        
        On Error Resume Next ' Some attributes may not be present, and will generate an error

        For intLoop = 0 To 15
            strEA(intLoop) = objUser.Get("extensionAttribute" & CStr(intLoop))
        Next intLoop
        
        On Error Resume Next
        GetDetails = True
    End If

    Set adoUsers = Nothing
    adoConn.Close
    Screen.MousePointer = vbDefault

End Function

'=======================================================================================
' Procedure : lstUsers_Click
' DateTime  : 14/10/2003 15:18
' Author    : Dave Robinson
' Returns   :
' Purpose   : Process the Listbox click
'=======================================================================================
'  V    Date            Author                 History
' 1.0   14/10/2003      Dave Robinson          Initial Version
'=======================================================================================
Private Sub lstUsers_Click()
    Dim strValue As String
    
    strValue = lstUsers.List(lstUsers.ListIndex)
    
    If strValue <> "" Then
        lblAttribute.Caption = Left(strValue, InStr(1, strValue, ":", vbTextCompare) - 1)
        txtValue.Text = Mid(strValue, 4)
    End If
    
    SplitAttribute txtValue.Text, UserPolicies()
    
    'frmMain.lstAttributes.Clear
    'frmMain.tabPolicies.Tabs = UBound(UserPolicies()) + 1
    'frmMain.tabPolicies.TabsPerRow = 3
    
    'For intLoop = 0 To UBound(UserPolicies())
        'frmMain.lstAttributes.AddItem UserPolicies(intLoop).PolicyName
        'frmMain.lstAttributes.AddItem UserPolicies(intLoop).Retries
        'frmMain.tabPolicies.TabCaption(intLoop) = UserPolicies(intLoop).PolicyName
    'Next intLoop
End Sub

'=======================================================================================
' Procedure : SetDetails
' DateTime  : 14/10/2003 15:18
' Author    : Dave Robinson
' Returns   :
' Purpose   : Set the user details in AD
'=======================================================================================
'  V    Date            Author                 History
' 1.0   14/10/2003      Dave Robinson          Initial Version
'=======================================================================================
Private Sub SetDetails(intAttribute As Integer, strDomainname As String, strUsername As String, strValue As String)
  Dim adoConn As ADODB.Connection
  Dim strADRoot As String
  Dim strFilter As String
  Dim adoCommand As ADODB.Command
  Dim adoUsers As ADODB.Recordset
  Dim intRecordCount As Integer
  Dim objUser As IADs
  
  ReDim strEA(15)

  'On Error Resume Next

    Screen.MousePointer = vbHourglass
    Set adoConn = CreateObject("ADODB.Connection")
    adoConn.Provider = "ADsDSOObject"
    adoConn.Open "Active Directory Provider"

    Set adoCommand = CreateObject("ADODB.Command")
    adoCommand.ActiveConnection = adoConn

    strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & strUsername & "))"

    strADRoot = "LDAP://" & strDomainname

    adoCommand.CommandText = "<" & strADRoot & ">;" & strFilter & _
                             ";name,distinguishedName,objectGuid;subTree"
    Set adoUsers = adoCommand.Execute
    intRecordCount = adoUsers.RecordCount

    If intRecordCount = 0 Then
        ' Cant find user
      Else 'NOT INTRECORDCOUNT...
        strDN = adoUsers.Fields("distinguishedName")
        Set objUser = GetObject("LDAP://" & strDN)
        
        On Error Resume Next ' Some attributes may not be present, and will generate an error
            
            If strValue = "" Then
                objUser.PutEx ADS_PROPERTY_CLEAR, "extensionAttribute" & CStr(intAttribute), vbNull
            Else
                objUser.Put "extensionAttribute" & CStr(intAttribute), strValue
            End If
            
            objUser.SetInfo
            
        On Error Resume Next
    End If

    Set adoUsers = Nothing
    
    txtValue.Text = ""
    lblAttribute.Caption = ""
    
    adoConn.Close
    Screen.MousePointer = vbDefault

End Sub

'=======================================================================================
' Procedure : SplitAttribute
' DateTime  : 14/10/2003 15:18
' Author    : Dave Robinson
' Returns   :
' Purpose   : Split the given Attribute value into manageable pieces
'=======================================================================================
'  V    Date            Author                 History
' 1.0   14/10/2003      Dave Robinson          Initial Version
'=======================================================================================
Private Function SplitAttribute(strAttributeValue As String, strResult() As UserViewedPolicy)
    Dim strPolicyDetails() As String
    Dim intLoop As Integer
    Dim strDetail() As String
    ' UserPolicies is a custom type with three attributes:
    '     PolicyName As String
    '     Retries As Integer
    '     AttributeNumber As Integer
    '
    ' This is used to store the list of policies
    
    On Error Resume Next ' Sometimes there will be no information.  This would generate an error
    
    ' strAttributeValue is the complete list of policies this user has viewed.
    ' eg RTWA_Policy2003-10,3|LogOff_Notice,1|Reboot_Notice,3
    
    strPolicyDetails() = Split(strAttributeValue, "|")
    
    For intLoop = 0 To UBound(strPolicyDetails())
        strDetail() = Split(strPolicyDetails(intLoop), ",")
        ReDim Preserve strResult(intLoop)
        strResult(intLoop).AttributeNumber = intLoop
        strResult(intLoop).PolicyName = strDetail(0)
        strResult(intLoop).Retries = strDetail(1)
    Next intLoop
    
End Function
