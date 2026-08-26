VERSION 5.00
Begin VB.UserControl drLogon 
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2160
   ScaleWidth      =   4500
   ToolboxBitmap   =   "drLogon.ctx":0000
   Begin VB.Image imgPicture 
      Height          =   255
      Left            =   0
      Picture         =   "drLogon.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "drLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_Username = ""
Const m_def_Password = "password"
Const m_def_Domain = ""
Const m_def_Authenticated = 0
'Property Variables:
Dim m_Username As String
Dim m_Password As String
Dim m_Domain As String
Dim m_Authenticated As Boolean

Const LOGON32_LOGON_BATCH = 4
Const LOGON32_LOGON_INTERACTIVE = 2
Const LOGON32_LOGON_NETWORK = 3
Const LOGON32_LOGON_SERVICE = 5
Const LOGON32_PROVIDER_DEFAULT = 0
Const LOGON32_PROVIDER_WINNT35 = 1


Private Declare Function LogonUser Lib "Advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
'Private Declare Function LogonUser Lib "Advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As Any, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get UserName() As String
  UserName = m_Username
End Property

Public Property Let UserName(ByVal New_Username As String)
  m_Username = New_Username
  PropertyChanged "Username"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Password() As String
  Password = m_Password
End Property

Public Property Let Password(ByVal New_Password As String)
  m_Password = New_Password
  PropertyChanged "Password"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Domain() As String
  Domain = m_Domain
End Property

Public Property Let Domain(ByVal New_Domain As String)
  m_Domain = New_Domain
  PropertyChanged "Domain"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Logon() As Variant
  Authenticated = DoLogon(m_Username, m_Password, m_Domain)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Authenticated() As Boolean
  Authenticated = m_Authenticated
End Property

Public Property Let Authenticated(ByVal New_Authenticated As Boolean)
  m_Authenticated = New_Authenticated
  PropertyChanged "Authenticated"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_Username = m_def_Username
  m_Password = m_def_Password
  m_Domain = m_def_Domain
  m_Authenticated = m_def_Authenticated
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  m_Username = PropBag.ReadProperty("Username", m_def_Username)
  m_Password = PropBag.ReadProperty("Password", m_def_Password)
  m_Domain = PropBag.ReadProperty("Domain", m_def_Domain)
  m_Authenticated = PropBag.ReadProperty("Authenticated", m_def_Authenticated)
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = imgPicture.Width
  UserControl.Height = imgPicture.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Username", m_Username, m_def_Username)
  Call PropBag.WriteProperty("Password", m_Password, m_def_Password)
  Call PropBag.WriteProperty("Domain", m_Domain, m_def_Domain)
  Call PropBag.WriteProperty("Authenticated", m_Authenticated, m_def_Authenticated)
End Sub

Function DoLogon(strUserID As String, strPassword As String, Optional strDomain As String) As Boolean
  Dim lngToken As Long, lngRtn As Long
  If Len(strDomain) = 0 Then strDomain = vbNullString
  lngRtn = LogonUser(strUserID, strDomain, strPassword, LOGON32_LOGON_NETWORK, LOGON32_PROVIDER_DEFAULT, lngToken)
  If lngRtn = 0 Then
    DoLogon = False
  Else
    DoLogon = True
    CloseHandle lngToken
  End If
End Function

