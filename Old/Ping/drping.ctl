VERSION 5.00
Begin VB.UserControl drPing 
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1440
   ScaleWidth      =   1590
   ToolboxBitmap   =   "drPing.ctx":0000
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   0
      Picture         =   "drPing.ctx":0312
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "drPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_PingTime = 0
Const m_def_ReturnCode = 0
'Property Variables:
Dim m_PingTime As Long
Dim m_ReturnCode As Variant
Dim m_Ping As New clsPing

Private Sub UserControl_Resize()
    UserControl.Width = imgIcon.Width
    UserControl.Height = imgIcon.Height
End Sub

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

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ReturnCode = PropBag.ReadProperty("ReturnCode", m_def_ReturnCode)
    m_PingTime = PropBag.ReadProperty("PingTime", m_def_PingTime)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ReturnCode", m_ReturnCode, m_def_ReturnCode)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=8
'Public Function Ping(sIPAddress As String, Optional ByVal lTimeout As Long) As Long
'    p = m_Ping.Ping(sIPAddress, lTimeout)
    'Call PropBag.WriteProperty("PingTime", m_PingTime, m_def_PingTime)
'End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ReturnCode = m_def_ReturnCode
    m_PingTime = m_def_PingTime
End Sub
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,1,0
Public Property Get ReturnCode() As Variant
    ReturnCode = m_Ping.ReturnCode
End Property

Public Property Let ReturnCode(ByVal New_ReturnCode As Variant)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_ReturnCode = New_ReturnCode
    PropertyChanged "ReturnCode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Ping(sIPAddress As String, Optional ByVal lTimeout As Long) As Variant
    Ping = m_Ping.Ping(sIPAddress, lTimeout)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,1,0
Public Property Get PingTime() As Long
    PingTime = m_PingTime
End Property

Public Property Let PingTime(ByVal New_PingTime As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_PingTime = New_PingTime
    PropertyChanged "PingTime"
End Property

