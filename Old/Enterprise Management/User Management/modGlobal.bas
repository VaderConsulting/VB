Attribute VB_Name = "modGlobal"
Type ScreenObject
    Index As Integer
    Name As String
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer
End Type

Public Enum ADS_USER_FLAG_ENUM
  ADS_UF_SCRIPT = &H1
  ADS_UF_ACCOUNTDISABLE = &H2
  ADS_UF_HOMEDIR_REQUIRED = &H3
  ADS_UF_LOCKOUT = &H10
  ADS_UF_PASSWD_NOTREQD = &H20
  ADS_UF_PASSWD_CANT_CHANGE = &H40
  ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H80
  ADS_UF_TEMP_DUPLICATE_ACCOUNT = &H100
  ADS_UF_NORMAL_ACCOUNT = &H200
  ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = &H800
  ADS_UF_WORKSTATION_TRUST_ACCOUNT = &H1000
  ADS_UF_SERVER_TRUST_ACCOUNT = &H2000
  ADS_UF_DONT_EXPIRE_PASSWD = &H10000
  ADS_UF_MNS_LOGON_ACCOUNT = &H20000
  ADS_UF_SMARTCARD_REQUIRED = &H40000
  ADS_UF_TRUSTED_FOR_DELEGATION = &H80000
  ADS_UF_NOT_DELEGATED = &H100000
End Enum

Public Const LB_SETTABSTOPS As Long = &H192

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public gstrDomain As String
Public gstrPDC As String

Public Sub SetListTabStops(ListHandle As Long, ParamArray ParmList() As Variant)
    Dim i As Long
    Dim ListTabs() As Long
    Dim NumColumns As Long

    ReDim ListTabs(UBound(ParmList))
    For i = 0 To UBound(ParmList)
        ListTabs(i) = ParmList(i) * 4 ' rough estimate.. 4 x parameter is tabstop
    Next i
    NumColumns = UBound(ParmList) + 1

    Call SendMessage(ListHandle, LB_SETTABSTOPS, NumColumns, ListTabs(0))
End Sub


