VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNetErr"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" _
   (lpError As Long, _
    ByVal lpErrorBuf As String, _
    ByVal nErrorBufSize As Long, _
    ByVal lpNameBuf As String, _
    ByVal nNameBufSize As Long) As Long

' These values can be set, but it's pointless
' they are intended only to be read

Public dwErrorCode&   ' Network specific error code
Public szDescription$ ' Description of the error
Public szName$        ' Name of the network provider

Private Sub Class_Initialize()
    Dim lRC As Long
    ' Preinitialize the strings (as always)
    szDescription = String$(512, 0)
    szName = String$(512, 0)
    Call WNetGetLastError(dwErrorCode, szDescription, 511, szName, 511)
    ' Strip off everything after the null terminator
    lRC = lstrcpy(szDescription, szDescription)
'    szDescription = agGetStringFromLPSTR(szDescription)
    lRC = lstrcpy(szName, szName)
'    szName = agGetStringFromLPSTR(szName)
End Sub

