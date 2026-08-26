VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim clsVBRasEntry As VBRasEntry
    Dim clsVBRasDevInfo() As VBRASDEVINFO
    Dim MyCountry As VBRASCTRYINFO
    Dim Modems() As String
    Dim rtn As Long
    Dim i As Integer
    Dim j As Integer
    
    ' Get Country code
    MyCountry.CountryID = 61
    rtn = VBRasGetCountryInfo(MyCountry)
    
    ' Get RAS Devices
    rtn = VBRasEnumDevices(clsVBRasDevInfo)
    
    ' Determine RAS Devices
    For j = 0 To rtn - 1
        If clsVBRasDevInfo(j).DeviceType = "modem" Then
            i = i + 1
            ReDim Preserve Modems(i)
            Modems(i) = clsVBRasDevInfo(j).DeviceName
        End If
    Next j
    
    ' Set RAS DUN Properties
    If i > 0 Then
        clsVBRasEntry.AreaCode = "08"
        clsVBRasEntry.LocalPhoneNumber = "012345678"
        clsVBRasEntry.fNetProtocols = RASNP_Ip
        clsVBRasEntry.DeviceName = Modems(1)
        clsVBRasEntry.CountryCode = MyCountry.CountryCode
        clsVBRasEntry.CountryID = MyCountry.CountryID
        clsVBRasEntry.Options = RASEO_ModemLights + RASEO_NetworkLogon + RASEO_PreviewDomain + RASEO_PreviewPhoneNumber + RASEO_PreviewUserPw + RASEO_ShowDialingProgress
        
        rtn = VBRasSetEntryProperties("MyConnection", clsVBRasEntry)
        If rtn <> 0 Then MsgBox VBRASErrorHandler(rtn)
    End If
End Sub
