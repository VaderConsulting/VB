VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ModifyEA"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim strDistinguishedName As String
    Dim intExtensionAttributeNumber As Integer
    Dim strNewAttributeValue As String
    Dim strTemp() As String
    Dim strCommand As String
    
    ' Remove chr(34)'s
    strCommand = Replace(Command$, Chr(34), "", 1, -1, vbTextCompare)
        
    strTemp = Split(strCommand, "*", -1, vbTextCompare)
    
    strDistinguishedName = strTemp(0)
    intExtensionAttributeNumber = strTemp(1)
    strNewAttributeValue = strTemp(2)
    
    If strDistinguishedName <> "" And intExtensionAttributeNumber <> 0 Then
        ModifyEA strDistinguishedName, intExtensionAttributeNumber, strNewAttributeValue
    End If
    
    ' Close application down
    End
End Sub

Private Sub ModifyEA(strDN As String, intAttributeNumber As Integer, strAttributeValue As String)
  Dim objUser As IADs
  
    Screen.MousePointer = vbHourglass

    Set objUser = GetObject("LDAP://" & strDN)
    
    On Error Resume Next ' Some attributes may not be present, and will generate an error
        
        If strAttributeValue = "" Then
            objUser.PutEx ADS_PROPERTY_CLEAR, "extensionAttribute" & intAttributeNumber, vbNull
        Else
            objUser.Put "extensionAttribute" & intAttributeNumber, strAttributeValue
        End If
        
        objUser.SetInfo
        
        Set objUser = Nothing
        
    On Error Resume Next

    Screen.MousePointer = vbDefault

End Sub
