VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim dacl As IADsAccessControlList
    Dim sd As IADsSecurityDescriptor
    Dim newACE As IADsAccessControlEntry
    Dim sdUtil As New ADsSecurityUtility
    
    Dim sFilename As String
    Dim sDomainName As String
    Dim sUsername As String
    
    sFilename = "c:\Downloads\New Folder"
    sDomainName = Environ$("COMPUTERNAME")
    sUsername = "test"
    
    Set sd = sdUtil.GetSecurityDescriptor(sFilename, ADS_PATH_FILE, ADS_SD_FORMAT_IID)
    Set dacl = sd.DiscretionaryAcl
    
    ' Add a new ACE
    Set newACE = CreateObject("AccessControlEntry")
    newACE.Trustee = sDomainName & "\" & sUsername
    newACE.AccessMask = GENERIC_ALL Or FILE_ALL_ACCESS
    newACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED
    newACE.AceFlags = ADS_ACEFLAG_INHERIT_ACE

    dacl.AddAce newACE
    sd.DiscretionaryAcl = dacl
    
    sdUtil.SetSecurityDescriptor sFilename, ADS_PATH_FILE, sd, ADS_SD_FORMAT_IID

    Set dacl = Nothing
    Set sd = Nothing
    Set newACE = Nothing
    Set sdUtil = Nothing
    End
End Sub

Private Sub AddACL_XP(strDomainname As String, strUsername As String, strPath As String, lAccessMask As Long, Optional bInherit As Boolean = False)
    Dim dacl As IADsAccessControlList
    Dim sd As IADsSecurityDescriptor
    Dim newACE As IADsAccessControlEntry
    Dim sdUtil As New ADsSecurityUtility
            
    Set sd = sdUtil.GetSecurityDescriptor(strPath, ADS_PATH_FILE, ADS_SD_FORMAT_IID)
    Set dacl = sd.DiscretionaryAcl
    
    ' Add a new ACE
    Set newACE = CreateObject("AccessControlEntry")
    newACE.Trustee = strDomainname & "\" & strUsername
    newACE.AccessMask = lAccessMask
    newACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED
    
    ' Inherit to child objects if applicable
    If bInherit Then
        newACE.AceFlags = ADS_ACEFLAG_INHERIT_ACE
    End If
    
    dacl.AddAce newACE
    sd.DiscretionaryAcl = dacl
    
    sdUtil.SetSecurityDescriptor strFilename, ADS_PATH_FILE, sd, ADS_SD_FORMAT_IID

    Set dacl = Nothing
    Set sd = Nothing
    Set newACE = Nothing
    Set sdUtil = Nothing

End Sub
