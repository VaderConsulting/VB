VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4395
      TabIndex        =   13
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1995
      TabIndex        =   12
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CheckBox chkLargeIcons 
      Caption         =   "Show Large Icons"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CheckBox chkTagFolders 
      Caption         =   "Tag Folders with DN"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CheckBox chkShowDLs 
      Caption         =   "Show Distribution Lists"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CheckBox chkShowAbViews 
      Caption         =   "Show Address Book Views"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CheckBox chkShowCustom 
      Caption         =   "Show Custom Recipients"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CheckBox chkShowUsers 
      Caption         =   "Show Users"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CheckBox chkShowGAL 
      Caption         =   "Show Global Address List"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   2415
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6588
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label lblLDAPPath 
      Caption         =   "LDAPPath"
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   5040
      Width           =   6735
   End
   Begin VB.Label Label4 
      Caption         =   "LDAP Path"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblFullPath 
      Caption         =   "Full Path"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   4680
      Width           =   6735
   End
   Begin VB.Label Label3 
      Caption         =   "Full Path"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Item Key"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Item Tag"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblKey 
      Caption         =   "Key"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   4320
      Width           =   6735
   End
   Begin VB.Label lblItemTag 
      Caption         =   "Tag"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   3960
      Width           =   6735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER = &H80000001

Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Dim bShowGal As Boolean
Dim bShowABView As Boolean
Dim bShowUsers As Boolean
Dim bShowCustom As Boolean
Dim bShowDLs As Boolean
Dim bTagFolder As Boolean
Dim bLargeIcons As Boolean

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Function ConvertDNToLDAP(strDN As String) As String
'DNs are in the format /o=org/ou=site/cn=container etc
'LDAP DN's are in the format /cn=container, ou=site, o=org etc
Dim strTempRev As String
Dim strTempDN As String
Dim I As Long
Dim bStop As Boolean
Dim strBuild As String
Dim strTempChunk As String
Dim lFoundAt As Long

    If Len(strDN) > 0 Then
'See if we have been passed an exhange address, if so strip of the start
        If InStr(1, strDN, "EX:") = 1 Then
'we have an exchange address, so strip it off
            strTempDN = Right$(strDN, Len(strDN) - 4)
        Else
'not an exchange address, so just strip off the first /
            strTempDN = Right$(strDN, Len(strDN) - 1)
        End If
'Reverse the text so we can use the standard instr functions
        For I = Len(strDN) To 1 Step -1 ' Reverse the text.
            strTempRev = strTempRev & Mid(strTempDN, I, 1)
        Next I
        bStop = False
        strBuild = ""
        Do
            lFoundAt = InStr(1, strTempRev, "/")
            If lFoundAt <> 0 Then
                strTempChunk = Right$(strTempDN, lFoundAt - 1)
            Else
                strTempChunk = strTempDN
                bStop = True
            End If
            strBuild = strBuild & strTempChunk & ", "
            strTempRev = Right$(strTempRev, Len(strTempRev) - lFoundAt)
            strTempDN = Left$(strTempDN, Len(strTempDN) - lFoundAt)
        Loop Until bStop = True
'trim off the last comma since we dont need it.
        strBuild = Left$(strBuild, Len(strBuild) - 2)
        ConvertDNToLDAP = strBuild
    Else
        ConvertDNToLDAP = ""
    End If
End Function

Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim I As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
    
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
    
        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
                tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
        Else                                                    ' WinNT Does NOT Null Terminate String...
                tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
        End If
        
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
    
        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
    
GetKeyError:
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Function FillMAPIFolderTree(TheTree As TreeView, Optional bShowGal As Boolean = False, _
    Optional bShowDLs As Boolean = False, Optional bShowUsers As Boolean = False, _
    Optional bShowCustom As Boolean = False, Optional bShowAbViews As Boolean = False, _
    Optional bTagFolder As Boolean = False, Optional bLargeIcons As Boolean = False) As Boolean
'Purpose: This function fills a treeview control with the MAPI folder structure.
'Comments: This walks the CDO supplied address book, and populates the treeview based on the
'           items found, and the flags supplied by the user.
'Parameters:
'TheTree      - The treeview to be populated
'bShowGAL     - Should the GAL be shown ?
'bShowDLs     - Should we show distribution lists ?
'bShowUsers   - Should we show users (mailboxes) ?
'bShowCustom  - Should we show custom recipients ?
'bShowAbViews - Should we show the Address Book views ?
'bTagFolder   - Should folders be tagged with the Exchange DN ?
'bLargeIcons  - Should we show using large icons ?

Dim strDefProfile As String
Dim OleMsg As MAPI.Session
Dim nodx As Node
Dim AddrLists As AddressLists
Dim AddrEntry As AddressList
Dim AddrContEntries As MAPI.AddressEntries
Dim AddrContEntry As MAPI.AddressEntry
Dim itmField As MAPI.Field
Dim strKey As String
Dim lCounter As Long
Dim strOrg As String
Dim strTreePath As String
Dim stopMe As Boolean
Dim FirstSlash As Long
Dim DNSlash As Long
Dim Level As Long
Dim strPrevitem As String
Dim strTempKey As String
Dim lDNFoundSlash As Long
Dim LoadIT As Boolean
Dim itmx As ListImage
Dim bAddedItem As Boolean
Dim strWhichIcon As String
Dim nody As Node
Dim strThisLevel As String

'Retrieve the default mapi profile from the registry. We use resource strings 1500 and 1501, which contain
' the paths where the item is located in the registry
    If GetKeyValue(HKEY_CURRENT_USER, LoadResString(1500), LoadResString(1501), strDefProfile) Then

    Else
        MsgBox "Error Getting Profile", vbExclamation, "Error"
'Setting this to an empty stirng will force mapi to pop up a dialog box asking for a profile
        strDefProfile = ""
    End If

    Set OleMsg = CreateObject("MAPI.Session")
'Since we will probablt be ading noes that have already been added, we use this to skip over the error
' Need a better way to handle this.
    On Error Resume Next
    OleMsg.Logon strDefProfile, "", True, True, 0

    If Err.Number <> 0 Then
        FillMAPIFolderTree = False
        Exit Function
    End If

    TheTree.ImageList.ListImages.Clear
'Set the icons to small ones
    If bLargeIcons = False Then
        TheTree.ImageList.ImageHeight = 16
        TheTree.ImageList.ImageWidth = 16
'Set a smaller indentation
        TheTree.Indentation = 300
    End If
'Fill the image list with icons from the resource file
    Set itmx = TheTree.ImageList.ListImages.Add(, "Container", LoadResPicture(1, vbResIcon))
    Set itmx = TheTree.ImageList.ListImages.Add(, "Custom", LoadResPicture(2, vbResIcon))
    Set itmx = TheTree.ImageList.ListImages.Add(, "RecipDistList", LoadResPicture(3, vbResIcon))
    Set itmx = TheTree.ImageList.ListImages.Add(, "WorldMail", LoadResPicture(4, vbResIcon))
    Set itmx = TheTree.ImageList.ListImages.Add(, "Site", LoadResPicture(5, vbResIcon))
    Set itmx = TheTree.ImageList.ListImages.Add(, "ExGal", LoadResPicture(6, vbResIcon))
    Set itmx = TheTree.ImageList.ListImages.Add(, "Mailbox", LoadResPicture(7, vbResIcon))
    Set itmx = TheTree.ImageList.ListImages.Add(, "ABView", LoadResPicture(8, vbResIcon))


'Clean out the old treeview
    TheTree.Nodes.Clear
'Loop throught all of the available address lists
    Set AddrLists = OleMsg.AddressLists
    lCounter = 1
    For Each AddrEntry In AddrLists
'Decide if we should load the item
'We hide the personal address list since you can't delete /create dl's directly in them
        If AddrEntry.Name = "Personal Address Book" Then
            LoadIT = False
        ElseIf (AddrEntry.Name = "Global Address List" And bShowGal = False) Then
            LoadIT = False
        Else
            LoadIT = True
        End If

        If LoadIT = True Then
'Loop thorugh the extended fields for the object and look for -2143551458 which
'equals the DN of the container and then add this as the key value for the field
'so we can extract it later
            Level = 1
            stopMe = False
            strKey = ""
            bAddedItem = False
'Firstly find the dn path for this item
            For Each itmField In AddrEntry.Fields
                If itmField.ID = -2143551458 Then
                    strKey = CStr(itmField.Value)
                    Exit For
                End If
            Next

            For Each itmField In AddrEntry.Fields
'loop through looking for -458722 which is the path in the MAPI tree of the object
                If itmField.ID = -458722 Then
'We need to fudge address book views because the appear under the wrong level
                    If InStr(1, strKey, "_ABViews_") <> 0 Then
'Get the string
                        strTreePath = itmField.Value
'Find out where the next slash is after the organisation
                        FirstSlash = InStr(2, itmField.Value, "\")
'Insert "Address Book Views"
                        strTreePath = Left$(itmField.Value, FirstSlash) & "Address Book Views\" & Right$(itmField.Value, Len(itmField.Value) - FirstSlash)
                    Else
'Otherwise just use the field as is.
                        strTreePath = itmField.Value
                    End If
                    strPrevitem = ""
                    strTreePath = Right$(strTreePath, Len(strTreePath) - 1)
                    FirstSlash = 1
                    lDNFoundSlash = 2
                    Do
                        bAddedItem = False
                        strWhichIcon = ""
'We loop here, pulling apart the path bit by bit, and add each bit
' progressivly, which gives our our actual tree !

'Since each object has a DN, we split that off as well, so the DN when added as the key
'value for the field will match it's path in the tree
                        DNSlash = InStr(lDNFoundSlash, strKey, "/")
                        If DNSlash <> 0 Then
                            strTempKey = Left$(strKey, DNSlash - 1)
                        Else
                            strTempKey = strKey
                        End If
'Pull apart the path
                        FirstSlash = InStr(1, strTreePath, "\")
                        If FirstSlash <> 0 Then
                            strOrg = Left$(strTreePath, FirstSlash - 1)
                        Else
'No slash, assume the rest of the string is the item to add
                            strOrg = strTreePath
                        End If
'Since we may be attempting to re-add an item which has already been added, just skip
' if we get an error
                        On Error Resume Next
'We set the key to include the current level because it is unlikely, but possible to
' come up with a duplicate (i'm not sure of the odds).
                        strThisLevel = CStr(Level) & strOrg
'See if we have already added a previous item. The first time this is called for each item in the address list
' it will be empty, and it is populated after each level is added
                        If strPrevitem <> "" Then
'See if the key (DN) has _ABViews_ in it. If so it is a Address book item. Address book
' items use the same icon for each level, so we don't bother checking which level
' the item is in, so we just set it to the ABView icon.
                            If (InStr(1, strKey, "_ABViews_") <> 0) Then
                                If bShowAbViews = True Then
                                    strWhichIcon = "ABView"
                                End If
                            Else
'Check what level we are. If it is 2 then we are at the site / container level in exchange
                                If Level = 2 Then
                                    strWhichIcon = "Site"
                                Else
                                    strWhichIcon = "Container"
                                End If
                            End If
'Only add the item if an icon has been specified
                            If strWhichIcon <> "" Then
                                Set nodx = TheTree.Nodes.Add(strPrevitem, tvwChild, strPrevitem & strThisLevel, strOrg, strWhichIcon)
                                If Err.Number = 0 Then
                                    bAddedItem = True
                                End If
                            End If
                        Else
'Since there is no previous item, we must be a root node
                            If strOrg = "Global Address List" Then
'If we are the GAL, use a different icon like exchange does
                                strWhichIcon = "ExGal"
                            Else
                                strWhichIcon = "WorldMail"
                            End If
                            Set nodx = TheTree.Nodes.Add(, , strThisLevel, strOrg, strWhichIcon)
                            If Err.Number = 0 Then
                                bAddedItem = True
                            End If
                        End If
'Should we set the tag for the folder ? Only do it if the item was added, otherwise we
'have problems with the tags
                        If (bAddedItem = True And bTagFolder = True) Then
                            nodx.Tag = strTempKey
                        End If
'Save the previtem, so that the tree is structured properly
                        strPrevitem = strPrevitem & strThisLevel
'Trim off the path we just added.
                        strTreePath = Right$(strTreePath, Len(strTreePath) - FirstSlash)
'increment the level, since we are now moving down into the tree
                        Level = Level + 1
'increment the distinguished name level as well. We only do this, instead of trimming
' like the
                        lDNFoundSlash = DNSlash + 1
'If firstslash=0 then we are at an endpoint, so load an items that are located in
' this container
                        If FirstSlash = 0 Then
'Go through all entries for this item
                            Set AddrContEntries = AddrEntry.AddressEntries
                            For Each AddrContEntry In AddrContEntries
'Decide what type the item is, and set the icon and tag accordingly
                                bAddedItem = False
                                If (AddrContEntry.DisplayType = CdoDistList And bShowDLs = True) Then
                                    Set nody = TheTree.Nodes.Add(strPrevitem, tvwChild, , AddrContEntry.Name, "RecipDistList")
                                    bAddedItem = True
                                ElseIf (AddrContEntry.DisplayType = CdoUser And bShowUsers = True) Then
                                    Set nody = TheTree.Nodes.Add(strPrevitem, tvwChild, , AddrContEntry.Name, "Mailbox")
                                    bAddedItem = True
                                ElseIf (AddrContEntry.DisplayType = CdoRemoteUser And bShowCustom = True) Then
                                    Set nody = TheTree.Nodes.Add(strPrevitem, tvwChild, , AddrContEntry.Name, "Custom")
                                    bAddedItem = True
                                End If

                                If bAddedItem = True Then
                                    nody.Tag = AddrContEntry.Address
                                End If
                            Next
'Cleanup
                            Set AddrContEntries = Nothing
'Since we have gone down this item as far as we can, we should stop and get the next item
                            stopMe = True
                        End If
'End of the loop for each item
                    Loop Until stopMe = True
'Since we found the field we are looking for, break the loop
                    Exit For
                End If
'End of the loop for each field in the address entry
            Next
        End If
        lCounter = lCounter + 1
'End of the loop for each item in the address list
    Next
    If OleMsg Is Nothing Then
'No session, so don't bother
    Else
        OleMsg.Logoff
    End If
End Function


Private Sub chkLargeIcons_Click()
bLargeIcons = Not bLargeIcons
End Sub

Private Sub chkShowAbViews_Click()
bShowABView = Not bShowABView
End Sub

Private Sub chkShowCustom_Click()
bShowCustom = Not bShowCustom
End Sub

Private Sub chkShowDLs_Click()
bShowDLs = Not bShowDLs
End Sub

Private Sub chkShowGAL_Click()
bShowGal = Not bShowGal
End Sub

Private Sub chkShowUsers_Click()
bShowUsers = Not bShowUsers
End Sub

Private Sub chkTagFolders_Click()
bTagFolder = Not bTagFolder
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRefresh_Click()
'This refills the list using the selected options
Call FillMAPIFolderTree(TreeView1, bShowGal, bShowDLs, bShowUsers, bShowCustom, bShowABView, bTagFolder, bLargeIcons)
End Sub

Private Sub Form_Load()
bShowGal = False
bShowABView = False
bShowUsers = False
bShowCustom = False
bShowDLs = False
bTagFolder = False
bLargeIcons = False
End Sub


Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
lblFullPath = Node.FullPath
lblItemTag = Node.Tag
lblKey = Node.Key
'Convert the DN (if available) to an LDAP path
lblLDAPPath = ConvertDNToLDAP(Node.Tag)
End Sub
