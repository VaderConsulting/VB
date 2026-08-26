VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outlook Analyzer"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objSession As MAPI.Session
     
Private Sub cmdConnect_Click()
    ' SQL = "Exchange 4.0;MAPILEVEL= storage|folderpath;TABLETYPE={0|1}; DATABASE=path"  ' ;[PROFILE=profile;PWD=password;]
    ' Source Table eg     :     "Mailbox – Pat Smith|Inbox\Big Project"
    ' Database eg         :     "h:\db1.mdb"
    ' Source Tabletype eg :     "TABLETYPE=0" ' = folders
    ' Source Tabletype eg :     "TABLETYPE=1" ' = address books
    ' ie
    '    SQL = "Exchange 4.0;MAPILEVEL= Mailbox - ROBINSON Dave [PD71044]|;TABLETYPE=0; DATABASE=h:\db1.mdb"
    
    Dim objInfoStore As InfoStore
    Set objSession = CreateObject("mapi.session")
    objSession.Logon
    Set objInfoStore = objSession.InfoStores("Public Folders")
    If objInfoStore.Fields(&H6632000B).Value Then  'PR_STORE_OFFLINE
        MsgBox "Store is offline", , "Offline Status"
    Else
        MsgBox "Store is online", , "Offline Status"
    End If
End Sub

Private Sub Form_Load()
    Set objSession = CreateObject("MAPI.Session")
    objSession.Logon showdialog:=False, newsession:=False
End Sub
     
Private Sub Command1_Click()
    Dim objPSTFolder As Folder
    Dim objMessages As Messages
    Dim objOneMessage As Message
    
    'Want to find the Messaging folder in my PST. Modify the "Messaging"
    'as needed to find the folder you are interested in.
    Set objPSTFolder = objFindTargetFolder("Top of Personal Folders", "Messaging")
    Set objMessages = objPSTFolder.Messages
    Set objOneMessage = objMessages.GetFirst
End Sub

Private Function objFindTargetFolder(strTargetTopFolder As String, strSearchName As String) As Folder
    Dim objInfoStores As InfoStores
    Dim objInfoStore As InfoStore
    Dim objTopFolder As Folder
    Dim objPSTFolders As Folders
    Dim i As Integer
    
    Set objInfoStores = objSession.InfoStores
    'This message box loop is just telling you what the current
    'objects are. It is not needed for the function to work
    'properly.
    For i = 1 To objInfoStores.Count
        Set objInfoStore = objInfoStores(i)
        Set objTopFolder = objInfoStore.RootFolder
        MsgBox "i= " & i & Chr(10) & "InfoStore.Name= " & objInfoStores(i).Name & Chr(10) & "TopFolder.Name= " & objTopFolder.Name
    Next i
    
    'This loop finds the TopFolder you specified.
    For i = 1 To objInfoStores.Count
        Set objInfoStore = objInfoStores(i)
        Set objTopFolder = objInfoStore.RootFolder
        If objTopFolder.Name = strTargetTopFolder Then 'Found PST
            'Because you can have more than one PST in a profile,
            'you may want to put another check here to make sure you have
            'the correct PST. This check would need to specify a string
            'that is the name of the PST you are looking for.
            'It would look something like this:
            'If objInfoStore.Name = "MyPST" Then 'Found own PST
            '   Exit For
            'End If
        End If
    Next i
    Set objPSTFolders = objTopFolder.Folders
    For i = 1 To objPSTFolders.Count
        MsgBox objPSTFolders.Item(i).Name
        If objPSTFolders.Item(i).Name = strSearchName Then
            Exit For
        End If
    Next i
    Set objFindTargetFolder = objPSTFolders.Item(i)
    Set objTopFolder = Nothing
    Set objPSTFolders = Nothing
    Set objInfoStores = Nothing
    Set objInfoStore = Nothing
End Function


