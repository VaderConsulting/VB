VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCatMgrTest 
   Caption         =   "Categories Manager Sample"
   ClientHeight    =   4290
   ClientLeft      =   2130
   ClientTop       =   2835
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstImplemented 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      IntegralHeight  =   0   'False
      Left            =   1980
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2190
      Width           =   1890
   End
   Begin ComctlLib.ListView lvwClasses 
      Height          =   2070
      Left            =   45
      TabIndex        =   3
      Top             =   2190
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   3651
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ProgID"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "CLSID"
         Object.Width           =   6350
      EndProperty
   End
   Begin ComctlLib.ListView lvwCats 
      Height          =   1650
      Left            =   0
      TabIndex        =   1
      Top             =   255
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   2910
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Category"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "CATID"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.Label lblImp 
      AutoSize        =   -1  'True
      Caption         =   "Implemented Categories:"
      Height          =   195
      Left            =   1980
      TabIndex        =   5
      Top             =   1950
      Width           =   1740
   End
   Begin VB.Label lblClasses 
      AutoSize        =   -1  'True
      Caption         =   "Classes:"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1965
      Width           =   585
   End
   Begin VB.Label lblCats 
      AutoSize        =   -1  'True
      Caption         =   "&Categories:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   795
   End
End
Attribute VB_Name = "frmCatMgrTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' Component Categories Manager Sample
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 01/23/2000
'
'*********************************************************************************************
Option Explicit

Dim CCM As StdComponentCategoriesMgr

Const HKEY_CLASSES_ROOT = &H80000000
Const KEY_QUERY_VALUE = 1
Const REG_SZ = 1

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal Str As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal L As Long)

Const S_OK = 0
Private Sub EnumCategories()
Dim oEnumGUID As IEnumCATEGORYINFO, CI As CATEGORYINFO
Dim GUID As String * 39

   With lvwCats
   
      ' Clear the listview
      .ListItems.Clear
            
      ' Get the categories enumerator
      ' object
      Set oEnumGUID = CCM.EnumCategories(0)
      
      ' Enumerate all categories
      Do While oEnumGUID.Next(1, CI) = S_OK
      
         With .ListItems.Add(, , CI.szDescription)
            
            ' Convert olelib.UUID struct to string
            StringFromGUID2 CI.CATID, GUID, Len(GUID)
            
            .SubItems(1) = GUID
            
         End With
      
      Loop
      
      .ListItems(1).Selected = True
           
   End With
   
End Sub

Private Sub EnumImplemented(ByVal sCLSID As String)
Dim oEnumGUID As IEnumGUID, CATID As olelib.UUID
Dim CLSID As olelib.UUID

   On Error Resume Next

   With lstImplemented
   
      ' Clear the list
      .Clear
            
      CLSIDFromString sCLSID, CLSID
      
      Set oEnumGUID = CCM.EnumImplCategoriesOfClass(CLSID)
      
      If Not oEnumGUID Is Nothing Then
         ' Enumerate all categories
         Do While oEnumGUID.Next(1, CATID) = S_OK
            
            .AddItem SysAllocString(CCM.GetCategoryDesc(CATID, 0))
         
         Loop
      End If
      
      Set oEnumGUID = CCM.EnumReqCategoriesOfClass(CLSID)
      
      If Not oEnumGUID Is Nothing Then
         
         ' Enumerate all categories
         Do While oEnumGUID.Next(1, CATID) = S_OK
            
            .AddItem SysAllocString(CCM.GetCategoryDesc(CATID, 0))
         
         Loop
         
      End If
      
   End With

End Sub

Private Sub EnumObjects(ByVal sCATID As String)
Dim oEnumGUID As IEnumGUID, CATID As olelib.UUID, CLSID As olelib.UUID, sGUID As String * 38
Dim hKey As Long, lKeylen As Long, sKeyValue As String

   With lvwClasses
   
      ' Clear the list
      .ListItems.Clear
      
      ' Convert string to olelib.UUID struct
      CLSIDFromString sCATID, CATID
      
      ' Get enumerator
      Set oEnumGUID = CCM.EnumClassesOfCategories(1, CATID, 0, CATID)
      
      Do While oEnumGUID.Next(1, CLSID) = S_OK
         
         ' Convert olelib.UUID struct to string
         StringFromGUID2 CLSID, sGUID, Len(sGUID) + 1
         
         With .ListItems.Add()
            
            .SubItems(2) = sGUID
            
            ' Open the registry key
            ' to get the class name
            If RegOpenKeyEx(HKEY_CLASSES_ROOT, "CLSID\" & sGUID, 0&, KEY_QUERY_VALUE, hKey) = 0 Then
            
               ' Get name length
               RegQueryValueEx hKey, vbNullString, 0&, REG_SZ, ByVal 0&, lKeylen
               
               ' Get name
               sKeyValue = String$(lKeylen, 0)
               RegQueryValueEx hKey, vbNullString, 0&, REG_SZ, ByVal sKeyValue, lKeylen
                           
               ' Close the key
               RegCloseKey hKey
               
               .Text = sKeyValue

            End If
         
            ' Get ProgID
            If RegOpenKeyEx(HKEY_CLASSES_ROOT, "CLSID\" & sGUID & "\ProgID", 0&, KEY_QUERY_VALUE, hKey) = 0 Then
            
               RegQueryValueEx hKey, vbNullString, 0&, REG_SZ, ByVal 0&, lKeylen
               
               sKeyValue = String$(lKeylen, 0)
               RegQueryValueEx hKey, vbNullString, 0&, REG_SZ, ByVal sKeyValue, lKeylen
               
               RegCloseKey hKey
               
               .SubItems(1) = sKeyValue

            End If
         
         End With
         
      Loop
   
      '.ListItems(1).Selected = True
   
   End With
   
End Sub


Private Sub Form_Load()

   Set CCM = New StdComponentCategoriesMgr
   
   EnumCategories
   
End Sub


Private Sub Form_Resize()
Dim Middle As Single, Center As Single

   Middle = (ScaleHeight - lblCats.Height - lblClasses.Height - 3) / 2
   Center = (ScaleWidth - 2) / 2
   
   lblCats.Move 0, 0
   lvwCats.Move 0, lblCats.Height + 1, ScaleWidth, Middle
   
   lblClasses.Move 0, lvwCats.Height + lvwCats.Top + 1
   lvwClasses.Move 0, lblClasses.Height + lblClasses.Top + 1, Center - 1, Middle
   
   lblImp.Move Center + 1, lblClasses.Top
   lstImplemented.Move Center + 1, lvwClasses.Top, Center, lvwClasses.Height
   
End Sub


Private Sub lvwCats_ItemClick(ByVal Item As ComctlLib.ListItem)
   
   EnumObjects Item.SubItems(1)
   DoEvents
End Sub


Private Sub lvwClasses_ItemClick(ByVal Item As ComctlLib.ListItem)

   EnumImplemented Item.SubItems(2)
   
End Sub


