VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asset Control"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tabRecords 
      Height          =   5775
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Existing data"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblField(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblField(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblField(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblField(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblField(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblField(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblAssetIDCaption"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblAssetID"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblField(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblExtraInfo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblField(14)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblProductID"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblProductIDCaption"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblField(16)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblExtraData"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtData(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtData(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtData(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtData(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtData(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtData(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdUpdate"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtData(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtData(14)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdBrowseDescription"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtData(16)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdViewExtraData"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdAddExtraData"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdAddAssociation"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdViewAssociated"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "New data"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBrowseSupplier"
      Tab(1).Control(1)=   "cmdBrowseManufacturer"
      Tab(1).Control(2)=   "txtData(15)"
      Tab(1).Control(3)=   "txtData(13)"
      Tab(1).Control(4)=   "txtData(12)"
      Tab(1).Control(5)=   "optDataType(1)"
      Tab(1).Control(6)=   "optDataType(0)"
      Tab(1).Control(7)=   "cmdAdd"
      Tab(1).Control(8)=   "txtData(11)"
      Tab(1).Control(9)=   "txtData(10)"
      Tab(1).Control(10)=   "txtData(9)"
      Tab(1).Control(11)=   "txtData(8)"
      Tab(1).Control(12)=   "txtData(7)"
      Tab(1).Control(13)=   "lblNewAssetProductID"
      Tab(1).Control(14)=   "lblNewAssetProductIDCaption"
      Tab(1).Control(15)=   "lblField(15)"
      Tab(1).Control(16)=   "lblField(13)"
      Tab(1).Control(17)=   "lblField(12)"
      Tab(1).Control(18)=   "lblField(11)"
      Tab(1).Control(19)=   "lblField(10)"
      Tab(1).Control(20)=   "lblField(9)"
      Tab(1).Control(21)=   "lblField(8)"
      Tab(1).Control(22)=   "lblField(7)"
      Tab(1).ControlCount=   23
      Begin VB.CommandButton cmdViewAssociated 
         Caption         =   "View assoc."
         Height          =   375
         Left            =   3480
         TabIndex        =   57
         Top             =   4800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdAddAssociation 
         Caption         =   "Add assoc."
         Height          =   375
         Left            =   4560
         TabIndex        =   56
         Top             =   4800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdAddExtraData 
         Caption         =   "Add data"
         Height          =   375
         Left            =   4560
         TabIndex        =   55
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdViewExtraData 
         Caption         =   "View"
         Height          =   375
         Left            =   3480
         TabIndex        =   54
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   52
         Top             =   3720
         Width           =   3735
      End
      Begin VB.CommandButton cmdBrowseDescription 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   5170
         TabIndex        =   48
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseSupplier 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -69830
         TabIndex        =   45
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseManufacturer 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   -69830
         TabIndex        =   44
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   15
         Left            =   -73560
         TabIndex        =   15
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   14
         Left            =   1440
         TabIndex        =   4
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   13
         Left            =   -73560
         TabIndex        =   20
         Top             =   3360
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   12
         Left            =   -73560
         TabIndex        =   19
         Top             =   3000
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   9
         Top             =   3360
         Width           =   3735
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   5280
         Width           =   975
      End
      Begin VB.OptionButton optDataType 
         Caption         =   "Product"
         Height          =   255
         Index           =   1
         Left            =   -71640
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optDataType 
         Caption         =   "Asset"
         Height          =   255
         Index           =   0
         Left            =   -73080
         TabIndex        =   11
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70440
         TabIndex        =   21
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   11
         Left            =   -73560
         TabIndex        =   18
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   10
         Left            =   -73560
         TabIndex        =   17
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   9
         Left            =   -73560
         TabIndex        =   16
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   8
         Left            =   -73560
         TabIndex        =   14
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   7
         Left            =   -73560
         TabIndex        =   13
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   5
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   7
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   8
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label lblExtraData 
         Caption         =   "Extra data is available.  -->"
         Height          =   255
         Left            =   1440
         TabIndex        =   53
         Top             =   4440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblField 
         Caption         =   "NSN"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   51
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblNewAssetProductID 
         Height          =   255
         Left            =   -73560
         TabIndex        =   50
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblNewAssetProductIDCaption 
         Caption         =   "Product ID"
         Height          =   255
         Left            =   -74880
         TabIndex        =   49
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblProductIDCaption 
         Caption         =   "Product ID"
         Height          =   255
         Left            =   2880
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblProductID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   46
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblField 
         Caption         =   "Product Key"
         Height          =   255
         Index           =   15
         Left            =   -74880
         TabIndex        =   43
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblField 
         Caption         =   "Product Key"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblExtraInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   41
         Top             =   4080
         Width           =   3735
      End
      Begin VB.Label lblField 
         Caption         =   "Supp. Code"
         Height          =   255
         Index           =   13
         Left            =   -74880
         TabIndex        =   40
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblField 
         Caption         =   "Supp. Code"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   39
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label lblField 
         Caption         =   "Man. Code"
         Height          =   255
         Index           =   12
         Left            =   -74880
         TabIndex        =   38
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblAssetID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   37
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblAssetIDCaption 
         Caption         =   "Asset ID"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblField 
         Caption         =   "Location"
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   35
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblField 
         Caption         =   "Date Added"
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   34
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblField 
         Caption         =   "Description"
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   33
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblField 
         Caption         =   "Asset Number"
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   32
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblField 
         Caption         =   "Serial Number"
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   31
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblField 
         Caption         =   "Serial Number"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblField 
         Caption         =   "Asset Number"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblField 
         Caption         =   "Description"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblField 
         Caption         =   "Date Added"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblField 
         Caption         =   "Location"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblField 
         Caption         =   "Man. Code"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdFetch 
      Caption         =   "Fetch"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtBarcode 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblBarcode 
      Caption         =   "Barcode:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6480
      Width           =   5655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenLast 
         Caption         =   "1"
         Index           =   0
      End
      Begin VB.Menu mnuOpenLast 
         Caption         =   "2"
         Index           =   1
      End
      Begin VB.Menu mnuOpenLast 
         Caption         =   "3"
         Index           =   2
      End
      Begin VB.Menu mnuOpenLast 
         Caption         =   "4"
         Index           =   3
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_RETURNFSANCESTORS = 8
Private Const BIF_EDITBOX = 16
Private Const BIF_NEWDIALOGSTYLE = 64
Private Const BIF_BROWSEINCLUDEFILES = 16384

Private Const MAX_PATH = 260

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Sub cmdBrowseDescription_Click()
    BrowseProducts 1
End Sub

Private Sub cmdBrowseManufacturer_Click()
    If optDataType(0).Value = True Then ' Asset
        BrowseProducts 2
        If txtData(9).Text <> "" Then cmdAdd.Enabled = True
    ElseIf optDataType(1).Value = True Then ' Product
        BrowseManufacturer
    End If
End Sub

Private Sub cmdBrowseSupplier_Click()
    BrowseSupplier
End Sub

Private Sub BrowseManufacturer()
    Dim strManufacturer As String
    
    frmBrowse.Caption = "Browse Manufacturers"
    frmBrowse.strQuery = "SELECT ID, Description, Manufacturer, Supplier, ManufacturerCode, SupplierCode FROM tblProducts ORDER BY Description" ' SELECT ID, Manufacturer FROM tblProducts ORDER BY Manufacturer"
    frmBrowse.GetData
    frmBrowse.iColToReturn = 2
    frmBrowse.Show vbModal
    strManufacturer = frmBrowse.strSelectedItem
    Set frmBrowse = Nothing
    txtData(9).Text = strManufacturer
    txtData(9).SetFocus
End Sub

Private Sub BrowseProducts(iCaller As Integer)
    Dim strDescription As String
    Dim strProductID As String
    
    ' iCaller = 1 is existing data,
    ' iCaller = 2 is new data
    
    frmBrowse.Caption = "Browse Products"
    frmBrowse.iColToReturn = 1
    frmBrowse.strQuery = "SELECT ID, Description, Manufacturer, Supplier, ManufacturerCode, SupplierCode FROM tblProducts ORDER BY Description"
    frmBrowse.GetData
    frmBrowse.Show vbModal
    strDescription = frmBrowse.strSelectedItem
    strProductID = frmBrowse.strProductID
    
    Select Case iCaller
        Case 1 ' existing
            If strDescription <> "" Then
                txtData(2).Text = strDescription
            End If
            txtData(2).SetFocus
        Case 2 ' new
            If strDescription <> "" Then
                txtData(9).Text = strDescription
            End If
            txtData(9).SetFocus
            If strProductID <> "" Then
                lblNewAssetProductID.Caption = strProductID
            End If
    End Select
    
    Set frmBrowse = Nothing
End Sub

Private Sub BrowseSupplier()
    Dim strSupplier As String
    
    frmBrowse.Caption = "Browse Suppliers"
    frmBrowse.strQuery = "SELECT ID, Description, Manufacturer, Supplier, ManufacturerCode, SupplierCode FROM tblProducts ORDER BY Description" ' SELECT ID, Supplier FROM tblProducts ORDER BY Supplier"
    frmBrowse.GetData
    frmBrowse.iColToReturn = 3
    frmBrowse.Show vbModal
    strSupplier = frmBrowse.strSelectedItem
    Set frmBrowse = Nothing
    txtData(10).Text = strSupplier
    txtData(10).SetFocus
End Sub

Private Sub cmdAddExtraData_Click()
    ExtraData Add
    HighlightBarcode
End Sub

Private Sub cmdViewExtraData_Click()
    ExtraData Update
    HighlightBarcode
End Sub

Private Sub ExtraData(iMethod As Method)
    Dim strSQL As String
    Dim strTable As String
    Dim Source As IDSource
    
    If Me.lblProductID.Visible = True Then
        frmExtraData.strID = Me.lblAssetID.Caption
        frmExtraData.strSource = Asset
        Source = Asset
    Else
        frmExtraData.strID = Me.lblAssetID.Caption
        frmExtraData.strSource = Product
        Source = Product
    End If
        
    Select Case iMethod
        Case Method.Add
            frmExtraData.iMethod = Add
            frmExtraData.Show vbModal
            
            If frmExtraData.strRecordID <> "" Then
                ' Update this record with the record number of the added extra data
                
                Select Case Source
                    Case IDSource.Asset
                        strTable = "tblAssets"
                    Case IDSource.Product
                        strTable = "tblProducts"
                End Select
                
                strSQL = ""
                strSQL = strSQL & "UPDATE " & strTable & " "
                strSQL = strSQL & "SET ExtraDataID = " & frmExtraData.strRecordID & " "
                strSQL = strSQL & "WHERE ID = " & frmExtraData.strID
                
                oConn.Execute strSQL
                cmdViewExtraData.Visible = True
                lblExtraData.Visible = True
                cmdAddExtraData.Visible = False
            End If
        Case Method.Update
            frmExtraData.iMethod = Update
            frmExtraData.Show vbModal
    End Select
    
    Set frmExtraData = Nothing
End Sub

Private Sub Form_Load()
    Dim LastFile(3) As String
    Dim i As Integer
    
    For i = 0 To 3
        LastFile(i) = GetSetting("Asset Control", "Previous", CStr(i), "")
        If LastFile(i) <> "" Then
            mnuOpenLast(i).Caption = LastFile(i)
            mnuOpenLast(i).Visible = True
        Else
            mnuOpenLast(i).Visible = False
        End If
    Next
    
    Set oDatabase = CreateObject("ADODB.Recordset")
    Set oConn = CreateObject("ADODB.Connection")
    
    cmdFetch.Enabled = False
    
    ' Open last database
    If LastFile(0) <> "" Then
        OpenDatabase LastFile(0)
        mnuOpen.Enabled = False
        mnuClose.Enabled = True
        txtBarcode.Enabled = True
    Else
        mnuOpen.Enabled = True
        mnuClose.Enabled = False
        txtBarcode.Enabled = False
    End If
    
    txtData(10).Text = FormatDateTime(Date, vbShortDate)
End Sub

Private Sub cmdAdd_Click()
    AddData
    cmdAdd.Enabled = False
End Sub

Private Sub AddData()
    Dim strSQL As String
    
    If optDataType(0).Value = True Then ' Asset
        strSQL = ""
        strSQL = strSQL & "INSERT INTO tblAssets "
        strSQL = strSQL & "(Serial, Asset, ProductKey, AddDate, LastDate, ProductID, ScanBy, Location) VALUES "
        strSQL = strSQL & "(""" & txtData(7).Text & """,""" & txtData(8).Text & """,""" & txtData(14).Text & """,""" & txtData(10).Text & """,""" & txtData(10).Text & """," & lblNewAssetProductID.Caption & ",""" & Environ$("USERNAME") & """,""" & txtData(11).Text & """)"
        lblInfo = "Asset Record added."
    Else ' Product
        strSQL = ""
        strSQL = strSQL & "INSERT INTO tblProducts "
        strSQL = strSQL & "(Description, Manufacturer, Supplier, ManufacturerCode, SupplierCode, NSN) VALUES "
        strSQL = strSQL & "(""" & txtData(8).Text & """,""" & txtData(9).Text & """,""" & txtData(10).Text & """,""" & txtData(11).Text & """,""" & txtData(12).Text & """,""" & txtData(13).Text & """)"
        lblInfo = "Product Record added."
    End If
    
    oConn.Execute strSQL
    
    txtBarcode.Text = ""
    HighlightBarcode
End Sub

Private Sub cmdFetch_Click()
    FetchData
End Sub

Private Sub cmdUpdate_Click()
    UpdateData
End Sub

Private Sub UpdateData()
    Dim strSQL As String
    
    If lblProductID.Visible = True Then ' Asset, not product
    
        ' Update Asset Data
        strSQL = ""
        strSQL = strSQL & "UPDATE tblAssets SET "
        strSQL = strSQL & "Serial = """ & txtData(0) & """, "
        strSQL = strSQL & "Asset = """ & txtData(1) & """, "
        strSQL = strSQL & "ProductKey = """ & txtData(14) & """, "
        strSQL = strSQL & "LastDate = """ & Date & """, "
        strSQL = strSQL & "Location = """ & txtData(4) & """ "
        strSQL = strSQL & "WHERE ID = " & lblAssetID.Caption
        
        oConn.Execute strSQL
        lblInfo = "Asset Record updated."
    Else
        ' Update Product Data (Description)
        If lblAssetID.Caption <> "" Then
            strSQL = ""
            strSQL = strSQL & "UPDATE tblProducts SET "
            strSQL = strSQL & "Description = """ & txtData(2) & """, "
            strSQL = strSQL & "ManufacturerCode = """ & txtData(5) & """, "
            strSQL = strSQL & "SupplierCode = """ & txtData(6) & """, "
            strSQL = strSQL & "NSN = """ & txtData(16) & """"
            strSQL = strSQL & "WHERE ID = " & lblAssetID.Caption
            
            oConn.Execute strSQL
            lblInfo = "Product Record updated."
        Else
            lblInfo = "No Product ID, unable to save record."
        End If
    End If
    
    txtBarcode.Text = ""
    HighlightBarcode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    ExitApp
End Sub

Private Sub mnuAbout_Click()
    MsgBox "For assisting ADI in Asset Control" & vbCrLf & vbCrLf & "(c) 2005 D. Robinson", vbOKOnly + vbInformation, "Asset Control"
End Sub

Private Sub mnuClose_Click()
    On Error GoTo Hell
    Select Case oDatabase.State
        Case ADODB.adStateOpen
            lblInfo = "Closing database..."
            frmMain.Refresh
            oDatabase.Close
        Case ADODB.adStateClosed
            ' Do nothing
        Case Else
            Do Until oDatabase.State = ADODB.adStateOpen
                lblInfo = "Waiting to close database..."
                frmMain.Refresh
                DoEvents
            Loop
            oDatabase.Close
    End Select
    lblInfo = "Database closed."
    
    mnuOpen.Enabled = True
    mnuClose.Enabled = False
    
    cmdBrowseManufacturer.Enabled = False
    cmdBrowseSupplier.Enabled = False
    cmdBrowseDescription.Enabled = False
    
    frmMain.Refresh
    Exit Sub
Hell:
    lblInfo = "Error " & Err.Number & " attempting to close database."
    frmMain.Refresh
End Sub

Private Sub mnuExit_Click()
    ExitApp
End Sub

Public Sub ExitApp()
    ' Close database objects,
    If oDatabase.State = ADODB.adStateOpen Then
        oDatabase.Close
    End If
    
    If oConn.State = ADODB.adStateOpen Then
        oConn.Close
    End If
    
    ' Remove references
    
    Set oDatabase = Nothing
    Set oConn = Nothing
    
    ' Close application
    End
End Sub

Private Sub mnuOpen_Click()
    Dim strDatabaseName As String
    Dim OpenFile As OPENFILENAME, SelectedFile As String
    Dim lReturn As Long
    Dim sFilter As String
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hWndOwner = frmMain.hWnd
    OpenFile.hInstance = App.hInstance
    sFilter = "All Files (*.*)" & Chr(0) & "*.*" & Chr(0) & "Databases (*.mdb)" & Chr(0) & "*.MDB" & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 2
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = App.Path
    OpenFile.lpstrTitle = "Select database"
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        'The User pressed the Cancel Button
    Else
        SelectedFile = Trim(OpenFile.lpstrFile)
        If SelectedFile <> "" Then
            strDatabaseName = SelectedFile
        Else
            strDatabaseName = ""
        End If
        
        ' Remove CHR(0)'s
        strDatabaseName = Replace(strDatabaseName, Chr(0), "")
        
        If strDatabaseName <> "" Then
            OpenDatabase strDatabaseName
        End If
        
    End If

End Sub

Private Sub mnuOpenLast_Click(Index As Integer)
    OpenDatabase mnuOpenLast(Index).Caption
End Sub

Private Sub optDataType_Click(Index As Integer)
    ClearDataEntryFields
    Select Case Index
        Case 0 ' Asset
            lblField(7) = "Serial Number"
            lblField(7).Visible = True
            txtData(7).Visible = True
            lblField(8) = "Asset Number"
            lblField(9) = "Description"
            lblField(10) = "Date Added"
            txtData(10).Text = FormatDateTime(Date, vbShortDate)
            lblField(11) = "Location"
            lblField(12) = ""
            lblField(12).Visible = False
            txtData(12).Visible = False
            lblField(13) = ""
            lblField(13).Visible = False
            txtData(13).Visible = False
            lblField(15).Visible = True
            txtData(15).Visible = True
            cmdBrowseSupplier.Visible = False
            lblNewAssetProductIDCaption.Visible = True
        Case 1 ' Product
            lblField(7) = "Serial Number"
            lblField(7).Visible = False
            txtData(7).Visible = False
            lblField(8) = "Description"
            lblField(9) = "Manufacturer"
            lblField(10) = "Supplier"
            txtData(10).Text = ""
            lblField(11) = "Man. Code"
            lblField(12) = "Supp. Code"
            lblField(12).Visible = True
            txtData(12).Visible = True
            lblField(13) = "NSN"
            lblField(13).Visible = True
            txtData(13).Visible = True
            lblField(15).Visible = False
            txtData(15).Visible = False
            'cmdBrowseManufacturer.Visible = True
            cmdBrowseSupplier.Visible = True
            lblNewAssetProductIDCaption.Visible = False
    End Select
End Sub

Private Sub tabRecords_Click(PreviousTab As Integer)
    If txtBarcode.Enabled Then
        txtBarcode.SetFocus
        HighlightBarcode
    End If
End Sub

Private Sub tabRecords_DblClick()
    If txtBarcode.Enabled Then
        txtBarcode.SetFocus
        HighlightBarcode
    End If
End Sub

Private Sub txtBarcode_Change()
    If txtBarcode.Text <> "" Then
        cmdFetch.Enabled = True
    Else
        cmdFetch.Enabled = False
    End If
End Sub

Private Sub txtBarcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        FetchData
        HighlightBarcode
    End If
End Sub

Public Sub OpenDatabase(strFilename As String)
    DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & strFilename & ";Mode=ReadWrite"
    If oConn.State = adStateOpen Then oConn.Close
    oConn.Open DSN
    oConn.CursorLocation = adUseClient
    
    mnuOpen.Enabled = False
    mnuClose.Enabled = True
    txtBarcode.Enabled = True
    lblInfo = "Database opened."
    If Not Me.txtBarcode Is Nothing Then
        HighlightBarcode
    End If
    
    cmdBrowseManufacturer.Enabled = True
    cmdBrowseSupplier.Enabled = True
    cmdBrowseDescription.Enabled = True
    
    SaveSetting "Asset Control", "Previous", 0, strFilename
End Sub

Private Sub FetchData()
    Dim strSQL As String
    Dim iRecordCount As Integer
    Dim strLabel As String
    Dim iRecordSource As Integer
    
    ClearDataEntryFields
    
    ' See if we can find this barcode in the Assets table
    iRecordSource = 1
    strSQL = "SELECT * FROM tblAssets WHERE Serial like '" & txtBarcode.Text & "' OR Asset LIKE '" & txtBarcode.Text & "'"
    iRecordCount = GetRecords(strSQL)
    
    If iRecordCount = 0 Then
        ' See if we can find this barcode in Products table
        iRecordSource = 2
        strSQL = "SELECT * FROM tblProducts WHERE ManufacturerCode like '" & txtBarcode.Text & "' OR "
        strSQL = strSQL & "SupplierCode LIKE '" & txtBarcode.Text & "'"
        iRecordCount = GetRecords(strSQL)
    End If
    
    If iRecordCount = 0 Then
        ' Record not found.
        tabRecords.Tab = 1
        frmNewData.Show vbModal
        HighlightBarcode
        txtData(8).SetFocus
    ElseIf iRecordCount = 1 Then ' Found a single record
        tabRecords.Tab = 0
        Select Case iRecordSource
            Case 1 ' tblAssets
                strSQL = ""
                strSQL = strSQL & "SELECT tblAssets.ID, tblAssets.Serial, tblAssets.Asset,tblAssets.ProductKey, "
                strSQL = strSQL & "tblAssets.AddDate, tblAssets.LastDate, "
                strSQL = strSQL & "tblAssets.ScanBy, tblAssets.Location, tblAssets.ProductID, "
                strSQL = strSQL & "tblProducts.Description, tblProducts.Manufacturer, tblProducts.Supplier, "
                strSQL = strSQL & "tblProducts.ManufacturerCode, tblProducts.SupplierCode, tblProducts.NSN, tblAssets.ExtraDataID "
                strSQL = strSQL & "FROM tblProducts RIGHT JOIN tblAssets ON tblProducts.ID = tblAssets.ProductID "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "tblAssets.Serial LIKE '" & txtBarcode.Text & "' "
                strSQL = strSQL & "OR tblAssets.Asset LIKE '" & txtBarcode.Text & "'"
                
                If oDatabase.State = adStateOpen Then oDatabase.Close
                oDatabase.Open strSQL, oConn, adOpenDynamic, adLockOptimistic
                
                lblAssetIDCaption.Caption = "Asset ID"
                lblAssetID.Caption = oDatabase.Fields("ID").Value
                txtData(0).Text = oDatabase.Fields("Serial").Value & ""
                txtData(0).Visible = True
                lblField(0).Visible = True
                txtData(1).Text = oDatabase.Fields("Asset").Value & ""
                lblField(1).Visible = True
                txtData(1).Visible = True
                txtData(2).Text = oDatabase.Fields("Description").Value & ""
                txtData(3).Text = oDatabase.Fields("AddDate").Value & ""
                lblField(3).Visible = True
                txtData(3).Visible = True
                txtData(4).Text = oDatabase.Fields("Location").Value & ""
                lblField(4).Visible = True
                txtData(4).Visible = True
                txtData(5).Text = oDatabase.Fields("ManufacturerCode").Value & ""
                txtData(6).Text = oDatabase.Fields("SupplierCode").Value & ""
                txtData(14).Text = oDatabase.Fields("ProductKey").Value & ""
                lblField(14).Visible = True
                txtData(14).Visible = True
                txtData(16).Text = oDatabase.Fields("NSN").Value & ""
                
                lblProductID.Caption = oDatabase.Fields("ProductID").Value & ""
                lblProductIDCaption.Visible = True
                lblProductID.Visible = True
                lblExtraInfo.Caption = "Asset data shown"
                
                ' Show the user which field was used to find this record
                If UCase(txtData(0).Text) = UCase(txtBarcode.Text) Then
                    lblField(0).FontBold = True  ' Serial
                    lblField(1).FontBold = False ' Asset
                    lblField(5).FontBold = False ' Manufacturer Code
                    lblField(6).FontBold = False ' Supplier Code
                Else
                    lblField(0).FontBold = False ' Serial
                    lblField(1).FontBold = True  ' Asset
                    lblField(5).FontBold = False ' Manufacturer Code
                    lblField(6).FontBold = False ' Supplier Code
                End If
                HighlightBarcode
            Case 2 ' tblProducts
                lblAssetIDCaption.Caption = "Product ID"
                lblAssetID.Caption = oDatabase.Fields("ID").Value
                txtData(0).Visible = False
                lblField(0).Visible = False
                txtData(1).Visible = False
                lblField(1).Visible = False
                txtData(2).Text = oDatabase.Fields("Description").Value & ""
                txtData(3).Visible = False
                lblField(3).Visible = False
                txtData(4).Visible = False
                lblField(4).Visible = False
                txtData(5).Text = oDatabase.Fields("ManufacturerCode").Value & ""
                txtData(6).Text = oDatabase.Fields("SupplierCode").Value & ""
                txtData(14).Visible = False
                lblField(14).Visible = False
                txtData(16).Text = oDatabase.Fields("NSN").Value & ""
                
                lblProductIDCaption.Visible = False
                lblProductID.Visible = False
                lblExtraInfo.Caption = "Product data shown"
                
                ' Show the user which field was used to find this record
                If txtData(5).Text = txtBarcode.Text Then
                    lblField(0).FontBold = False ' Serial
                    lblField(1).FontBold = False ' Asset
                    lblField(5).FontBold = True  ' Manufacturer Code
                    lblField(6).FontBold = False ' Supplier Code
                Else
                    lblField(0).FontBold = False ' Serial
                    lblField(1).FontBold = False ' Asset
                    lblField(5).FontBold = False ' Manufacturer Code
                    lblField(6).FontBold = True  ' Supplier Code
                End If
                
                HighlightBarcode
        End Select
    Else
        MsgBox "More than 1 record returned.  This is currently not supported.", vbInformation, "Unsupported"
        Exit Sub
    End If
    
    If oDatabase.EOF Then
        tabRecords.Tab = 0
        HighlightBarcode
        Exit Sub
    End If
    
    If oDatabase.Fields("ExtraDataID") <> "0" Then
        Me.lblExtraData.Visible = True
        Me.cmdViewExtraData.Visible = True
        Me.cmdAddExtraData.Visible = False
    Else
        Me.lblExtraData.Visible = False
        Me.cmdViewExtraData.Visible = False
        Me.cmdAddExtraData.Visible = True
    End If
    
    oDatabase.Close
End Sub

Private Sub ClearDataEntryFields()
    txtData(0).Text = ""
    txtData(1).Text = ""
    txtData(2).Text = ""
    txtData(3).Text = ""
    txtData(4).Text = ""
    txtData(5).Text = ""
    txtData(6).Text = ""
    txtData(7).Text = ""
    txtData(8).Text = ""
    txtData(9).Text = ""
    txtData(10).Text = ""
    txtData(11).Text = ""
    txtData(12).Text = ""
    txtData(13).Text = ""
    txtData(14).Text = ""
    txtData(15).Text = ""
    txtData(16).Text = ""
    lblAssetID.Caption = ""
End Sub

Private Sub HighlightBarcode()
    On Error Resume Next
    txtBarcode.SetFocus
    txtBarcode.SelStart = 0
    txtBarcode.SelLength = Len(txtBarcode.Text)
    On Error GoTo 0
End Sub

Private Sub txtData_Change(Index As Integer)
    cmdUpdate.Enabled = True
End Sub

Private Sub txtData_GotFocus(Index As Integer)
    txtData(Index).SelStart = 0
    txtData(Index).SelLength = Len(txtData(Index).Text)
End Sub
