VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNewData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Data"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmNewData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin TabDlg.SSTab tabNewData 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detail type"
      TabPicture(0)   =   "frmNewData.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "optDetailType(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "optDetailType(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Data Type"
      TabPicture(1)   =   "frmNewData.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optDataType(2)"
      Tab(1).Control(1)=   "optDataType(1)"
      Tab(1).Control(2)=   "optDataType(0)"
      Tab(1).ControlCount=   3
      Begin VB.OptionButton optDataType 
         Caption         =   "Product Key"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   7
         Top             =   1320
         Width           =   4095
      End
      Begin VB.OptionButton optDataType 
         Caption         =   "Serial"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optDataType 
         Caption         =   "Barcode"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optDetailType 
         Caption         =   "Product"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optDetailType 
         Caption         =   "Asset"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Label lblNewData 
      Alignment       =   2  'Center
      Caption         =   "You have scanned new data.  Please complete the following information."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmNewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum DetailType
    Asset = 0
    Product = 1
End Enum

Private Enum DataType
    Barcode = 0
    Serial = 1
    ProductKey = 2
    ManufacturerCode = 3
    ProductCode = 4
    Other = 5
End Enum

Private Enum PageChangeType
    TabClick = 1
    ButtonClick = 2
End Enum

Dim iDetailType As DetailType
Dim iDataType As DataType

Private Sub optDataType_Click(Index As Integer)
    Select Case iDetailType
        Case DetailType.Asset
            Select Case Index
                Case 0
                    iDataType = Barcode
                    frmMain.txtData(7).Text = frmMain.txtBarcode.Text
                    frmMain.txtData(8).Text = ""
                    frmMain.txtData(15).Text = ""
                    frmMain.txtData(11).Text = ""
                    frmMain.txtData(12).Text = ""
                    frmMain.cmdAdd.Enabled = True
                Case 1
                    iDataType = Serial
                    frmMain.txtData(7).Text = ""
                    frmMain.txtData(8).Text = frmMain.txtBarcode.Text
                    frmMain.txtData(15).Text = ""
                    frmMain.txtData(11).Text = ""
                    frmMain.txtData(12).Text = ""
                    frmMain.cmdAdd.Enabled = True
                Case 2
                    iDataType = ProductKey
                    frmMain.txtData(7).Text = ""
                    frmMain.txtData(8).Text = ""
                    frmMain.txtData(15).Text = frmMain.txtBarcode.Text
                    frmMain.txtData(11).Text = ""
                    frmMain.txtData(12).Text = ""
                    frmMain.cmdAdd.Enabled = True
            End Select
        Case DetailType.Product
            Select Case Index
                Case 0
                    iDataType = ManufacturerCode
                    frmMain.txtData(7).Text = ""
                    frmMain.txtData(8).Text = ""
                    frmMain.txtData(15).Text = ""
                    frmMain.txtData(11).Text = frmMain.txtBarcode.Text
                    frmMain.txtData(12).Text = ""
                    frmMain.cmdAdd.Enabled = True
                Case 1
                    iDataType = ProductCode
                    frmMain.txtData(7).Text = ""
                    frmMain.txtData(8).Text = ""
                    frmMain.txtData(15).Text = ""
                    frmMain.txtData(11).Text = ""
                    frmMain.txtData(12).Text = frmMain.txtBarcode.Text
                    frmMain.cmdAdd.Enabled = True
                Case 2
                    iDataType = Other
                    frmMain.txtData(7).Text = ""
                    frmMain.txtData(8).Text = ""
                    frmMain.txtData(15).Text = ""
                    frmMain.txtData(11).Text = ""
                    frmMain.txtData(12).Text = ""
                    frmMain.cmdAdd.Enabled = False
            End Select
    End Select
End Sub

Private Sub optDetailType_Click(Index As Integer)
    Select Case Index
        Case 0
            iDetailType = Asset
            frmMain.optDataType(0).Value = True
        Case 1
            iDetailType = Product
            frmMain.optDataType(1).Value = True
    End Select
End Sub

Private Sub DoCommandButton(iType As PageChangeType)
    If iType = ButtonClick Then
        If cmdNext.Caption = "Finish" Then
            Unload Me
        Else
            tabNewData.Tab = 1
        End If
    End If
    
    Select Case iDetailType
        Case DetailType.Asset
            optDataType(0).Caption = "Barcode serial number"
            optDataType(1).Caption = "Asset sticker"
            optDataType(2).Caption = "Product Key"
        Case DetailType.Product
            optDataType(0).Caption = "Manufacturer Code"
            optDataType(1).Caption = "Supplier Code"
            optDataType(2).Caption = "I don't want to use this barcode."
    End Select
    
    If tabNewData.Tab = 0 Then cmdNext.Caption = "Next"
    If tabNewData.Tab = 1 Then cmdNext.Caption = "Finish"
    
End Sub

Private Sub cmdNext_Click()
    DoCommandButton ButtonClick
End Sub

Private Sub tabNewData_Click(PreviousTab As Integer)
    DoCommandButton TabClick
End Sub
