VERSION 5.00
Begin VB.Form frmNewProduct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Product"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "frmNewProduct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   5
      Left            =   5290
      TabIndex        =   11
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   4
      Left            =   5290
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   3
      Left            =   5290
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   2
      Left            =   5290
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   1
      Left            =   5290
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Index           =   0
      Left            =   5290
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   10
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   8
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblField 
      Caption         =   "NSN"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblField 
      Caption         =   "Supplier Code"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblField 
      Caption         =   "Manufacturer Code"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblField 
      Caption         =   "Supplier"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblField 
      Caption         =   "Manufacturer"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblField 
      Caption         =   "Description"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strSelected As String
Public strReturnID As String

Private Sub cmdBrowse_Click(Index As Integer)
    Dim strFieldToReturn As String
    Dim strQuery As String
    Dim strResult As String
    
    Select Case Index
        Case 0 ' Description
            strFieldToReturn = "Description"
        Case 1 ' Manufacturer
            strFieldToReturn = "Manufacturer"
        Case 2 ' Supplier
            strFieldToReturn = "Supplier"
        Case 3 ' ManufacturerCode
            strFieldToReturn = "ManufacturerCode"
        Case 4 ' SupplierCode
            strFieldToReturn = "SupplierCode"
        Case 5 ' NSN
            strFieldToReturn = "NSN"
    End Select
    
    strQuery = "SELECT ID, " & strFieldToReturn & " FROM tblProducts ORDER BY " & strFieldToReturn
    frmSelect.strReturnFieldDisplayName = lblField(Index).Caption
    frmSelect.strFieldToReturn = strFieldToReturn
    frmSelect.strQuery = strQuery
    
    frmSelect.GetData
    
    frmSelect.Show vbModal
    strResult = frmSelect.strSelectedItem
    
    Set frmSelect = Nothing
    txtData(Index).Text = strResult
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim iCount As Integer
    
    ' Before we add the record, lets see if one already exists.
    strSQL = ""
    strSQL = strSQL & "SELECT ID FROM tblProducts WHERE "
    strSQL = strSQL & "Description = """ & txtData(0).Text & """ AND "
    strSQL = strSQL & "Manufacturer = """ & txtData(1).Text & """ AND "
    strSQL = strSQL & "Supplier = """ & txtData(2).Text & """ AND "
    strSQL = strSQL & "ManufacturerCode = """ & txtData(3).Text & """ AND "
    strSQL = strSQL & "SupplierCode = """ & txtData(4).Text & """ AND "
    strSQL = strSQL & "NSN = """ & txtData(5).Text & """"
    
    iCount = GetRecords(strSQL)
    
    If iCount = 0 Then
        ' Not found.  Lets add it
        strSQL = ""
        strSQL = strSQL & "INSERT INTO tblProducts "
        strSQL = strSQL & "(Description, Manufacturer, Supplier, ManufacturerCode, SupplierCode, NSN) VALUES "
        strSQL = strSQL & "(""" & txtData(0).Text & """,""" & txtData(1).Text & """,""" & txtData(2).Text & """,""" & txtData(3).Text & """,""" & txtData(4).Text & """,""" & txtData(5).Text & """)"
        frmMain.lblInfo.Caption = "Product Record added."
        
        oConn.Execute strSQL
        
        ' Get the ID of this product
        strSQL = ""
        strSQL = strSQL & "SELECT ID FROM tblProducts WHERE "
        strSQL = strSQL & "Description = """ & txtData(0).Text & """ AND "
        strSQL = strSQL & "Manufacturer = """ & txtData(1).Text & """ AND "
        strSQL = strSQL & "Supplier = """ & txtData(2).Text & """ AND "
        strSQL = strSQL & "ManufacturerCode = """ & txtData(3).Text & """ AND "
        strSQL = strSQL & "SupplierCode = """ & txtData(4).Text & """ AND "
        strSQL = strSQL & "NSN = """ & txtData(5).Text & """"
        
        iCount = GetRecords(strSQL)
    Else
        frmMain.lblInfo.Caption = "Product already exists."
    End If
    
    ' Now return the ID
    If iCount > 0 Then
        strReturnID = oDatabase.Fields(0)
    End If
    
    Unload Me
End Sub

