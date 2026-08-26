VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Barcode to Access Interface"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5145
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox lstStatus 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4935
   End
   Begin VB.CommandButton cmdManage 
      Caption         =   "Manage..."
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtBarcode 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label lblBarcode 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle.  Waiting for Model/Part Number."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoConn As ADODB.Connection
Dim aodRS As ADODB.Recordset

Dim adoConn2 As ADODB.Connection
Dim aodRS2 As ADODB.Recordset

Dim DSN As String

Dim Mode As Integer

Private Sub ReceiveInfo(Text As String)
    Dim SQL As String, WillInsert As Boolean
    Static Model As String, Serial As String
    
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    
    DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "Barcodes.mdb;Persist Security Info=False"
    Mode = Mode + 1
    
    Select Case Mode
        Case 1
            lblStatus = "Received Model/Part Number.  Waiting for Serial No"
            lstStatus.AddItem "Received Model/Part Number.  Waiting for Serial No", 0
            Model = Text
            lblBarcode = Text
            txtBarcode = ""
        Case 2
            lblStatus = "Received Serial No.  Waiting for Asset No"
            lstStatus.AddItem "Received Serial No.  Waiting for Asset No", 0
            Serial = Text
            lblBarcode = Text
            txtBarcode = ""
            
            ' Do we already have this item?
            If InDatabase(Serial, Model) = True Then
                lblStatus = "Waiting for Model/Part Number."
                lstStatus.AddItem "This item is already in the database.", 0
                Mode = 0
                WillInsert = False
                Serial = ""
                Model = ""
                Text = ""
                Exit Sub
            End If
            
            ' Should it have an Asset Number?
            If HasAsset(Model) = False Then
                WillInsert = True
                Text = "N/A"
            Else
                WillInsert = False
            End If
        Case 3
            txtBarcode = ""
            WillInsert = True
    End Select
    
    If WillInsert Then
        ' Get details for this model and part number
        SQL = "SELECT tblModels.ID FROM tblModels INNER JOIN tblTypes ON tblModels.TypeID = tblTypes.ID WHERE tblModels.Barcode ='" & Model & "'"
        adoConn.Open DSN
        adoRS.Open SQL, adoConn
        If Not adoRS.EOF Then
            SQL = "INSERT INTO tblItems (Serial, Asset, ModelID) VALUES ('" & Serial & "','" & Text & "','" & adoRS("ID") & "')"
            adoConn.Execute SQL
            lstStatus.AddItem "Inserted item into Database.", 0
        End If
        lblStatus = "Received all Details.  Waiting for Model/Part Number."
        WillInsert = False
        Text = ""
        Mode = 0
        lblBarcode = ""
    End If
    
    lblStatus.Refresh
End Sub

Private Function HasAsset(Model As String) As Boolean
    SQL = "SELECT tblTypes.AssetReqd FROM tblModels INNER JOIN tblTypes ON tblModels.TypeID = tblTypes.ID WHERE tblModels.Barcode ='" & Model & "'"
    
    Set adoConn2 = CreateObject("ADODB.Connection")
    Set adors2 = CreateObject("ADODB.Recordset")
    adoConn2.Open DSN
    adors2.Open SQL, adoConn2
    If Not adors2.EOF Then
        If adors2("AssetReqd") = True Then
            HasAsset = True
        Else
            HasAsset = False
        End If
    Else
        lblStatus = "Model not found!"
    End If
    adors2.Close
    adoConn2.Close
End Function

Private Function InDatabase(Serial As String, Model As String) As Boolean
    SQL = "SELECT tblItems.Serial, tblModels.Barcode FROM tblItems INNER JOIN tblModels ON tblItems.ModelID = tblModels.ID WHERE (((tblItems.Serial)='" & Serial & "') AND ((tblModels.Barcode)='" & Model & "'))"
    
    Set adoConn2 = CreateObject("ADODB.Connection")
    Set adors2 = CreateObject("ADODB.Recordset")
    adoConn2.Open DSN
    adors2.Open SQL, adoConn2
    If Not adors2.EOF Then
        InDatabase = True
    End If
    adors2.Close
    adoConn2.Close
End Function

Private Sub cmdReset_Click()
    Mode = 0
    lblStatus = "Reset Status.  Waiting for Model/Part Number."
End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
     If KeyAscii = Asc(vbCr) Or KeyAscii = Asc(vbCrLf) Then
        ReceiveInfo Trim(txtBarcode.Text)
    End If
End Sub
