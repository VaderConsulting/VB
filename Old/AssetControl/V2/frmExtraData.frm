VERSION 5.00
Begin VB.Form frmExtraData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra Data"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9195
   Icon            =   "frmExtraData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   9195
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtField 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   8
      Top             =   6360
      Width           =   6615
   End
   Begin VB.TextBox txtField 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   6
      Top             =   4800
      Width           =   6615
   End
   Begin VB.TextBox txtField 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   3240
      Width           =   6615
   End
   Begin VB.TextBox txtField 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   6615
   End
   Begin VB.TextBox txtField 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1005
      Index           =   4
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   6720
      Width           =   9015
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1005
      Index           =   3
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   5160
      Width           =   9015
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1005
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   3600
      Width           =   9015
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1005
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2040
      Width           =   9015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   7920
      Width           =   975
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1005
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   480
      Width           =   9015
   End
   Begin VB.Label lblFieldCaption 
      Caption         =   "Field 5 Name"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblFieldCaption 
      Caption         =   "Field 4 Name"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblFieldCaption 
      Caption         =   "Field 3 Name"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblFieldCaption 
      Caption         =   "Field 2 Name"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblFieldCaption 
      Caption         =   "Field 1 Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmExtraData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum IDSource
    Asset = 1
    Product = 2
End Enum

Public Enum Method
    Add = 1
    Update = 2
End Enum

Public strID As String           ' Input
Public strSource As IDSource     ' Input
Public iMethod As Method         ' Input
Public strRecordID As String     ' Output

Private strExtraDataID As String ' Local


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Dim strSQL As String
    Dim oRecordset As New ADODB.Recordset
    
    Select Case iMethod
        Case Method.Add
            ' Add
            strSQL = ""
            strSQL = strSQL & "INSERT INTO tblExtraData"
            strSQL = strSQL & "("
            strSQL = strSQL & "Field1Name, Field1Data, "
            strSQL = strSQL & "Field2Name, Field2Data, "
            strSQL = strSQL & "Field3Name, Field3Data, "
            strSQL = strSQL & "Field4Name, Field4Data, "
            strSQL = strSQL & "Field5Name, Field5Data "
            strSQL = strSQL & ") "
            strSQL = strSQL & "VALUES ("
            strSQL = strSQL & Chr(34) & Me.txtField(0).Text & Chr(34) & ", " & Chr(34) & Me.txtData(0).Text & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & Me.txtField(1).Text & Chr(34) & ", " & Chr(34) & Me.txtData(1).Text & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & Me.txtField(2).Text & Chr(34) & ", " & Chr(34) & Me.txtData(2).Text & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & Me.txtField(3).Text & Chr(34) & ", " & Chr(34) & Me.txtData(3).Text & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & Me.txtField(4).Text & Chr(34) & ", " & Chr(34) & Me.txtData(4).Text & Chr(34)
            strSQL = strSQL & ")"
            
            oConn.Execute strSQL
            
            ' Retrieve the ID of the data we just added
            strSQL = ""
            strSQL = strSQL & "SELECT ID FROM tblExtraData WHERE "
            strSQL = strSQL & "Field1Name = """ & Me.txtField(0).Text & """ AND Field1Data = """ & Me.txtData(0).Text & """ AND "
            strSQL = strSQL & "Field2Name = """ & Me.txtField(1).Text & """ AND Field2Data = """ & Me.txtData(1).Text & """ AND "
            strSQL = strSQL & "Field3Name = """ & Me.txtField(2).Text & """ AND Field3Data = """ & Me.txtData(2).Text & """ AND "
            strSQL = strSQL & "Field4Name = """ & Me.txtField(3).Text & """ AND Field4Data = """ & Me.txtData(3).Text & """ AND "
            strSQL = strSQL & "Field5Name = """ & Me.txtField(4).Text & """ AND Field5Data = """ & Me.txtData(4).Text & """ "
            
            Set oRecordset = oConn.Execute(strSQL)
            strRecordID = oRecordset.Fields("ID")
            
        Case Method.Update
            strSQL = ""
            strSQL = strSQL & "UPDATE tblExtraData "
            strSQL = strSQL & "SET "
            strSQL = strSQL & "Field1Name = """ & Me.txtField(0).Text & """, Field1Data = """ & Me.txtData(0).Text & """, "
            strSQL = strSQL & "Field2Name = """ & Me.txtField(1).Text & """, Field2Data = """ & Me.txtData(1).Text & """, "
            strSQL = strSQL & "Field3Name = """ & Me.txtField(2).Text & """, Field3Data = """ & Me.txtData(2).Text & """, "
            strSQL = strSQL & "Field4Name = """ & Me.txtField(3).Text & """, Field4Data = """ & Me.txtData(3).Text & """, "
            strSQL = strSQL & "Field5Name = """ & Me.txtField(4).Text & """, Field5Data = """ & Me.txtData(4).Text & """ "
            strSQL = strSQL & "WHERE ID = " & strExtraDataID
            
            oConn.Execute strSQL
    End Select
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim iCount As Integer
    Dim strTable As String
    
    Select Case strSource
        Case IDSource.Asset
            strTable = "tblAssets"
        Case IDSource.Product
            strTable = "tblProducts"
    End Select
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "tblExtraData.ID, "
    strSQL = strSQL & "tblExtraData.Field1Name, tblExtraData.Field1Data, "
    strSQL = strSQL & "tblExtraData.Field2Name, tblExtraData.Field2Data, "
    strSQL = strSQL & "tblExtraData.Field3Name, tblExtraData.Field3Data, "
    strSQL = strSQL & "tblExtraData.Field4Name, tblExtraData.Field4Data, "
    strSQL = strSQL & "tblExtraData.Field5Name, tblExtraData.Field5Data "
    strSQL = strSQL & "FROM " & strTable & " "
    strSQL = strSQL & "INNER JOIN tblExtraData "
    strSQL = strSQL & "ON "
    strSQL = strSQL & strTable & ".ExtraDataID = tblExtraData.ID "
    strSQL = strSQL & "WHERE " & strTable & ".ID = " & strID
    
    iCount = GetRecords(strSQL)
    
    If iCount > 0 Then ' Update
        txtField(0).Text = oDatabase.Fields("Field1Name") & ""
        txtData(0).Text = oDatabase.Fields("Field1Data") & ""
        txtField(1).Text = oDatabase.Fields("Field2Name") & ""
        txtData(1).Text = oDatabase.Fields("Field2Data") & ""
        txtField(2).Text = oDatabase.Fields("Field3Name") & ""
        txtData(2).Text = oDatabase.Fields("Field3Data") & ""
        txtField(3).Text = oDatabase.Fields("Field4Name") & ""
        txtData(3).Text = oDatabase.Fields("Field4Data") & ""
        txtField(4).Text = oDatabase.Fields("Field5Name") & ""
        txtData(4).Text = oDatabase.Fields("Field5Data") & ""
        
        strExtraDataID = oDatabase.Fields("ID")
    Else ' Add
        lblFieldCaption(0).Visible = True
        lblFieldCaption(1).Visible = True
        lblFieldCaption(2).Visible = True
        lblFieldCaption(3).Visible = True
        lblFieldCaption(4).Visible = True
        
        txtField(0).Locked = False
        txtField(1).Locked = False
        txtField(2).Locked = False
        txtField(3).Locked = False
        txtField(4).Locked = False
    End If
    
    
    
End Sub
