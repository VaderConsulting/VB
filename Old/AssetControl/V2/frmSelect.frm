VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdSelect 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4048
      _Version        =   393216
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strQuery As String
Public strFieldToReturn As String
Public strReturnFieldDisplayName As String
Public strSelectedItem As String
Public strProductID As String

Public Sub GetData()
    Dim iCount As Integer
    Dim iCol As Integer
    Dim iRow As Integer
    Dim strCaptions(1) As String
    Dim iWidth(1) As Integer
    
    strCaptions(0) = "ID"
    strCaptions(1) = strReturnFieldDisplayName
    
    iWidth(0) = 600
    iWidth(1) = 2500
    
    grdSelect.Row = 0
    For iCol = 0 To 1
        grdSelect.Col = iCol
        grdSelect.Text = strCaptions(iCol)
        grdSelect.ColWidth(iCol) = iWidth(iCol)
        grdSelect.ColAlignment(iCol) = flexAlignLeftCenter
    Next iCol
    
    
    Me.Refresh
    
    iCount = GetRecords(strQuery)
    
    If iCount > 0 Then
    
        grdSelect.Rows = iCount + 1
        
        iRow = 1
        
        Do Until oDatabase.EOF
            ' Prevented from adding blank records
            If oDatabase.Fields(1) & "" <> "" Then
                iCol = 0
                For Each Field In oDatabase.Fields
                    grdSelect.Col = iCol
                    grdSelect.Row = iRow
                    grdSelect.Text = oDatabase.Fields(iCol) & ""
                    iCol = iCol + 1
                Next
                iRow = iRow + 1
            End If
            oDatabase.MoveNext
        Loop
        
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    frmNewProduct.strSelected = strSelectedItem
    Unload Me
End Sub

Private Sub grdSelect_Click()
    Dim iRow As Integer
    Dim strDescription As String
    Dim strID As String
    
    iRow = grdSelect.Row
    
    strID = grdSelect.TextMatrix(iRow, 0)
    strDescription = grdSelect.TextMatrix(iRow, 1)
    
    ' return this info through the forms public variables
    strSelectedItem = strDescription
    strProductID = strID
End Sub
