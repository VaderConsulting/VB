VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   11430
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   4200
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdBrowse 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   7
      AllowBigSelection=   -1  'True
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10320
      TabIndex        =   1
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Top             =   4200
      Width           =   975
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strQuery As String
Public strSelectedItem As String
Public strProductID As String
Public iColToReturn As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub GetData()
    Dim iCount As Integer
    Dim iCol As Integer
    Dim iRow As Integer
    Dim strCaptions(6) As String
    Dim iWidth(6) As Integer
    
    strCaptions(0) = "ID"
    strCaptions(1) = "Description"
    strCaptions(2) = "Manufacturer"
    strCaptions(3) = "Supplier"
    strCaptions(4) = "Manufacturer Code"
    strCaptions(5) = "Supplier Code"
    strCaptions(6) = "NSN"
    
    iWidth(0) = 600
    iWidth(1) = 2200
    iWidth(2) = 2200
    iWidth(3) = 2200
    iWidth(4) = 2200
    iWidth(5) = 2200
    iWidth(6) = 2200
    
    grdBrowse.Row = 0
    For iCol = 0 To 6
        grdBrowse.Col = iCol
        grdBrowse.Text = strCaptions(iCol)
        grdBrowse.ColWidth(iCol) = iWidth(iCol)
        grdBrowse.ColAlignment(iCol) = flexAlignLeftCenter
    Next iCol
    
    
    Me.Refresh
    
    iCount = GetRecords(strQuery)
    
    If iCount > 0 Then
    
        grdBrowse.Rows = iCount + 1
        
        iRow = 1
        
        Do Until oDatabase.EOF
            ' Prevented from adding blank records
            If oDatabase.Fields(1) & "" <> "" Then
                'lstBrowse.AddItem oDatabase.Fields(1) & ""
                iCol = 0
                For Each Field In oDatabase.Fields
                    grdBrowse.Col = iCol
                    grdBrowse.Row = iRow
                    grdBrowse.Text = oDatabase.Fields(iCol) & ""
                    iCol = iCol + 1
                Next
                iRow = iRow + 1
            End If
            oDatabase.MoveNext
        Loop
        
    End If
End Sub

Private Sub cmdNew_Click()
    frmNewProduct.Show vbModal
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub grdBrowse_Click()
    Dim iRow As Integer
    Dim strDescription As String
    Dim strID As String
    
    iRow = grdBrowse.Row
    
    strID = grdBrowse.TextMatrix(iRow, 0)
    strDescription = grdBrowse.TextMatrix(iRow, iColToReturn)
    
    ' return this info through the forms public variables
    strSelectedItem = strDescription
    strProductID = strID
End Sub
