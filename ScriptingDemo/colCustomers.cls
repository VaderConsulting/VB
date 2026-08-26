VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Custom collection modelling customers from the NorthWind database
'Used for a play object for ScriptingDemo.vbp
Private m_oCust As cCustomer
Private m_oCol As Collection

Private m_rsCust As ADODB.Recordset

Private Sub Class_Initialize()
    Set m_oCol = New Collection
End Sub

Public Sub Add(oCust As cCustomer)
    m_oCol.Add oCust, oCust.CustomerId
End Sub

Public Property Get Count() As Integer
    Count = m_oCol.Count
End Property

Public Sub Remove(vIndex As Variant)
    m_oCol.Remove vIndex
End Sub
Public Function Item(vIndex As Variant) As cCustomer
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Err.Clear
    Set Item = m_oCol.Item(vIndex)
    If Err.Number > 0 Then
        Err.Clear
        If IsNumeric(vIndex) Then
            On Error GoTo 0
            Set Item = m_oCol.Item(CInt(vIndex))
        End If
     End If
End Function
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Provides support for For.. Each
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub LoadData(strDSN As String, strUID As String, strPWD As String)
    Dim cn As New ADODB.Connection
    Dim oCust As New cCustomer
    
    cn.ConnectionString = "DSN=" & strDSN & ";UID=" & strUID & ";PWD=" & strPWD & ";APP=Visual Basic;DATABASE=Northwind"
    cn.Open
    Set m_rsCust = cn.Execute("SELECT * FROM CUSTOMERS;")
    m_rsCust.MoveFirst
    Do Until m_rsCust.EOF
        Set oCust = New cCustomer
        oCust.CustomerId = m_rsCust("CustomerId") & ""
        oCust.CompanyName = m_rsCust("CompanyName") & ""
        oCust.ContactName = m_rsCust("ContactName") & ""
        oCust.Address = m_rsCust("Address") & ""
        oCust.City = m_rsCust("City") & ""
        oCust.PostalCode = m_rsCust("PostalCode") & ""
        oCust.Country = m_rsCust("Country") & ""
        oCust.Phone = m_rsCust("Phone") & ""
        oCust.Fax = m_rsCust("Fax") & ""
        Add oCust
        m_rsCust.MoveNext
    Loop
        
End Sub
