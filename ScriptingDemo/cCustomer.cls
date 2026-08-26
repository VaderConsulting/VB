VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Support class for colCustomers, models a customer from
'the NorthWind database
Public CompanyName As String
Public ContactName As String
Public Address As String
Public City As String
Public State As String
Public PostalCode As String
Public Country As String
Public Phone As String
Public Fax As String
Private ID As String

Public Property Get CustomerId() As String
    CustomerId = ID
End Property

Friend Property Let CustomerId(value As String)
    ID = value
End Property
