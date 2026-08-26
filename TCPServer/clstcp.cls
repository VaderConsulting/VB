VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum ConnState
    CONN_STATE_CONNECTED = 0
    CONN_STATE_WAITING_FOR_AUTHORISATION
    
End Enum

'local variable(s) to hold property value(s)
Private mvarConnectionState As Integer 'local copy
Public Index As Integer
Attribute Index.VB_VarDescription = "Defines the index used by the Winsock object."

Public Property Let ConnectionState(ByVal vData As Integer)
Attribute ConnectionState.VB_Description = "Internal State Machine.  This Integer is used to represent what state the object is in."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConnectionState = 5
    mvarConnectionState = vData
End Property


Public Property Get ConnectionState() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConnectionState
    ConnectionState = mvarConnectionState
End Property



