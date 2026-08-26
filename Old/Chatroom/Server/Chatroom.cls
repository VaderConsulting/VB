VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mOwner As String        'Not used as yet ;)
Private mRoomName As String

Property Get Owner() As String
    Owner = mOwner
End Property

Property Let Owner(aUserName As String)
    mOwner = aUserName
End Property

Property Get RoomName() As String
    RoomName = mRoomName
End Property

Property Let RoomName(aRoomName As String)
    mRoomName = aRoomName
End Property
