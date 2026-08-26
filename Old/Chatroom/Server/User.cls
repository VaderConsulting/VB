VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ChatRoom"
Private mUserName As String
Private mOldUserName As String
Private mChatRoom As String
Private mOldChatRoom As String
Private mLimboReady As Boolean
Private mMySocketIndex As Integer

Property Get SocketIndex() As Integer
    SocketIndex = mMySocketIndex
End Property

Property Let SocketIndex(anIndex As Integer)
    mMySocketIndex = anIndex
    mUserName = Str(mMySocketIndex)
End Property

Property Get UserName() As String
    UserName = mUserName
End Property

Property Let UserName(aUserName As String)
    mUserName = aUserName
End Property

Property Get OldUserName() As String
    OldUserName = mOldUserName
End Property

Property Let OldUserName(aUserName As String)
    mOldUserName = aUserName
End Property

Property Get ChatRoom() As String
    ChatRoom = mChatRoom
End Property

Property Let ChatRoom(aChatRoom As String)
    mChatRoom = aChatRoom
End Property

Property Get OldChatRoom() As String
    OldChatRoom = mOldChatRoom
End Property

Property Let OldChatRoom(aChatRoom As String)
    mOldChatRoom = aChatRoom
End Property

Private Sub Class_Initialize()
    mChatRoom = DEFAULT_CHATROOM
    mLimboReady = False
End Sub

Private Sub Class_Terminate()
    ' MainForm.AcceptedSocket_Close mMySocketIndex
End Sub
