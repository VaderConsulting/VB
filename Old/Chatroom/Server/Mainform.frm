VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Collective"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkServerRunning 
      Caption         =   "Server Running"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3480
      TabIndex        =   12
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtServerPort 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Text            =   "6123"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdStartServer 
      Caption         =   "Start the Server"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdminMsg 
      Caption         =   "Send Admin Message"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txtChatRoom 
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.ListBox lstAllUsers 
      Height          =   2400
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddChatRoom 
      Caption         =   "Add a Chat Room"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ListBox lstChatRoomList 
      Height          =   2400
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock AcceptedSocket 
      Index           =   0
      Left            =   2520
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ListenSocket 
      Left            =   2520
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6123
   End
   Begin VB.Label lblServerPort 
      Caption         =   "On Port"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblChatRoom 
      Caption         =   "Selected User's Chat Room"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label lblUserList 
      Caption         =   "All Connected Users"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblChatRoomList 
      Caption         =   "Chat Room List"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Users As UserCollection             'For keeping track of the user names and
                                        'chat rooms for connected users, as well as
                                        'associating a communication socket for
                                        'each one
Dim ChatRooms As ChatRoomCollection     'This could probably just be a collection of
                                        'Strings, but I wasn't sure of the functionality
                                        'I needed for a chatroom during design time.
                                        'I'll leave it this way for future expandability
Dim WSAStartupData As WSADataType       'This is for use of the Winsock API.  See
                                        'the client code comments for a more complete
                                        'explanation.

Public Sub AcceptedSocket_Close(Index As Integer)
'What to do when a client disconnects
    Dim ChatRoomName As String
    
    'Go ahead and close the socket, if not done already
    'Not sure if this is necessary, but what the hay.
    AcceptedSocket(Index).Close
    
    'Get the appropriate user's chat room
    ChatRoomName = Users(Str(Index)).ChatRoom
    
    'Destroy the invalid user!
    Users.Remove Str(Index)
    
    'Get rid of the Winsock OCX instance
    If Not Index = 0 Then Unload AcceptedSocket(Index)
    
    'Tell everyone in the defunct user's chat room (subtly)
    'that they have been downsized
    NotifyUpdateUserList ChatRoomName
    
    'Update our own list of connected users.
    UpdateConnUsersList
    
End Sub

Private Function UpdateConnUsersList()
'This function resets the user list.  That's all.  Oh, and it sets the caption
'to indicate how many users are connected.
    Dim UserElement As New User
    
    lstAllUsers.Clear
    
    For Each UserElement In Users
        lstAllUsers.AddItem UserElement.UserName
    Next UserElement
    
    MainForm.Caption = "The Collective: " & Str(Users.Count) & " Drones"
    
End Function

Private Function FindRequestType(Data As String) As Integer
'This function simply does a string compare and passes out
'an integer value based on what it found.
    Dim comp As Integer
    
    comp = InStr(1, Data, USER_CHANGE, vbTextCompare)
    
    If comp = 1 Then
        FindRequestType = iUSER_CHANGE
        Exit Function
    End If
    
    comp = InStr(1, Data, CHATROOM_CHANGE, vbTextCompare)
    
    If comp = 1 Then
        FindRequestType = iCHATROOM_CHANGE
        Exit Function
    End If

    comp = InStr(1, Data, INCOMING_MESSAGE, vbTextCompare)
    
    If comp = 1 Then
        FindRequestType = iINCOMING_MESSAGE
        Exit Function
    End If

End Function

Private Sub AcceptedSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'This is a very "meaty" function.  Here is where we do the processing
'of incoming data.  We check for the request type, and do the appropriate
'thing.
    Dim tempString As String
    Dim strValue() As String
    Dim RequestType As Integer
    Dim strCommand() As String
    Dim CommandElement As Variant
    
    'I don't know enough about the Winsock control to know if you have to
    'do multiple reads for split packets, as you do with the API and with
    'UNIX sockets (which are all but identical) but it seems to work as is.
    'Probably because we're not sending large enough packets to be split up.
    ' Receive the data.
    AcceptedSocket(Index).GetData tempString, bytesTotal
    
    'This Split function and the For Each on the resulting array were due
    'to delayed sends from the Winsock OCX.  When I switched to the API for
    'sending data, I don't seem to have the problem of receiving multiple
    'commands in a single receive anymore, but I figured it doesn't hurt
    'to leave it in just in case.
    strCommand = Split(Right(tempString, (Len(tempString) - 1)), "#")

    For Each CommandElement In strCommand
    
        ' Split the string into (hopefully) two values,
        ' right and left of the command separator
        strValue = Split(CommandElement, COMMAND_SEPARATOR)
        
        ' Check to see if we have a data change, such as
        ' chat room change or user name change, or if we
        ' have data to send.
        RequestType = FindRequestType(strValue(0))
        
        If RequestType = iUSER_CHANGE Then
        'We have been notified that a user has changed their user name.
        'This only happens after we send an initial message.
            Users(Str(Index)).OldUserName = Users(Str(Index)).UserName
            Users(Str(Index)).UserName = strValue(1)
            NotifyUserNameChange Index
        ElseIf RequestType = iCHATROOM_CHANGE Then
        'We have been notified that a user has changed chat rooms.  We
        'will send out a message to all users in both the old room and the
        'new room to update their user lists.
            Users(Str(Index)).OldChatRoom = Users(Str(Index)).ChatRoom
            Users(Str(Index)).ChatRoom = strValue(1)
            NotifyUserChatRoomChange Index
        ElseIf RequestType = iINCOMING_MESSAGE Then
        'Some user wants to communicate.  Imagine the bother.  Well, if we
        'must, we'll echo his message to everyone in his chat room.
            Dim tempStr As String
            tempStr = strValue(1)
            NotifyUserMessage Index, tempStr
        End If
    Next CommandElement
              
End Sub

Private Sub NotifyUpdateUserList(RoomName As String)
'We call this function whenever we need to give an updated user list to
'all users in a chat room.
    Dim Element As New User
    Dim strSend As String
    Dim Count As Integer
        
    'Form the first portion of the string
    strSend = "#" & UPDATE_USERLIST & COMMAND_SEPARATOR
    
    'Find the number of users in <RoomName> chatroom
    Count = 0
    For Each Element In Users
        If Element.ChatRoom = RoomName Then
            Count = Count + 1
        End If
    Next Element
        
    'Add the count and the count separator to the string
    strSend = strSend + Str(Count) & COUNT_SEPARATOR
    
    Count = 0
    
    'Add each user name preceeded by a value separator (only if not the
    'first element we are adding).
    For Each Element In Users
        If Element.ChatRoom = RoomName And Not Count = 0 Then
            strSend = strSend + VALUE_SEPARATOR
            strSend = strSend + Element.UserName
            Count = Count + 1
        ElseIf Element.ChatRoom = RoomName Then
            strSend = strSend + Element.UserName
            Count = Count + 1
        End If
    Next Element
    
    ' Send the message
    For Each Element In Users
        If Element.ChatRoom = RoomName Then
            NotifyUserMessageSingle Element.SocketIndex, strSend
        End If
    Next Element
    
End Sub

Private Sub NotifyUserNameChange(Index As Integer)
'We call this after receiving a user name change during initialization of a user
'connection.  This function sends two messages: an updated user list, and the
'welcome message.
    Dim WelcomeMessage As String
    
    NotifyUpdateUserList Users(Str(Index)).ChatRoom
    UpdateConnUsersList
    
    WelcomeMessage = "The Collective welcomes " & Users(Str(Index)).UserName & _
        ".  You have been assimilated."
    NotifyUserMessage Index, WelcomeMessage

End Sub

Private Sub NotifyUserChatRoomChange(Index As Integer)
    ' Update each user's chatroom user list for the old
    ' chat room
    NotifyUpdateUserList Users(Str(Index)).OldChatRoom
    
    ' Update each user's chatroom user list for the new
    ' chat room
    NotifyUpdateUserList Users(Str(Index)).ChatRoom
    
End Sub

Private Sub NotifyUserMessage(Index As Integer, Message As String)
'Echoes a message to everyone in the indexed user's chat room.
    Dim strSend As String
    Dim Element As New User
    
    strSend = "#" & INCOMING_MESSAGE & COMMAND_SEPARATOR & _
        Users(Str(Index)).UserName & VALUE_SEPARATOR & Message
        
    For Each Element In Users
        If Element.ChatRoom = Users(Str(Index)).ChatRoom Then
            'AcceptedSocket(Element.SocketIndex).SendData strSend
            NotifyUserMessageSingle Element.SocketIndex, strSend
        End If
    Next Element

End Sub

Private Sub NotifyUserMessageSingle(Index As Integer, Message As String)
'Echoes a message to only one user.
    send AcceptedSocket(Index).SocketHandle, ByVal Message, Len(Message), 0
End Sub

Private Function UpdateChatRoomListSingle(Index As Integer)
'Sends an updated chat room list to only one user.
    Dim strSend As String
    Dim Element As New ChatRoom
    Dim Count As Integer
    
    'Form the initial portion of the message string
    strSend = "#" & UPDATE_CHATROOM_LIST & COMMAND_SEPARATOR & _
        Str(ChatRooms.Count) & COUNT_SEPARATOR
    
    'Add the chat room count to the string
    Count = 0
    For Each Element In ChatRooms
        If Not Count = 0 Then
            strSend = strSend + VALUE_SEPARATOR
            strSend = strSend + Element.RoomName
            Count = Count + 1
        Else
            strSend = strSend + Element.RoomName
            Count = Count + 1
        End If
    Next Element
    
    'Send the message
    NotifyUserMessageSingle Index, strSend

End Function

Private Sub UpdateChatRoomListAll()
'Sends an updated chat room list to everyone.
    Dim strSend As String
    Dim ChatElement As New ChatRoom
    Dim UserElement As New User
    Dim Count As Integer
    
    'Form the initial portion of the message string
    strSend = "#" & UPDATE_CHATROOM_LIST & COMMAND_SEPARATOR & _
        Str(ChatRooms.Count) & COUNT_SEPARATOR
    
    'Add the chat room count to the string
    Count = 0
    For Each ChatElement In ChatRooms
        If Not Count = 0 Then
            strSend = strSend + VALUE_SEPARATOR
            strSend = strSend + ChatElement.RoomName
            Count = Count + 1
        Else
            strSend = strSend + ChatElement.RoomName
            Count = Count + 1
        End If
    Next ChatElement
    
    'Send the message to all users
    For Each UserElement In Users
        NotifyUserMessageSingle UserElement.SocketIndex, strSend
    Next UserElement
    
End Sub

Private Sub AcceptedSocket_Error(Index As Integer, ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print "Socket Error"
    AcceptedSocket_Close Index
End Sub

Private Sub cmdAddChatRoom_Click()
    Dim NewChatRoom As New ChatRoom
    
    NewChatRoom.RoomName = InputBox("Enter new Chat Room Name:", "Add a New Chat Room")
    lstChatRoomList.AddItem NewChatRoom.RoomName
    ChatRooms.Add NewChatRoom, NewChatRoom.RoomName
        
    UpdateChatRoomListAll
End Sub

Private Sub cmdAdminMsg_Click()
    Dim strAdminMsg As String
    Dim strSend As String
    Dim UserElement As User
        
    strAdminMsg = InputBox("Type the message to send to all users.", "Admin Message")
    strSend = "#" & INCOMING_MESSAGE & COMMAND_SEPARATOR & ADMIN_USERNAME & _
        VALUE_SEPARATOR & strAdminMsg
    
    If Not strSend = "" Then
        For Each UserElement In Users
            NotifyUserMessageSingle UserElement.SocketIndex, strSend
        Next UserElement
    End If
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdStartServer_Click()
    ListenSocket.LocalPort = txtServerPort.Text
    
    ListenSocket.bind
    ListenSocket.listen
    
    txtServerPort.Locked = True
    chkServerRunning.Value = 1
    cmdStartServer.Enabled = False
    cmdAddChatRoom.Default = True
    
End Sub

Private Sub Form_Load()
    Dim DefaultChatRoom As New ChatRoom
    Dim intReturnCode As Integer
    
    SetVars
    
    intReturnCode = WSAStartup(&H101, WSAStartupData)
    
    Set Users = New UserCollection
    Set ChatRooms = New ChatRoomCollection
    
    DefaultChatRoom.RoomName = DEFAULT_CHATROOM
    lstChatRoomList.AddItem DEFAULT_CHATROOM
    ChatRooms.Add DefaultChatRoom, DefaultChatRoom.RoomName
    
End Sub

Private Sub ListenSocket_ConnectionRequest(ByVal requestID As Long)
    ' Create a user to add to the collection
    Dim NewSocketIndex As Integer
    Dim NewUser As New User
    
    ' Find the first available socket index, and load a new
    ' socket into it
    NewSocketIndex = GetAvailableSocketIndex
    
    ' Use the new socket to accept the incoming connection
    AcceptedSocket(NewSocketIndex).accept requestID
    
    ' Set the user's socket index value to the one we just
    ' created, so it can reference it from inside the user
    ' instance
    NewUser.SocketIndex = NewSocketIndex
    
    ' Add the user to the collection with the socket
    ' index as the key to the user
    Users.Add NewUser, Str(NewSocketIndex)
    UpdateConnUsersList
    
    ' Send the init message to the newbie
    SendInitMessage NewSocketIndex
    
    ' Send the ChatRoom list to the newbie
    UpdateChatRoomListSingle NewSocketIndex
       
End Sub

Private Sub SendInitMessage(Index As Integer)
    Dim strSend As String
    
    strSend = "#" & INIT_MESSAGE
    
    'AcceptedSocket(Index).SendData strSend
    'send AcceptedSocket(Index).SocketHandle, _
    '    ByVal strSend, Len(strSend), 0
    NotifyUserMessageSingle Index, strSend


End Sub

Private Function GetAvailableSocketIndex() As Integer
    Dim AvailIndex As Integer
    Dim SocketElement As Variant
    
    AvailIndex = 0
    
    ' First check for available sockets
    
    For Each SocketElement In AcceptedSocket
        If AcceptedSocket(SocketElement.Index).State = sckClosed Then
            If SocketElement.Index <> 0 Then AvailIndex = SocketElement.Index
        End If
    Next SocketElement
    
    ' Next, if AvailIndex is 0 at this point, then we either don't have a created
    ' winsock that is closed, or we only have the original winsock, which we don't
    ' want to use.  Therefore, we need to create a winsock and pass out it's index.
    
    If AvailIndex = 0 Then
        AvailIndex = AcceptedSocket.Count
        Load AcceptedSocket(AvailIndex)
    End If
    
    GetAvailableSocketIndex = AvailIndex
End Function

Private Sub ListenSocket_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print "Listen Socket Error!!!"
    End
End Sub

Private Sub lstAllUsers_Click()
    ' Note: THIS F(N) _ASSUMES_ NO DUPLICATE USERNAMES EXIST
    
    Dim UserElement As User
    Dim Index As Integer
    
    ' Find the user with the selected name
    For Each UserElement In Users
        If UserElement.UserName = lstAllUsers.Text Then
            Index = UserElement.SocketIndex
            Exit For
        End If
    Next UserElement
    
    txtChatRoom.Text = Users(Str(Index)).ChatRoom
    
End Sub

