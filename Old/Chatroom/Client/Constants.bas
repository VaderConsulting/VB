Attribute VB_Name = "Constants"
Public VALUE_SEPARATOR As String
Public COMMAND_SEPARATOR As String
Public COUNT_SEPARATOR As String

Public Const ADMIN_USERNAME = "ADMIN"
Public Const INIT_MESSAGE = "init_message"
Public Const USER_CHANGE = "user_change"
Public Const CHATROOM_CHANGE = "chatroom_change"
Public Const UPDATE_CHATROOM_LIST = "update_chatroom_list"
Public Const INCOMING_MESSAGE = "incoming_message"
Public Const UPDATE_USERLIST = "update_userlist"
Public Const DEFAULT_CHATROOM = "Main"
Public Const iUSER_CHANGE = 1
Public Const iCHATROOM_CHANGE = 2
Public Const iINCOMING_MESSAGE = 3
Public Const iUPDATE_USERLIST = 4
Public Const iUPDATE_CHATROOM_LIST = 5
Public Const iINIT_MESSAGE = 6

Public Sub SetVars()
'Call this function from FormLoad!
'This sets up the value separators to characters that
'hopefully will have a very low chance of being typed.
    VALUE_SEPARATOR = Chr(161)
    COMMAND_SEPARATOR = Chr(162)
    COUNT_SEPARATOR = Chr(163)
End Sub
