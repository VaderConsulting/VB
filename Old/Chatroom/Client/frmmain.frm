VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Collective - Drone"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Text            =   "6123"
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtServerName 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   5055
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.ComboBox cmbChatRooms 
      Height          =   315
      Left            =   5520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.ListBox lstUsers 
      Height          =   3375
      Left            =   5520
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtMain 
      Height          =   4095
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
   End
   Begin MSWinsockLib.Winsock ConnSocket 
      Left            =   7800
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblUserName 
      Caption         =   "Your User Name"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblServerName 
      Caption         =   "Server"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblMessages 
      Caption         =   "Messages"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblChatRooms 
      Caption         =   "Chat Rooms"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblData 
      Caption         =   "Type Your Message and Click ""Send,"" or Hit Enter"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Label lblUserList 
      Caption         =   "People in this Room"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim UserName As String
Dim Port As Integer                 'The remote server port
Dim CurrentChatRoom As String       'The chat room you are in
Dim ServerName As String            'The network name (or IP) of
                                    'your chosen chat server
Dim WSAStartupData As WSADataType   'For the Winsock API:
                                    'This is because I chose to
                                    'use part of the Winsock API
                                    'to "adjust" for the OCX bug
                                    'where it sometimes doesn't send
                                    'output for a while, then all at
                                    'once.  Not good in a real-time
                                    'communication application

Option Explicit

Private Sub cmbChatRooms_Click()
'When a user selects a chatroom from the dropdown list, we send a
'message to the server to notify of the change.  The server then
'will send an updated user list message to everyone (including us)
'in the old chat room and the new chat room.
    Dim strSend As String
    
    'Create the message string
    strSend = "#" & CHATROOM_CHANGE & COMMAND_SEPARATOR & cmbChatRooms.Text
    'Use the API to send the string
    send ConnSocket.SocketHandle, ByVal strSend, Len(strSend), 0
    
    CurrentChatRoom = cmbChatRooms.Text
End Sub

Private Sub cmdConnect_Click()
'When a user clicks the Connect button, some conditions must be met
'first.  There must be text in the ServerName, Port, and UserName
'text boxes.  This is taken care of in the change event procedures
'for those controls.
    Dim strSend As String
    
    'Set the port and server name from the text boxes on the form.
    'If either of these are blank, the Connect button will not be
    'enabled.
    Port = txtPort.Text
    ServerName = txtServerName.Text
        
    'Error handling for the socket is now handled in the Error event,
    'as it should be.
    ConnSocket.connect ServerName, Port
    
    'Lock down the text boxes while we are connected, disable the
    'Connect button, and enable the Send and Disconnect buttons.
    txtServerName.Locked = True
    txtPort.Locked = True
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    cmdSend.Enabled = True

End Sub

Private Sub cmdDisconnect_Click()
'Closes the connection and performs some GUI cleanup.

    'Actually close the connection
    ConnSocket.Close
    
    'Clear the User list box and chat rooms list
    lstUsers.Clear
    cmbChatRooms.Clear
    
    'Enable editing of the text boxes, disable the disconnect and send
    'buttons, and enable the connect button.
    txtServerName.Locked = False
    txtPort.Locked = False
    cmdDisconnect.Enabled = False
    cmdSend.Enabled = False
    cmdConnect.Enabled = True
End Sub

Private Sub cmdSend_Click()
'When the user clicks Send, we send the text in the txtSend control
'to the server.  The server then sends everyone a copy of the message
'who is in our chat room with us (including us).
    Dim strSend As String
    
    'If we have no text, do nothing.
    If Len(txtSend.Text) = 0 Then
        Exit Sub
    End If
    
    'Form the message to send
    strSend = "#" & INCOMING_MESSAGE & COMMAND_SEPARATOR & txtSend.Text
    
    'Actually send the message
    send ConnSocket.SocketHandle, ByVal strSend, Len(strSend), 0

    'Set the text in the txtSend control to nothing.
    txtSend.Text = ""
End Sub

Private Function FindRequestType(Data As String) As Integer
'This function simply does a string compare and passes out
'an integer value based on what it found.
    Dim comp As Integer
    
    comp = InStr(1, Data, INIT_MESSAGE, vbTextCompare)
    
    If comp = 1 Then
        FindRequestType = iINIT_MESSAGE
        Exit Function
    End If
    
    comp = InStr(1, Data, UPDATE_USERLIST, vbTextCompare)
    
    If comp = 1 Then
        FindRequestType = iUPDATE_USERLIST
        Exit Function
    End If
    
    comp = InStr(1, Data, UPDATE_CHATROOM_LIST, vbTextCompare)
    
    If comp = 1 Then
        FindRequestType = iUPDATE_CHATROOM_LIST
        Exit Function
    End If

    comp = InStr(1, Data, INCOMING_MESSAGE, vbTextCompare)
    
    If comp = 1 Then
        FindRequestType = iINCOMING_MESSAGE
        Exit Function
    End If

End Function

Private Sub ConnSocket_Close()
'What to do when the server disconnects us
    cmdDisconnect_Click
End Sub

Private Sub ConnSocket_DataArrival(ByVal bytesTotal As Long)
'This is a very "meaty" function.  Here is where we do the processing
'of incoming data.  We check for the request type, and do the appropriate
'thing.
    Dim CommandElement As Variant
    Dim tempString As String
    Dim strCommand() As String
    Dim strValue() As String
    Dim RequestType As Integer
    Dim strChatList() As String
    Dim strCount() As String
    Dim Count As Double
    Dim Counter As Double
    Dim strUserName As String
    Dim strUserList() As String
    Dim strNameAndMessage() As String
    Dim strMessage As String
    Dim strOutput As String
    Dim strSend As String

    
    'I don't know enough about the Winsock control to know if you have to
    'do multiple reads for split packets, as you do with the API and with
    'UNIX sockets (which are all but identical) but it seems to work as is.
    'Probably because we're not sending large enough packets to be split up.
    ' Receive the data.
    ConnSocket.GetData tempString, bytesTotal
    
    'This Split function and the For Each on the resulting array were due
    'to delayed sends from the Winsock OCX.  When I switched to the API for
    'sending data, I don't seem to have the problem of receiving multiple
    'commands in a single receive anymore, but I figured it doesn't hurt
    'to leave it in just in case.
    strCommand = Split(Right(tempString, (Len(tempString) - 1)), "#")

    For Each CommandElement In strCommand
        
        ' Split the string into all of it glorious values,
        ' using the command separator as a value separator
        strValue = Split(CommandElement, COMMAND_SEPARATOR)
    
        ' Check to see if we have a data change, such as
        ' chat room change or user name change, or if we
        ' have data to display.
        RequestType = FindRequestType(strValue(0))
        
        If RequestType = iUPDATE_USERLIST Then
        'We have received a message to update our user list, which
        'will include all the user names in our current chat room,
        'preceded by the user name count for that chat room.
            lstUsers.Clear
            
            'Here we extract the number of user names in the message
            'from the count value
            strCount = Split(strValue(1), COUNT_SEPARATOR)
            Count = Val(strCount(0))
            
            'Here we split the string of user names into it's own array
            strUserList = Split(strCount(1), VALUE_SEPARATOR)
            
            'And then add them all to the user list control
            For Counter = 0 To Count - 1
                lstUsers.AddItem strUserList(Counter)
            Next
        ElseIf RequestType = iINIT_MESSAGE Then
        'We have received the initial message from the server, which
        'is our cue to tell them who we are.  We will only receive
        'this message once each connection.
        
            'Form the data string to send
            strSend = "#" & USER_CHANGE & COMMAND_SEPARATOR & UserName
        
            'Actually send the message, using the API
            send ConnSocket.SocketHandle, ByVal strSend, Len(strSend), 0
        
            'Enter the default chat room, so we are where the server
            'thinks we are
            CurrentChatRoom = DEFAULT_CHATROOM
            cmbChatRooms.AddItem DEFAULT_CHATROOM
            cmbChatRooms.Text = DEFAULT_CHATROOM
    
        ElseIf RequestType = iUPDATE_CHATROOM_LIST Then
        'We have been told to update our chat room list, usually because
        'the administrator has added a chat room, or we have just signed on.
        'This message contains a count of chat rooms, and the name of each
        'chat room.
            cmbChatRooms.Clear
            
            'Here we extract the number of chat rooms from the message
            strCount = Split(strValue(1), COUNT_SEPARATOR)
            Count = Val(strCount(0))
            
            'Here we split out the string of chat room names into its own
            'array
            strChatList = Split(strCount(1), VALUE_SEPARATOR)
            
            'And add them all to our chat room list control
            For Counter = 0 To Count - 1
                cmbChatRooms.AddItem strChatList(Counter)
            Next
            
            'Now we give the chat room list control our current information
            cmbChatRooms.Text = CurrentChatRoom
            
       ElseIf RequestType = iINCOMING_MESSAGE Then
       'We have received an incoming message from someone!  They must like us
       'if they want to talk to us...  We eats it, my precious.
       
            'Here we split the user name of the sender and the actual message
            'into an array
            strNameAndMessage = Split(strValue(1), VALUE_SEPARATOR)
            strUserName = strNameAndMessage(0)
            strMessage = strNameAndMessage(1)
            
            'Form the output
            strOutput = "[" & strUserName & "] " & strMessage
            
            'And display it
            txtMain.Text = txtMain.Text & vbCrLf & strOutput
            
        End If
    Next CommandElement
        
End Sub

Private Sub ConnSocket_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "The connection socket encountered an error.  Please reconnect.  Error: " & Description
    cmdDisconnect_Click
End Sub

Private Sub txtMain_Change()
    'This is so we will always see the end of the text in the text box.
    'If you don't do this, it get's anoying for long multiline data, as
    'you have to scroll down again every time something new is added.
    txtMain.SelStart = Len(txtMain.Text)
End Sub

Private Sub Form_Load()
    'From UserNames.mod
    SetVars
    
    'For the Winsock API
    WSAStartup &H101, WSAStartupData
    
    'Now this gets your login name for your local box.  In other words,
    'doing it this way is preferable on a corporate intranet, but probably
    'not over the internet.
    txtUserName.Text = UserName
    
End Sub


Private Sub txtPort_Change()
    If txtServerName.Text = "" _
        Or txtPort.Text = "" _
        Or txtUserName = "" _
        Or ConnSocket.State = sckConnected Then
        cmdConnect.Enabled = False
    Else
        cmdConnect.Enabled = True
    End If
    
End Sub

Private Sub txtServerName_Change()
    If txtServerName.Text = "" _
        Or txtPort.Text = "" _
        Or txtUserName = "" _
        Or ConnSocket.State = sckConnected Then
        cmdConnect.Enabled = False
    Else
        cmdConnect.Enabled = True
    End If
    
End Sub

Private Sub txtUserName_Change()
    If txtServerName.Text = "" _
        Or txtPort.Text = "" _
        Or txtUserName = "" _
        Or ConnSocket.State = sckConnected Then
        cmdConnect.Enabled = False
    Else
        cmdConnect.Enabled = True
    End If

End Sub

