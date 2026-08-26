VERSION 5.00
Begin VB.UserControl uPing 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "Ping.ctx":0000
   ScaleHeight     =   390
   ScaleWidth      =   1215
   ToolboxBitmap   =   "Ping.ctx":0013
   Begin VB.TextBox txtTimeout 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "3000"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtRoundTripTime 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPing 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ping Control"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      Caption         =   "Result"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "uPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
Const m_def_Octet1 = "127"
Const m_def_Octet2 = "0"
Const m_def_Octet3 = "0"
Const m_def_Octet4 = "1"
Const m_def_IPAddress = "127.0.0.1"
'Property Variables:
Dim m_Octet1 As String
Dim m_Octet2 As String
Dim m_Octet3 As String
Dim m_Octet4 As String
Dim m_IPAddress As String
'Dim m_IPAddress As String
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=cmdPing,cmdPing,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=cmdPing,cmdPing,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=cmdPing,cmdPing,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Sub cmdPing_Click()
    Dim fMatch As Boolean, sRTT As String, sHost As String
    
    PING_TIMEOUT = txtTimeout
    txtIP(0) = Octet1
    txtIP(1) = Octet2
    txtIP(2) = Octet3
    txtIP(3) = Octet4
    If IPAddress <> "" Then
        sHost = IPAddress
    Else
        sHost = txtIP(0) & "." & txtIP(1) & "." & txtIP(2) & "." & txtIP(3)
    End If
    
    If sHost <> "" Then
        pingresult = Ping1(sHost, sRTT, True, 32, PING_TIMEOUT)
        Select Case pingresult
        Case "Error"
            txtRoundTripTime = "Error"
        Case "458772"
            txtRoundTripTime = "Timeout"
        Case Else
            txtRoundTripTime = pingresult '& " mS"
        End Select
    End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdPing,cmdPing,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "IPProperties"
    Enabled = cmdPing.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    cmdPing.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub cmdPing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub cmdPing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cmdPing_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,127.000.000.001
'Public Property Get IPAddress() As String
'    IPAddress = m_IPAddress
'End Property
'
'Public Property Let IPAddress(ByVal New_IPAddress As String)
'    m_IPAddress = New_IPAddress
'    PropertyChanged "IPAddress"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtRoundTripTime,txtRoundTripTime,-1,Text
Public Property Get RoundTripTime() As String
Attribute RoundTripTime.VB_Description = "Returns/sets the text contained in the control."
    RoundTripTime = txtRoundTripTime.Text
End Property

Public Property Let RoundTripTime(ByVal New_RoundTripTime As String)
    txtRoundTripTime.Text() = New_RoundTripTime
    PropertyChanged "RoundTripTime"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13
Public Function Ping() As String
    cmdPing_Click
End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    txtIP(0) = ""
'    m_IPAddress = m_def_IPAddress
    m_IPAddress = m_def_IPAddress
    m_Octet1 = m_def_Octet1
    m_Octet2 = m_def_Octet2
    m_Octet3 = m_def_Octet3
    m_Octet4 = m_def_Octet4
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    cmdPing.Enabled = PropBag.ReadProperty("Enabled", True)
'    m_IPAddress = PropBag.ReadProperty("IPAddress", m_def_IPAddress)
    txtRoundTripTime.Text = PropBag.ReadProperty("RoundTripTime", "")
    m_IPAddress = PropBag.ReadProperty("IPAddress", m_def_IPAddress)
    txtIP(0).Text = PropBag.ReadProperty("Octet1", "")
    txtIP(1).Text = PropBag.ReadProperty("Octet2", "")
    txtIP(2).Text = PropBag.ReadProperty("Octet3", "")
    txtIP(3).Text = PropBag.ReadProperty("Octet4", "")
    m_Octet1 = PropBag.ReadProperty("Octet1", m_def_Octet1)
    m_Octet2 = PropBag.ReadProperty("Octet2", m_def_Octet2)
    m_Octet3 = PropBag.ReadProperty("Octet3", m_def_Octet3)
    m_Octet4 = PropBag.ReadProperty("Octet4", m_def_Octet4)
    txtTimeout.Text = PropBag.ReadProperty("PING_TIMEOUT", "3000")
End Sub

Private Sub UserControl_Resize()
    Height = 465
    Width = 1245
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", cmdPing.Enabled, True)
'    Call PropBag.WriteProperty("IPAddress", m_IPAddress, m_def_IPAddress)
    Call PropBag.WriteProperty("RoundTripTime", txtRoundTripTime.Text, "")
    Call PropBag.WriteProperty("IPAddress", m_IPAddress, m_def_IPAddress)
    Call PropBag.WriteProperty("Octet1", txtIP(0).Text, "")
    Call PropBag.WriteProperty("Octet2", txtIP(1).Text, "")
    Call PropBag.WriteProperty("Octet3", txtIP(2).Text, "")
    Call PropBag.WriteProperty("Octet4", txtIP(3).Text, "")
    Call PropBag.WriteProperty("Octet1", m_Octet1, m_def_Octet1)
    Call PropBag.WriteProperty("Octet2", m_Octet2, m_def_Octet2)
    Call PropBag.WriteProperty("Octet3", m_Octet3, m_def_Octet3)
    Call PropBag.WriteProperty("Octet4", m_Octet4, m_def_Octet4)
    Call PropBag.WriteProperty("PING_TIMEOUT", txtTimeout.Text, "3000")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,127.0.0.1
Public Property Get IPAddress() As String
Attribute IPAddress.VB_Description = "IP Address of Host to ping in the form xxx.xxx.xxx.xxx"
Attribute IPAddress.VB_ProcData.VB_Invoke_Property = "IPProperties"
    IPAddress = m_IPAddress
End Property

Public Property Let IPAddress(ByVal New_IPAddress As String)
    m_IPAddress = New_IPAddress
    PropertyChanged "IPAddress"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtIP(0),txtIP,0,Text
'Public Property Get Octet1() As String
'    Octet1 = txtIP(0).Text
'End Property
'
'Public Property Let Octet1(ByVal New_Octet1 As String)
'    txtIP(0).Text() = New_Octet1
'    PropertyChanged "Octet1"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtIP(1),txtIP,1,Text
'Public Property Get Octet2() As String
'    Octet2 = txtIP(1).Text
'End Property
'
'Public Property Let Octet2(ByVal New_Octet2 As String)
'    txtIP(1).Text() = New_Octet2
'    PropertyChanged "Octet2"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtIP(2),txtIP,2,Text
'Public Property Get Octet3() As String
'    Octet3 = txtIP(2).Text
'End Property
'
'Public Property Let Octet3(ByVal New_Octet3 As String)
'    txtIP(2).Text() = New_Octet3
'    PropertyChanged "Octet3"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtIP(3),txtIP,3,Text
'Public Property Get Octet4() As String
'    Octet4 = txtIP(3).Text
'End Property
'
'Public Property Let Octet4(ByVal New_Octet4 As String)
'    txtIP(3).Text() = New_Octet4
'    PropertyChanged "Octet4"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Octet1() As String
Attribute Octet1.VB_Description = "Returns/sets the text contained in the control."
Attribute Octet1.VB_ProcData.VB_Invoke_Property = "IPProperties"
    Octet1 = m_Octet1
End Property

Public Property Let Octet1(ByVal New_Octet1 As String)
    m_Octet1 = New_Octet1
    PropertyChanged "Octet1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Octet2() As String
Attribute Octet2.VB_Description = "Returns/sets the text contained in the control."
Attribute Octet2.VB_ProcData.VB_Invoke_Property = "IPProperties"
    Octet2 = m_Octet2
End Property

Public Property Let Octet2(ByVal New_Octet2 As String)
    m_Octet2 = New_Octet2
    PropertyChanged "Octet2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Octet3() As String
Attribute Octet3.VB_Description = "Returns/sets the text contained in the control."
Attribute Octet3.VB_ProcData.VB_Invoke_Property = "IPProperties"
    Octet3 = m_Octet3
End Property

Public Property Let Octet3(ByVal New_Octet3 As String)
    m_Octet3 = New_Octet3
    PropertyChanged "Octet3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Octet4() As String
Attribute Octet4.VB_Description = "Returns/sets the text contained in the control."
Attribute Octet4.VB_ProcData.VB_Invoke_Property = "IPProperties"
    Octet4 = m_Octet4
End Property

Public Property Let Octet4(ByVal New_Octet4 As String)
    m_Octet4 = New_Octet4
    PropertyChanged "Octet4"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtTimeout,txtTimeout,-1,Text
Public Property Get PING_TIMEOUT() As String
Attribute PING_TIMEOUT.VB_Description = "Returns/sets the text contained in the control."
Attribute PING_TIMEOUT.VB_ProcData.VB_Invoke_Property = "IPProperties"
    PING_TIMEOUT = txtTimeout.Text
End Property

Public Property Let PING_TIMEOUT(ByVal New_PING_TIMEOUT As String)
    txtTimeout.Text() = New_PING_TIMEOUT
    PropertyChanged "PING_TIMEOUT"
End Property

