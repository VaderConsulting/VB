VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl cbbBarcode 
   BackStyle       =   0  'Transparent
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "Barcode.ctx":0000
   ScaleHeight     =   645
   ScaleWidth      =   2295
   ToolboxBitmap   =   "Barcode.ctx":0017
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1792
      Top             =   0
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1152
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   16384
      NullDiscard     =   -1  'True
      RTSEnable       =   -1  'True
      InputMode       =   1
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   384
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   768
   End
   Begin VB.Image imgPicture 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   0
      Picture         =   "Barcode.ctx":0111
      Stretch         =   -1  'True
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "cbbBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Event Declarations:
Event Scan() 'MappingInfo=MSComm1,MSComm1,-1,OnComm
Event ValidScan() 'MappingInfo=txtCode,txtCode,-1,Change
Attribute ValidScan.VB_Description = "Occurs when the contents of a control have changed."

Private Sub Timer1_Timer()
    If MSComm1.PortOpen = False Then Exit Sub
    'If MSComm1.Input = "" Then Exit Sub
    Dim strBarcode As String
    Static strBuffer As String
    strBuffer = strBuffer & MSComm1.Input
    If strBuffer <> "" Then
        Debug.Print "Buffer: '" & strBuffer & "'"
    End If
    If CBool(InStr(1, strBuffer, vbCr, vbTextCompare)) Then
        strBarcode = Mid(strBuffer, 1, InStr(1, strBuffer, vbCr, vbTextCompare) - 1)
        strBuffer = Mid(strBuffer, InStr(1, strBuffer, vbCr, vbTextCompare) + 1)
    End If
    If strBarcode <> "" Then
        Debug.Print "Barcode: " & strBarcode
        txtCode.Text = strBarcode
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.Width = imgPicture.Width
    Height = imgPicture.Height
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = imgPicture.Width
    Height = imgPicture.Height
End Sub
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtCode,txtCode,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = txtCode.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    txtCode.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtCode,txtCode,-1,ForeColor
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = txtCode.ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    txtCode.ForeColor() = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "Custom"
    Enabled = Timer1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Timer1.Enabled() = New_Enabled
    MSComm1.PortOpen = New_Enabled
    PropertyChanged "Enabled"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtCode,txtCode,-1,Font
'Public Property Get Font() As Font
'    Set Font = txtCode.Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Set txtCode.Font = New_Font
'    PropertyChanged "Font"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtCode,txtCode,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'    BorderStyle = txtCode.BorderStyle
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtCode,txtCode,-1,Refresh
'Public Sub Refresh()
'    txtCode.Refresh
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get Delay() As Long
Attribute Delay.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
Attribute Delay.VB_ProcData.VB_Invoke_Property = "Custom"
    Delay = Timer1.Interval
End Property

Public Property Let Delay(ByVal New_Delay As Long)
    Timer1.Interval() = New_Delay
    PropertyChanged "Delay"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtCode,txtCode,-1,Text
Public Property Get Code() As String
Attribute Code.VB_Description = "Returns/sets the text contained in the control."
Attribute Code.VB_ProcData.VB_Invoke_Property = "Custom"
    Code = txtCode.Text
End Property

Public Property Let Code(ByVal New_Code As String)
    txtCode.Text() = New_Code
    PropertyChanged "Code"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSComm1,MSComm1,-1,Settings
Public Property Get Settings() As String
Attribute Settings.VB_Description = "Sets/returns the baud rate, parity, data bit, and stop bit parameters."
Attribute Settings.VB_ProcData.VB_Invoke_Property = "Custom"
    Settings = MSComm1.Settings
End Property

Public Property Let Settings(ByVal New_Settings As String)
    MSComm1.Settings() = New_Settings
    PropertyChanged "Settings"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(propbag As PropertyBag)
    txtCode.BackColor = propbag.ReadProperty("BackColor", &H80000005)
    txtCode.ForeColor = propbag.ReadProperty("ForeColor", &H80000008)
    Timer1.Enabled = propbag.ReadProperty("Enabled", False)
    Set Font = propbag.ReadProperty("Font", Ambient.Font)
    Timer1.Interval = propbag.ReadProperty("Delay", 500)
    txtCode.Text = propbag.ReadProperty("Code", "")
    MSComm1.Settings = propbag.ReadProperty("Settings", "9600,n,8,1")
    MSComm1.CommPort = propbag.ReadProperty("CommPort", 1)
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(propbag As PropertyBag)

    Call propbag.WriteProperty("BackColor", txtCode.BackColor, &H80000005)
    Call propbag.WriteProperty("ForeColor", txtCode.ForeColor, &H80000008)
    Call propbag.WriteProperty("Enabled", Timer1.Enabled, True)
    Call propbag.WriteProperty("Font", Font, Ambient.Font)
    Call propbag.WriteProperty("Delay", Timer1.Interval, 500)
    Call propbag.WriteProperty("Code", txtCode.Text, "")
    Call propbag.WriteProperty("Settings", MSComm1.Settings, "9600,n,8,1")
    Call propbag.WriteProperty("CommPort", MSComm1.CommPort, 1)
End Sub
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=MSComm1,MSComm1,-1,CommPort
''Public Property Get Port() As Integer
''    Port = MSComm1.CommPort
''End Property
''
''Public Property Let Port(ByVal New_Port As Integer)
''    MSComm1.CommPort() = New_Port
''    PropertyChanged "Port"
''End Property
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!

Private Sub txtCode_Change()
    RaiseEvent ValidScan
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSComm1,MSComm1,-1,CommPort
Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "Returns/sets the communications port number."
Attribute CommPort.VB_ProcData.VB_Invoke_Property = "Custom"
    CommPort = MSComm1.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    MSComm1.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

Private Sub MSComm1_OnComm()
    'Debug.Print "Data: " & MSComm1.Input
    'RaiseEvent Scan
    
End Sub

