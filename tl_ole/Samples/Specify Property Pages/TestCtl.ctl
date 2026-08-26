VERSION 5.00
Begin VB.UserControl ISPPSample 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "TestCtl.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdClr 
      Caption         =   "&Colors"
      Height          =   300
      Left            =   1845
      TabIndex        =   3
      Top             =   3285
      Width           =   1110
   End
   Begin VB.CommandButton cmdPict 
      Caption         =   "&Picture"
      Height          =   300
      Left            =   3690
      TabIndex        =   2
      Top             =   0
      Width           =   1110
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "&Font"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1110
   End
   Begin VB.Label lblTxt 
      AutoSize        =   -1  'True
      Caption         =   "Click to show my property pages"
      Height          =   195
      Left            =   1350
      TabIndex        =   0
      Top             =   1650
      Width           =   2280
   End
End
Attribute VB_Name = "ISPPSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (A As Any, b As Any, ByVal c As Long)

Dim WithEvents m_Text As Caption
Attribute m_Text.VB_VarHelpID = -1
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "ISPPSamplePage"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTxt,lblTxt,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblTxt.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    
    Set lblTxt.Font = New_Font
    PropertyChanged "Font"
    
    UserControl_Resize
    
End Property

Private Sub cmdClr_Click()
Dim ClrPage As UUID

    CLSIDFromString IIDSTR_StandardColor, ClrPage
        
    OleCreatePropertyFrame UserControl.hWnd, 0, 0, Ambient.DisplayName & " Color", 1, Me, 1, ClrPage, &H40D, 0, 0

End Sub

Private Sub cmdFont_Click()
Dim FntPage As UUID

    CLSIDFromString IIDSTR_StandardFont, FntPage
        
    OleCreatePropertyFrame UserControl.hWnd, 0, 0, Ambient.DisplayName & " Font", 1, Me, 1, FntPage, &H40D, 0, 0

End Sub


Private Sub cmdPict_Click()
Dim FntPage As UUID

    CLSIDFromString IIDSTR_StandardPicture, FntPage
        
    OleCreatePropertyFrame UserControl.hWnd, 0, 0, Ambient.DisplayName & " Picture", 1, Me, 1, FntPage, &H40D, 0, 0

End Sub


Private Sub lblTxt_Change()

    UserControl_Resize
    
End Sub

Private Sub lblTxt_Click()
Dim ClrPage As UUID

    CLSIDFromString IIDSTR_StandardColor, ClrPage
        
    OleCreatePropertyFrame UserControl.hWnd, 0, 0, "Caption Object", 1, m_Text, 1, ClrPage, &H40D, 0, 0

End Sub

Private Sub m_Text_PropertyChanged(ByVal PropName As String)

    Select Case PropName
        Case "Text"
            lblTxt.Caption = m_Text.Text
        Case "ForeColor"
            lblTxt.ForeColor = m_Text.ForeColor
        Case "BackColor"
            lblTxt.BackColor = m_Text.BackColor
    End Select
    
End Sub

Private Sub UserControl_Click()
Dim ISPP As ISpecifyPropertyPages
Dim Pages As CAUUID

    Set ISPP = Me
    
    ISPP.GetPages Pages

    OleCreatePropertyFrame Parent.hWnd, 0, 0, "axISPP Sample Control", 1, Me, Pages.cElems, ByVal Pages.pElems, Ambient.LocaleID, 0, 0
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Private Sub UserControl_Initialize()
    
    Set m_Text = New Caption

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    Set Font = Ambient.Font
    
    m_Text.Text = "Click here to chage text colors"
    m_Text.BackColor = vb3DFace
    m_Text.ForeColor = vbButtonText
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    
    m_Text.Text = PropBag.ReadProperty("Text")
    m_Text.BackColor = PropBag.ReadProperty("BackColor1", vb3DFace)
    m_Text.ForeColor = PropBag.ReadProperty("ForeColor1", vbButtonText)

End Sub

Private Sub UserControl_Resize()

    lblTxt.Move (ScaleWidth - lblTxt.Width) / 2, (ScaleHeight - lblTxt.Height) / 2
    cmdPict.Move ScaleWidth - cmdPict.Width
    cmdClr.Move (ScaleWidth - cmdClr.Width) / 2, ScaleHeight - cmdClr.Height
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    
    Call PropBag.WriteProperty("Text", m_Text.Text)
    Call PropBag.WriteProperty("BackColor1", m_Text.BackColor)
    Call PropBag.WriteProperty("ForeColor1", m_Text.ForeColor)

End Sub

Public Property Get Text() As Object
Attribute Text.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Text.VB_ProcData.VB_Invoke_Property = "ISPPSamplePage"
    Set Text = m_Text
End Property

