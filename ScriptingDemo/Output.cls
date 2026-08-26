VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Abstraction for a textbox to provide an output window
'for scripts run by ScriptingDemo.vbp
Private m_Text As TextBox
Friend Property Set OutputBox(data As TextBox)
    Set m_Text = data
    
End Property
Public Sub PrintLine(strData As Variant)
    m_Text.Text = m_Text.Text & strData & vbCrLf
    
End Sub

Public Sub Clear()
    m_Text.Text = ""
End Sub
