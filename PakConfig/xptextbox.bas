Attribute VB_Name = "xptextbox"
Public Enum states
    Normal = 0
    Disable = 1
    ReadOnly = 2
End Enum

Public Function xptxt(txt As TextBox, pic As PictureBox, BackColor As ColorConstants, State As states)
'Make sure you have a border of 6 pixels around
'textbox before you use this function

'states = normal, disable, read only
    pic.BackColor = BackColor
    pic.ScaleMode = 3
    txt.Appearance = 0
    txt.BorderStyle = 0
    pic.AutoRedraw = True
    Dim a, b, c, d As Integer
    a = txt.top + txt.Height + 1
    b = txt.left - 1
    c = txt.left + txt.Width + 1
    d = txt.top - 1
    pic.DrawWidth = 1
    pic.Line (b, a)-(c, a), RGB(127, 157, 185)
    pic.Line (c, a)-(c, d), RGB(127, 157, 185)
    pic.Line (c, d)-(b, d), RGB(127, 157, 185)
    pic.Line (b, d)-(b, a), RGB(127, 157, 185)
    
    If State = Normal Then
        txt.BackColor = vbWhite
        txt.Enabled = True
        txt.Locked = False
    ElseIf State = Disable Then
        txt.Enabled = False
        txt.BackColor = RGB(235, 235, 228)
        txt.ForeColor = RGB(161, 161, 146)
    ElseIf State = ReadOnly Then
        txt.Enabled = True
        txt.Locked = True
    End If
End Function
