Attribute VB_Name = "mod_gradient"
Public Function Horizontal(Newform As Form, Colour1 As ColorConstants, Colour2 As ColorConstants)
    'This function sets gradient for the form
    Dim VR, VG, VB As Single
    Dim Color1, Color2 As Long
    Dim R, G, b, R2, G2, B2 As Integer
    Dim temp As Long

    Color1 = Colour1
    Color2 = Colour2

    temp = (Color1 And 255)
    R = temp And 255
    temp = Int(Color1 / 256)
    G = temp And 255
    temp = Int(Color1 / 65536)
    b = temp And 255
    temp = (Color2 And 255)
    R2 = temp And 255
    temp = Int(Color2 / 256)
    G2 = temp And 255
    temp = Int(Color2 / 65536)
    B2 = temp And 255

    VR = Abs(R - R2) / Newform.ScaleWidth
    VG = Abs(G - G2) / Newform.ScaleWidth
    VB = Abs(b - B2) / Newform.ScaleWidth

    If R2 < R Then VR = -VR
    If G2 < G Then VG = -VG
    If B2 < b Then VB = -VB

    For x = 0 To Newform.ScaleWidth
        R2 = R + VR * x
        G2 = G + VG * x
        B2 = b + VB * x
        Newform.Line (x, 0)-(x, Newform.ScaleHeight), RGB(R2, G2, B2)
    Next x
End Function
