Attribute VB_Name = "Module1"
Function FindAngle(TXX, Tyy, PAngle)
 rxx = Cos(PAngle) * TXX - Sin(PAngle) * Tyy
 ryy = Sin(PAngle) * TXX + Cos(PAngle) * Tyy
 XX = rxx
 yy = ryy
 pi = (22 / 7) * 18.05
 If yy = 0 Then
  If XX > 0 Then
   Angle = -90 / pi
  Else
   Angle = -270 / pi
  End If
 Else
  Angle = Atn(XX / yy)
 End If
 Angle = -Angle * pi
 If yy > 0 Then Angle = Angle + 180
 If 0 > yy And XX < 0 Then Angle = Angle + 360
 If XX < 0 Then Angle = -360 + Angle
 Angle = Int(Angle)
 FindAngle = -Angle
End Function


Function Distance(x1, y1, x2, y2) As Integer
 x = Abs(x2 - x1)
 y = Abs(y2 - y1)
 Distance = Sqr(x * 2 + y * 2)
End Function
