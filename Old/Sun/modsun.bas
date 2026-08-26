Attribute VB_Name = "Module1"
    Public MeanAnomaly As Double
    Public Eccentricity As Double
    Public JulianDay As Double
    Public EccentricAnomaly As Double
    Public TrueAnomalyXCoord As Double
    Public TrueAnomalyYCoord As Double
    Public vDay
    Public vMonth
    Public vYear
    Public Const PI = 3.14159265358979

Sub main()
    frmSun.Show
End Sub

Public Function RadtoDeg(x As Double)
    x = x * (180 / PI)
End Function

Public Function DegtoRad(x As Double)
    x = x * (PI / 180)
End Function

Public Function SinD(x As Double)
    x = RadtoDeg(Sin(x))
End Function

Public Function CosD(x As Double)
    x = RadtoDeg(Cos(x))
End Function

Public Function TanD(x As Double)
    x = RadtoDeg(Tan(x))
End Function


