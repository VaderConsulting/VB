Attribute VB_Name = "Module1"
Option Explicit

Public Const cHeight = 1370
Public Const cWidth = 1200

Public glTwip As Long

Public Type POINTAPI
   X As Long
   Y As Long
End Type

Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


Public Sub pGradient(frm As Form, ByVal sColor As String)
Dim i As Integer
Dim W As Long

With frm
    .AutoRedraw = True
    .DrawStyle = vbInsideSolid
    .DrawMode = vbCopyPen
    .ScaleMode = vbPixels
    .DrawWidth = 2
    .ScaleHeight = 256
End With
W = Screen.Width

Select Case sColor
    Case "Red"
        For i = 0 To 255 Step 3
            frm.Line (0, i)-(W, i - 1), RGB(255 - i, 0, 0), B
        Next
    Case "Blue"
        For i = 0 To 255 Step 3
            frm.Line (0, i)-(W, i - 1), RGB(0, 0, 255 - i), B
        Next
    Case "Green"
        For i = 0 To 255 Step 3
            frm.Line (0, i)-(W, i - 1), RGB(0, 255 - i, 0), B
        Next
    Case "Yellow"
        For i = 0 To 255 Step 3
            frm.Line (0, i)-(W, i - 1), RGB(i, i, 255 - i), B
        Next
End Select
End Sub

Public Sub pQuit()
Unload frmT
Unload frmH
Unload frmE
Unload frmS1
Unload frmA
Unload frmC
Unload frmR
Unload frmM
Unload frmS2
End Sub
