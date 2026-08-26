VERSION 5.00
Begin VB.Form frmM 
   BorderStyle     =   0  'None
   ClientHeight    =   1605
   ClientLeft      =   8430
   ClientTop       =   2640
   ClientWidth     =   1560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Call pQuit
End Sub

Private Sub Form_Load()
Dim hRgn    As Long
Dim lH3     As Long
Dim lW3     As Long
ReDim XY(12) As POINTAPI

With Me
    .ScaleMode = vbPixels
    .Width = cWidth
    .Height = cHeight

    lH3 = .ScaleHeight / 3
    lW3 = .ScaleWidth / 3
    
    XY(0).X = 0
    XY(0).Y = 0
    XY(1).X = lW3
    XY(1).Y = 0
    XY(2).X = .ScaleWidth / 2
    XY(2).Y = .ScaleHeight / 3
    XY(3).X = 2 * lW3
    XY(3).Y = 0
    XY(4).X = .ScaleWidth
    XY(4).Y = 0
    XY(5).X = .ScaleWidth
    XY(5).Y = .ScaleHeight
    XY(6).X = 2 * lW3
    XY(6).Y = .ScaleHeight
    XY(7).X = 2 * lW3
    XY(7).Y = .ScaleHeight / 2
    XY(8).X = .ScaleWidth / 2
    XY(8).Y = 2 * (.ScaleHeight / 3)
    XY(9).X = lW3
    XY(9).Y = .ScaleHeight / 2
    XY(10).X = lW3
    XY(10).Y = .ScaleHeight
    XY(11).X = 0
    XY(11).Y = .ScaleHeight
End With

hRgn = CreatePolygonRgn(XY(0), 12, 2)
Call SetWindowRgn(Me.hWnd, hRgn, True)
Call pGradient(Me, "Red")

End Sub


