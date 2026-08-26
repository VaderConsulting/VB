VERSION 5.00
Begin VB.Form frmA 
   BorderStyle     =   0  'None
   ClientHeight    =   1485
   ClientLeft      =   8100
   ClientTop       =   2100
   ClientWidth     =   1560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmA"
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
ReDim XY(6) As POINTAPI

With Me
    .ScaleMode = vbPixels
    .Width = cWidth
    .Height = cHeight

    lH3 = .ScaleHeight / 3
    lW3 = .ScaleWidth / 3

    XY(0).X = .ScaleWidth / 2
    XY(0).Y = 0
    XY(1).X = .ScaleWidth
    XY(1).Y = .ScaleHeight
    XY(2).X = .ScaleWidth - lW3
    XY(2).Y = .ScaleHeight
    XY(3).X = .ScaleWidth / 2
    XY(3).Y = .ScaleHeight - lH3
    XY(4).X = lW3
    XY(4).Y = .ScaleHeight
    XY(5).X = 0
    XY(5).Y = .ScaleHeight
End With

hRgn = CreatePolygonRgn(XY(0), 6, 2)
Call SetWindowRgn(Me.hWnd, hRgn, True)
Call pGradient(Me, "Red")
End Sub


