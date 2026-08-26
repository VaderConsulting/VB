VERSION 5.00
Begin VB.Form frmC 
   BorderStyle     =   0  'None
   ClientHeight    =   1485
   ClientLeft      =   7515
   ClientTop       =   2445
   ClientWidth     =   1560
   ClipControls    =   0   'False
   Icon            =   "frmC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmC"
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
ReDim XY(8) As POINTAPI

With Me
    .ScaleMode = vbPixels
    .Width = cWidth
    .Height = cHeight
    
    lH3 = .ScaleHeight / 3
    
    XY(0).X = 0
    XY(0).Y = 0
    XY(1).X = .ScaleWidth
    XY(1).Y = 0
    XY(2).X = .ScaleWidth
    XY(2).Y = lH3
    XY(3).X = .ScaleWidth / 3
    XY(3).Y = lH3
    XY(4).X = .ScaleWidth / 3
    XY(4).Y = 2 * lH3
    XY(5).X = .ScaleWidth
    XY(5).Y = 2 * lH3
    XY(6).X = .ScaleWidth
    XY(6).Y = .ScaleHeight
    XY(7).X = 0
    XY(7).Y = .ScaleHeight
End With

hRgn = CreatePolygonRgn(XY(0), 8, 2)
Call SetWindowRgn(Me.hWnd, hRgn, True)
Call pGradient(Me, "Yellow")
End Sub


