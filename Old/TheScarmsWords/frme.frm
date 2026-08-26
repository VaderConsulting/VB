VERSION 5.00
Begin VB.Form frmE 
   BorderStyle     =   0  'None
   ClientHeight    =   1485
   ClientLeft      =   4365
   ClientTop       =   2205
   ClientWidth     =   1560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Call pQuit
End Sub

Private Sub Form_Load()
Dim hRgn     As Long
Dim lH5      As Long
Dim lW2      As Long
ReDim XY(12) As POINTAPI

With Me
    .ScaleMode = vbPixels
    .Width = cWidth
    .Height = cHeight
    
    lH5 = .ScaleHeight / 5
    lW2 = .ScaleWidth / 2
    
    XY(0).X = 0
    XY(0).Y = 0
    XY(1).X = .ScaleWidth
    XY(1).Y = 0
    XY(2).X = .ScaleWidth
    XY(2).Y = lH5
    XY(3).X = .ScaleWidth - lW2
    XY(3).Y = lH5
    XY(4).X = .ScaleWidth - lW2
    XY(4).Y = 2 * lH5
    XY(5).X = .ScaleWidth
    XY(5).Y = 2 * lH5
    XY(6).X = .ScaleWidth
    XY(6).Y = 3 * lH5
    XY(7).X = .ScaleWidth - lW2
    XY(7).Y = 3 * lH5
    XY(8).X = .ScaleWidth - lW2
    XY(8).Y = 4 * lH5
    XY(9).X = .ScaleWidth
    XY(9).Y = 4 * lH5
    XY(10).X = .ScaleWidth
    XY(10).Y = .ScaleHeight
    XY(11).X = 0
    XY(11).Y = .ScaleHeight
End With

hRgn = CreatePolygonRgn(XY(0), 12, 2)
Call SetWindowRgn(Me.hWnd, hRgn, True)
Call pGradient(Me, "Red")
End Sub

