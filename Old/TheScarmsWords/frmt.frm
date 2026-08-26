VERSION 5.00
Begin VB.Form frmT 
   BorderStyle     =   0  'None
   ClientHeight    =   1065
   ClientLeft      =   6120
   ClientTop       =   1965
   ClientWidth     =   1560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmT"
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
Dim lW3     As Long
ReDim XY(7) As POINTAPI

glTwip = Screen.TwipsPerPixelX
With Me
    .ScaleMode = vbPixels
    .Width = cWidth
    .Height = cHeight
    
    frmH.Show
    frmE.Show
    frmS1.Show
    frmC.Show
    frmA.Show
    frmR.Show
    frmM.Show
    frmS2.Show
    
    
    frmH.Top = .Top
    frmE.Top = .Top
    frmH.Left = frmT.Left + cWidth + glTwip
    frmE.Left = frmH.Left + cWidth + glTwip
    
    frmS1.Top = .Top + .Height + glTwip
    frmC.Top = frmS1.Top
    frmA.Top = frmS1.Top
    frmR.Top = frmS1.Top
    frmM.Top = frmS1.Top
    frmS2.Top = frmS1.Top
    
    frmS1.Left = frmT.Left - (1.5 * cWidth)
    frmC.Left = frmS1.Left + cWidth + glTwip
    frmA.Left = frmC.Left + cWidth + glTwip
    frmR.Left = frmA.Left + cWidth + glTwip
    frmM.Left = frmR.Left + cWidth + glTwip
    frmS2.Left = frmM.Left + cWidth + glTwip
   
    .ScaleMode = vbPixels
    XY(0).X = 0
    XY(0).Y = 0
    XY(1).X = .ScaleWidth
    XY(1).Y = 0
    XY(2).X = .ScaleWidth
    XY(2).Y = .ScaleHeight / 3
    XY(3).X = .ScaleWidth - (.ScaleWidth / 3)
    XY(3).Y = .ScaleHeight / 3
    XY(4).X = .ScaleWidth - (.ScaleWidth / 3)
    XY(4).Y = .ScaleHeight
    XY(5).X = .ScaleWidth / 3
    XY(5).Y = .ScaleHeight
    XY(6).X = .ScaleWidth / 3
    XY(6).Y = .ScaleHeight / 3
    XY(7).X = 0
    XY(7).Y = .ScaleHeight / 3
End With

hRgn = CreatePolygonRgn(XY(0), 8, 2)
Call SetWindowRgn(Me.hWnd, hRgn, True)
Call pGradient(Me, "Yellow")
End Sub

Private Sub Timer1_Timer()
Static iCount As Integer

Select Case iCount
    Case 0
        Call pGradient(frmT, "Red")
        Call pGradient(frmH, "Blue")
        Call pGradient(frmE, "Green")
        Call pGradient(frmS1, "Yellow")
        Call pGradient(frmC, "Red")
        Call pGradient(frmA, "Blue")
        Call pGradient(frmM, "Green")
        Call pGradient(frmR, "Yellow")
        Call pGradient(frmS2, "Red")
    Case 1
        Call pGradient(frmT, "Blue")
        Call pGradient(frmH, "Green")
        Call pGradient(frmE, "Yellow")
        Call pGradient(frmS1, "Red")
        Call pGradient(frmC, "Blue")
        Call pGradient(frmA, "Green")
        Call pGradient(frmM, "Yellow")
        Call pGradient(frmR, "Red")
        Call pGradient(frmS2, "Blue")
    Case 2
        Call pGradient(frmT, "Green")
        Call pGradient(frmH, "Yellow")
        Call pGradient(frmE, "Red")
        Call pGradient(frmS1, "Blue")
        Call pGradient(frmC, "Green")
        Call pGradient(frmA, "Yellow")
        Call pGradient(frmM, "Red")
        Call pGradient(frmR, "Blue")
        Call pGradient(frmS2, "Green")
    Case 3
        Call pGradient(frmT, "Yellow")
        Call pGradient(frmH, "Red")
        Call pGradient(frmE, "Blue")
        Call pGradient(frmS1, "Green")
        Call pGradient(frmC, "Yellow")
        Call pGradient(frmA, "Red")
        Call pGradient(frmM, "Blue")
        Call pGradient(frmR, "Green")
        Call pGradient(frmS2, "Yellow")
End Select

iCount = iCount + 1
If iCount > 3 Then iCount = 0
DoEvents
End Sub


