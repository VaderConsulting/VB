Attribute VB_Name = "mdlIOleObject"
'*********************************************************************************************
'
' Adding items to UserControl context menu
'
' IOleObject functions
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 02/09/2000
'
'*********************************************************************************************
Option Explicit

Public Function IOleObject_GetExtent(ByVal This As IOleObject, Aspect As DVASPECT, Size As Size) As Long
Dim PixelSize As Size

   ' Size contains the ideal control
   ' size in HIMETRIC. Changing it
   ' will change the control size.
   
   ' Convert from HIMETRIC to pixels
   ' (each HIMETRIC = 72/127 Twips):
   ' Pixels = HIMETRIC * (72/127) / Screen.TwipsPerPixel
   ' HIMETRIC = Pixels * Screen.TwipsPerPixel / (72/127)
   
   PixelSize.cx = MulDiv(Size.cx, 72, 127) / Screen.TwipsPerPixelX
   PixelSize.cy = MulDiv(Size.cy, 72, 127) / Screen.TwipsPerPixelY
   
   ' 64x64 is the minimun control size
   ' and the control resizes is steps
   ' of 16 pixels
   
   If PixelSize.cx < 64 Then
      Size.cx = MulDiv(64 * Screen.TwipsPerPixelX, 127, 72)
   Else
      Size.cx = MulDiv(Int(PixelSize.cx / 16) * 16 * Screen.TwipsPerPixelX, 127, 72)
   End If
   
   If PixelSize.cy < 64 Then
      Size.cy = MulDiv(64 * Screen.TwipsPerPixelY, 127, 72)
   Else
      Size.cy = MulDiv(Int(PixelSize.cy / 16) * 16 * Screen.TwipsPerPixelY, 127, 72)
   End If

End Function


