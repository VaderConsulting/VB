Attribute VB_Name = "trans_colour"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Dim CurRgn, TempRgn As Long  ' Region variables

Public Function AutoFormShape(bg As Form, transColor)
    Dim X, Y As Integer
    CurRgn = CreateRectRgn(0, 0, bg.ScaleWidth, bg.ScaleHeight)  ' Create base region which is the current whole window

    While Y <= 5 ' Go through each column of pixels on form
        While X <= bg.ScaleWidth  ' Go through each line of pixels on form
            If GetPixel(bg.hdc, X, Y) = transColor Then  ' If the pixels color is the transparency color (bright purple is a good one to use)
                TempRgn = CreateRectRgn(X, Y, X + 1, Y + 1)  ' Create a temporary pixel region for this pixel
                success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)  ' Combine temp pixel region with base region using RGN_DIFF to extract the pixel and make it transparent
                DeleteObject (TempRgn)  ' Delete the temporary pixel region and clear up very important resources
            End If
            X = X + 1
        Wend
            Y = Y + 1
            X = 0
    Wend
    success = SetWindowRgn(bg.hWnd, CurRgn, True)  ' Finally set the windows region to the final product
    DeleteObject (CurRgn)  ' Delete the now un-needed base region and free resources
    ' This code is by Chris Yates (1999)
End Function
