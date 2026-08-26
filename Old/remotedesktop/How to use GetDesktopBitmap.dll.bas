Attribute VB_Name = "MgetDesktop"
Public Declare Function getDesktop Lib "GetDesktopBitmap.dll" (ByVal nWidth As Integer, ByVal nHeight As Integer, blnJpeg As Boolean, ByVal JPGCompressQuality As Integer, ByVal strFileName As String) As Integer
Sub Main()
    ' if nWidth = 0 then the size of the saved picture is equal with the size of the screen.
    getDesktop 0, 600, True, 60, "C:\test.jpg"
End Sub
