VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cThumbnail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ************************************************************************
'  The EliteVB thumbnail class
' ************************************************************************
'
'      Written By: Garrett D. Sever aka The Hand
'      Written On: 12/20/02
'      Updated On: 01 July 2003 By D Robinson aka Dangerous
'
'     Description: A simple extension of the cDIBSection written by
'                  Steve McMahon on vbAccelerator to create thumbnail images
'                  from a larger image.
'
'    Terms of use: This module and source code can be re-distributed however
'                  you want and wherever you want. All we ask is that you
'                  give credit to EliteVB if the code is published elsewhere.
'
'    Converted into DLL for ASP web site usage - specifically for WoLs LAN
'
' ********************************************************************************
'      Visit http://www.EliteVB.com for more high-powered solutions!!
' ********************************************************************************

Option Explicit

' **********************************************************************************
'  API DECLARATIONS
' **********************************************************************************

Private Type BITMAP
    bmType          As Long
    bmWidth         As Long
    bmHeight        As Long
    bmWidthBytes    As Long
    bmPlanes        As Integer
    bmBitsPixel     As Integer
    bmBits          As Long
End Type

Private Type RECT
    left            As Long
    top             As Long
    right           As Long
    bottom          As Long
End Type

Private Type GUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(7)        As Byte
End Type
Private Type PicBmp
    Size            As Long
    Type            As Long
    hBmp            As Long
    hPal            As Long
    Reserved        As Long
End Type

' API declarations
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CopyImage Lib "USER32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Private Declare Function ReleaseDC Lib "USER32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetRect Lib "USER32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

' **********************************************************************************
'  Class property variables
' **********************************************************************************

Private mMaxWidth       As Long
Private mMaxHeight      As Long
Private mBackColor      As Long
Private mDestFilePath   As String
Private mForcedDim      As Boolean
Public Property Let ForcedDimensions(aBool As Boolean)
    mForcedDim = aBool
End Property
Public Property Get ForcedDimensions() As Boolean
    ForcedDimensions = mForcedDim
End Property
Public Property Let MaxWidth(aLong As Long)
    mMaxWidth = aLong
End Property
Public Property Get MaxWidth() As Long
    MaxWidth = mMaxWidth
End Property
Public Property Let MaxHeight(aLong As Long)
    mMaxHeight = aLong
End Property
Public Property Get MaxHeight() As Long
    MaxHeight = mMaxHeight
End Property
Public Property Let BackColor(aColor As OLE_COLOR)
    mBackColor = aColor
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property
Public Property Let DestFilePath(aString As String)
    mDestFilePath = aString
End Property
Public Property Get DestFilePath() As String
    DestFilePath = mDestFilePath
End Property

Public Sub CreateFromStdPicture(aStdPicture As StdPicture)

    Dim lDestDIB        As cDIBSection  ' DIB section - from INTEL - ITS MAGIC!!! lol... not
    Dim lDestStdPic     As StdPicture   ' Standard picture
    Dim lDestDC         As Long         ' Device context used to fill the DIB class
    Dim lDestBmp        As Long         ' Bitmap used to transfer the image to the DIBSection
    
    Dim lBackBrush      As Long         ' a brush to fill in the background color
    Dim lBackColor      As Long         ' a background color (RGB long)
    Dim lRECT           As RECT         ' a rectangle used to fill in the bg color
    
    Dim lTempBmp        As Long         ' handle to the resized bitmap image
    Dim lPicDC          As Long         ' device context to hold the resized bitmap

    Dim lDestX          As Long         ' thumbnail image x destination, pixels
    Dim lDestY          As Long         ' thumbnail image y destination, pixels
    Dim lDestWid        As Long         ' thumbnail image width size, pixels
    Dim lDestHgt        As Long         ' thumbnail image height size, pixels
    
    Dim lBMP            As BITMAP       ' Bitmap UDT - used to get the StdPic dimensions
    Dim saveType        As String       ' extension for the saved thumbnail
    
    Dim screenDC        As Long         ' Screen's device context.
    screenDC = GetDC(0)
    
' STEP 1: Calculate our destination x, y, width, and height.
    
    GetObject aStdPicture.handle, Len(lBMP), lBMP
    lDestWid = lBMP.bmWidth
    lDestHgt = lBMP.bmHeight
    ' Scale width and height so they fit inside the thumbnail
    If lDestWid > mMaxWidth Then
        lDestHgt = lDestHgt * (mMaxWidth / lDestWid)
        lDestWid = mMaxWidth
    End If
    If lDestHgt > mMaxHeight Then
        lDestWid = lDestWid * (mMaxHeight / lDestHgt)
        lDestHgt = mMaxHeight
    End If
    ' Calculate the destination x, y for BitBlting
    If mForcedDim Then
        lDestX = (mMaxWidth - lDestWid) / 2
        lDestY = (mMaxHeight - lDestHgt) / 2
    Else
        lDestX = 0
        lDestY = 0
    End If
' STEP 2: Create destination image and fill in the background color
    
    ' Create the Destination DCs and bitmaps
    lDestDC = CreateCompatibleDC(screenDC)
    If mForcedDim Then
        lDestBmp = CreateCompatibleBitmap(screenDC, mMaxWidth, mMaxHeight)
    Else
        lDestBmp = CreateCompatibleBitmap(screenDC, lDestWid, lDestHgt)
    End If
    DeleteObject SelectObject(lDestDC, lDestBmp)
    
    If mForcedDim Then
        ' Determine the RGB value for the back color specified (might be a
        '  system color, so we "translate" it first)
        OleTranslateColor mBackColor, 0, lBackColor
        ' Create a brush to fill in the background color
        lBackBrush = CreateSolidBrush(lBackColor)
        ' Fill in the background color
        SetRect lRECT, 0, 0, mMaxWidth, mMaxHeight
        FillRect lDestDC, lRECT, lBackBrush
        ' Delete our background brush
        DeleteObject lBackBrush
    End If
    
' STEP 3: Resize the original bitmap and bitblt it into our temporary DC

    lPicDC = CreateCompatibleDC(screenDC)
    lTempBmp = SelectObject(lPicDC, aStdPicture.handle)
    ' Transfer the image
    SetStretchBltMode lDestDC, 4
    StretchBlt lDestDC, lDestX, lDestY, lDestWid, lDestHgt, lPicDC, 0, 0, lBMP.bmWidth, lBMP.bmHeight, vbSrcCopy
    ' Replace the original bitmap
    SelectObject lPicDC, lTempBmp
    
' STEP 4: Create the objects and save
    saveType = LCase$(right$(mDestFilePath, 3))
    
    Select Case saveType
        Case "jpg"
            ' Create the DIB
            Set lDestDIB = New cDIBSection
            If mForcedDim Then
                lDestDIB.Create mMaxWidth, mMaxHeight
                ' Transfer the image into it
                BitBlt lDestDIB.hdc, 0, 0, mMaxWidth, mMaxHeight, lDestDC, 0, 0, vbSrcCopy
            Else
                lDestDIB.Create lDestWid, lDestHgt
                ' Transfer the image into it
                BitBlt lDestDIB.hdc, 0, 0, lDestWid, lDestHgt, lDestDC, 0, 0, vbSrcCopy
            End If
            SaveJPG lDestDIB, mDestFilePath
        Case "bmp"
            Set lDestStdPic = CreateBitmapPicture(lDestBmp, 0)
            SavePicture lDestStdPic, mDestFilePath
    End Select
    ' Clean up our stuff
    DeleteDC lDestDC
    DeleteObject lDestBmp

' STEP 5: Clean up
    
    Set lDestDIB = Nothing
    Set lDestStdPic = Nothing
    ReleaseDC 0, screenDC
    
End Sub

Public Sub CreateFromFile(aFilePath As String)
    
    Dim aStdPic As StdPicture
    On Error Resume Next
        Set aStdPic = LoadPicture(aFilePath)
        If Err.Number = 0 Then CreateFromStdPicture aStdPic
    On Error GoTo 0
    
End Sub

Private Sub Class_Initialize()
    mBackColor = vbWhite
    mMaxWidth = 50
    mMaxHeight = 50
End Sub

Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    
    Dim r               As Long
    Dim Pic             As PicBmp
    Dim IPic            As IPicture
    Dim IID_IDispatch   As GUID

    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With Pic
        .Size = Len(Pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    'Create the picture
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function
