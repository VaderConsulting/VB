VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DoubleClick to save the icon"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   240
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   375
      Left            =   80
      TabIndex        =   1
      Top             =   160
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   600
      Index           =   0
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type PicBmp
   Size As Long
   tType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private err As Boolean

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hicon As Long) As Long

Public Function GetIconFromFile(FileName As String, IconIndex As Long, UseLargeIcon As Boolean) As Picture
    'Parameters:
    'FileName - File (EXE or DLL) containing icons
    'IconIndex - Index of icon to extract, starting with 0
    'UseLargeIcon-True for a large icon, False for a small icon
    'Returns: Picture object, containing icon
    
    Dim hlargeicon As Long
    Dim hsmallicon As Long
    Dim selhandle As Long
    
    ' IPicture requires a reference to "Standard OLE Types."
    Dim pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID
    
    If ExtractIconEx(FileName, IconIndex, hlargeicon, hsmallicon, 1) > 0 Then
        If UseLargeIcon Then
            selhandle = hlargeicon
        Else
            selhandle = hsmallicon
        End If
        
        ' Fill in with IDispatch Interface ID.
        With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        ' Fill Pic with necessary parts.
        With pic
            .Size = Len(pic) ' Length of structure.
            .tType = vbPicTypeIcon ' Type of Picture (bitmap).
            .hBmp = selhandle ' Handle to bitmap.
        End With
        
        ' Create Picture object.
        Call OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
        
        ' Return the new Picture object.
        Set GetIconFromFile = IPic
        
        DestroyIcon hsmallicon
        DestroyIcon hlargeicon
        
    End If
End Function

Private Function LoadAllIacons(strFileName As String)
    On Error GoTo errHandle
    Dim i As Integer
    For i = 1 To Picture1.Count - 1
        Unload Picture1(i)
    Next
    i = 1
    While Not err
        Load Picture1(i)
        Picture1(i).Visible = True
        Picture1(i).Top = Picture1(0).Height * Int(i / 15) + 100
        Picture1(i).Left = Picture1(0).Width * (i Mod 15 - 1) + cmdOpen.Width + 200
        retval = GetIconFromFile(strFileName, i - 1, True)
        If retval = 0 Then err = True
        Set Picture1(i).Picture = GetIconFromFile(strFileName, i - 1, True)
        i = i + 1
    Wend
    Picture1(0).Visible = False
    Exit Function
errHandle:
    Picture1(0).Visible = False
    Unload Picture1(i)
End Function

Private Sub cmdOpen_Click()
    Dim strFileName
    cDialog.FileName = ""
    cDialog.ShowOpen
    strFileName = cDialog.FileName
    LoadAllIacons (strFileName)
End Sub

Private Sub Form_Load()
    LoadAllIacons (App.Path & "\moricons.dll")
End Sub

Private Sub Picture1_DblClick(Index As Integer)
    SavePicture Picture1(Index).Picture, App.Path & "\" & InputBox("Icon name", "Extract icons", "test") & ".ico"
End Sub
