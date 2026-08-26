VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "Simple Clipboard 'Viewer'"
   ClientHeight    =   3495
   ClientLeft      =   3975
   ClientTop       =   2250
   ClientWidth     =   5190
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5190
   Begin VB.CheckBox chkHex 
      Caption         =   "Show Hex"
      Height          =   255
      Left            =   2460
      TabIndex        =   5
      Top             =   300
      Width           =   2535
   End
   Begin VB.TextBox txtView 
      Height          =   2775
      Left            =   2460
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.ListBox lstFormats 
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   2295
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   2460
      ScaleHeight     =   2715
      ScaleWidth      =   2595
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label lblView 
      Caption         =   "Text/Picture View:"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   60
      Width           =   2655
   End
   Begin VB.Label lblFormats 
      Caption         =   "Clipboard Formats"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2295
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Class provides clipboard viewer properties and methods:
Private WithEvents m_cCV As cClipboardViewer
Attribute m_cCV.VB_VarHelpID = -1
' Class to access the clipboard without restriction:
Private m_cClip As New cCustomClipboard

Private Sub chkHex_Click()
    
    lstFormats_Click
End Sub

Private Sub Form_Load()
   ' Here we initialise the clipboard viewer class.
   ' This makes the application part of the clipboard
   ' chain, so that whenever the clipboard contents
   ' change, the class will raise a ClipboardChanged
   ' event:
   Set m_cCV = New cClipboardViewer
   m_cCV.InitClipboardChangeNotification Me.hWnd
End Sub

Private Sub lstFormats_Click()
Dim lId As Long
Dim sText As String
Dim hBmp As Long
Dim sOut As String
Dim sHex As String
Dim bData() As Byte
Dim l As Long

    Screen.MousePointer = vbHourglass
    txtView.Text = ""
    If (lstFormats.ListCount > 0) Then
    lId = lstFormats.ItemData(lstFormats.ListIndex)
    Select Case lId
    Case CF_BITMAP
        ' Here is the simple way!
        Set picView = Clipboard.GetData(CF_BITMAP)
        picView.ZOrder
        
        ' As an example, here is the hard way to do it
        ' - the clipboard contains a bitmap handle:
        'picView.Cls
        'If (m_cClip.ClipboardOpen(Me.hWnd)) Then
        '   dim hBmp as long, hBmpOld as long, hDC as long, tBM as BITMAP
        '    hBmp = m_cClip.GetClipboardMemoryHandle(CF_BITMAP)
        '    hDC = CreateCompatibleDC(picView.hDC)
        '    hBmpOld = SelectObject(hDC, hBmp)
        '    GetObjectAPI pic.Handle, Len(tBM), tBM
        '    BitBlt picView.hDC,0,0, tBM.bmWidth,tBM.bmHeight, hDC,0,0,SRCCOPY
        '    SelectObject hDC,hBmpOld
        '    DeleteObject hDC
        '    m_cClip.ClipboardClose
        'End If
    Case CF_DIB
        ' The clipboard handle points to a BITMAPINFO structure
        ' followed by the bitmap bits:
        Set picView = Clipboard.GetData(CF_DIB)
        picView.ZOrder
        
    Case CF_ENHMETAFILE
        ' The clipboard handle is a handle to an Enhanced Metafile:
        Set picView = Clipboard.GetData(CF_ENHMETAFILE)
        picView.ZOrder
        
    Case CF_HDROP
        ' the clipboard handle can be passed to the DragQueryFile
        ' function to get the information:
        txtView.ZOrder
        txtView.Text = ""
        
    Case CF_METAFILEPICT
        ' The clipboard handle is a handle to an old style Metafile:
        Set picView = Clipboard.GetData(CF_ENHMETAFILE)
        
    Case Else
        ' Assume Text
        If (m_cClip.ClipboardOpen(Me.hWnd)) Then
            If chkHex.Value = 1 Then
                ' Do Hex - get a byte array of clipboard
                ' data and show that:
                m_cClip.GetBinaryData lId, bData()
                For l = LBound(bData) To UBound(bData)
                    sHex = Hex$(bData(l))
                    If (Len(sHex) < 2) Then sHex = "0" & sHex
                    sOut = sOut & sHex
                    If ((l + 1) Mod 8) = 0 Then
                        sOut = sOut & vbCrLf
                    Else
                        sOut = sOut & "  "
                    End If
                Next l
                txtView = sOut
            Else
                ' Get a string and show that:
                If (m_cClip.GetTextData(lId, sText)) Then
                    txtView.Text = sText
                End If
            End If
            m_cClip.ClipboardClose
        End If
        txtView.ZOrder
    End Select
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub m_cCV_ClipboardChanged()
Dim lCount As Long
Dim lFormat As Long
Dim lId As Long
   Debug.Print "Clipboard changed"
   With m_cClip
      lCount = .GetCurrentFormats(Me.hWnd)
      lstFormats.Clear
      For lFormat = 1 To lCount
         lId = .GetCurrentFormatID(lFormat)
         lstFormats.AddItem CStr(lId) & vbTab & .GetCurrentFormatName(lFormat)
         lstFormats.ItemData(lstFormats.NewIndex) = lId
         If (lId = 1) Then ' Text
            lstFormats.ListIndex = lstFormats.NewIndex
         End If
      Next lFormat
      
   End With
End Sub
