VERSION 5.00
Begin VB.UserControl Test2 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   795
   ScaleWidth      =   1440
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This control has 2 items added to the context menu."
      Height          =   615
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1335
   End
End
Attribute VB_Name = "Test2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Design-Time Context Menu Sample - Test2"
'*********************************************************************************************
'
' UserControl design context menu sample
'
' Test control 2
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on: Windows 98 + VB5
'
' History:
'          02/09/2000 - The file was released
'
'*********************************************************************************************
Option Explicit

' Next item to be enumerated
Dim m_lCurrentVerb As Long

' Verbs array
Dim m_aOleVerbs() As OLEVERB_VB

' Implement IEnumOLEVERB so
' we can add items to the
' context menu
Implements olelib2.IEnumOLEVERB

' Implement the callback interface
Implements IOleVerbCallback

' Index of the border verb
Dim m_lBorderIdx As Long

Private Function IOleVerbCallback_AddMenuItem(ByVal Name As String, ByVal iVerb As Long, Optional ByVal Flags As ovcb.MenuFlags = 0&) As Long
Dim lIdx As Long

   On Error Resume Next
   
   lIdx = UBound(m_aOleVerbs) + 1
   
   ReDim Preserve m_aOleVerbs(0 To lIdx)
   
   With m_aOleVerbs(lIdx)
      .lpszVerbName = Name
      .lVerb = iVerb
      .fuFlags = Flags
      .grfAttribs = OLEVERBATTRIB_ONCONTAINERMENU
   End With

   IOleVerbCallback_AddMenuItem = lIdx
   
End Function

Private Function IOleVerbCallback_DoVerb(ByVal iVerb As Long, lpmsg As olelib.MSG, ByVal pActiveSite As olelib.IOleClientSite, ByVal lindex As Long, ByVal hwndParent As Long, lprcPosRect As olelib.RECT) As Long
   
   On Error Resume Next
   
   Select Case iVerb

      Case 3
         MsgBox "Test2 Control" & vbCrLf & "Eduardo A. Morcillo 2001"
      
      Case 4
         
         If UserControl.BorderStyle = 0 Then
            UserControl.BorderStyle = 1
            m_aOleVerbs(m_lBorderIdx).lpszVerbName = "Hide &Border"
         Else
            UserControl.BorderStyle = 0
            m_aOleVerbs(m_lBorderIdx).lpszVerbName = "Show &Border"
         End If
         
   End Select

End Function


Private Function IEnumOLEVERB_Clone() As IEnumOLEVERB

   Err.Raise vbObjectError
   
End Function

Private Function IEnumOLEVERB_Next(ByVal celt As Long, rgelt As OLEVERB) As Long
   
   If m_lCurrentVerb <= UBound(m_aOleVerbs) Then
      
      With m_aOleVerbs(m_lCurrentVerb)
         rgelt.grfAttribs = OLEVERBATTRIB_ONCONTAINERMENU
         rgelt.fuFlags = .fuFlags
         rgelt.lpszVerbName = StrPtr(.lpszVerbName)
         rgelt.lVerb = .lVerb
      End With
      
      m_lCurrentVerb = m_lCurrentVerb + 1
      
      IEnumOLEVERB_Next = 1
            
   Else
      
      m_lCurrentVerb = 0
      IEnumOLEVERB_Next = 0
      
      ' Raise some error to stop the enumeration
      Err.Raise E_FAIL
      
   End If
   
End Function

Private Sub IEnumOLEVERB_Reset()

   m_lCurrentVerb = 0
   
End Sub

Private Sub IEnumOLEVERB_Skip(ByVal celt As Long)
   
   m_lCurrentVerb = m_lCurrentVerb + celt
   
End Sub

Private Sub UserControl_Initialize()
   
   InitializeVerbs Me
   
   m_lBorderIdx = IOleVerbCallback_AddMenuItem("Hide &Border", 4)
   IOleVerbCallback_AddMenuItem "About...", 3
   
End Sub

Private Sub UserControl_Terminate()

   TerminateVerbs Me
   
End Sub


Private Sub UserControl_Resize()

   Width = 1440
   Height = 795
   
End Sub
