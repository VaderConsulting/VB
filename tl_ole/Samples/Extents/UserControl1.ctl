VERSION 5.00
Begin VB.UserControl UC_Extents 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
End
Attribute VB_Name = "UC_Extents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' UserControl extents
'
' Sample control
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

Dim m_Org_GetExtent As Long

Private Sub UserControl_Initialize()
Dim IOleObj As IOleObject
   
   ' Get the IOleObject interface
   ' of the control
   Set IOleObj = Me
   
   ' Replace IOleObject.GetExtent
   m_Org_GetExtent = ReplaceVTableEntry(ObjPtr(IOleObj), IDX_GetExtent, AddressOf mdlIOleObject.IOleObject_GetExtent)

End Sub

Private Sub UserControl_Paint()
Dim lX As Long

   For lX = -2 To ScaleWidth Step 16
      Line (lX, 0)-(lX, ScaleHeight), vbBlack
   Next
   
   For lX = -2 To ScaleHeight Step 16
      Line (0, lX)-(ScaleWidth, lX), vbRed
   Next
   
End Sub


Private Sub UserControl_Terminate()
Dim IOleObj As IOleObject
   
   ' Get the IOleObject interface
   ' of the control
   Set IOleObj = Me
   
   ' Restore original GetExtent method
   ReplaceVTableEntry ObjPtr(IOleObj), IDX_GetExtent, m_Org_GetExtent
   
End Sub

