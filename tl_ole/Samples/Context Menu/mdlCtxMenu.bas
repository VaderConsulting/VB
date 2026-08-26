Attribute VB_Name = "mdlCtxMenu"
'*********************************************************************************************
'
' UserControl design context menu sample
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

Public Const E_FAIL = &H80004005

' Control count
Dim m_lCount As Long

' Pointers to the original
' functions
Dim m_lEnumVerbs As Long
Dim m_lDoVerb As Long
Dim m_lNext As Long

Type OLEVERB_VB
    lVerb As Long
    lpszVerbName As String
    fuFlags As Long
    grfAttribs As OLEVERBATTRIB
End Type

Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Const PAGE_EXECUTE_READWRITE& = &H40&

Declare Function lstrcpyW Lib "kernel32" (ByVal P As Long, ByVal P1 As Long) As Long
Declare Function lstrlenW Lib "kernel32" (ByVal P1 As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)

Declare Sub CoTaskMemFree Lib "ole32" (ByVal ptr As Long)

Private Function pvPtr2Str(ByVal ptr As Long) As String

   pvPtr2Str = String$(lstrlenW(ptr), 0)
   lstrcpyW StrPtr(pvPtr2Str), ptr
   
End Function

'
' Replaces an entry in a object v-table
'
Private Function pvReplaceVTableEntry(ByVal oObject As Long, ByVal nEntry As Integer, ByVal pFunc As Long) As Long
Dim pFuncOld As Long, pVTableHead As Long
Dim pFuncTmp As Long, lOldProtect As Long
     
    ' Object pointer contains a pointer to v-table--copy it to temporary
    ' pVTableHead = *oObject;
    CopyMemory pVTableHead, ByVal oObject, 4
    
    ' Calculate pointer to specified entry
    pFuncTmp = pVTableHead + (nEntry - 1) * 4
    
    ' Save address of previous method for return
    ' pFuncOld = *pFuncTmp;
    CopyMemory pFuncOld, ByVal pFuncTmp, 4
    
    ' Ignore if they're already the same
    If pFuncOld <> pFunc Then
        ' Need to change page protection to write to code
        VirtualProtect pFuncTmp, 4, PAGE_EXECUTE_READWRITE, lOldProtect
        
        ' Write the new function address into the v-table
        CopyMemory ByVal pFuncTmp, pFunc, 4     ' *pFuncTmp = pfunc;
        
        ' Restore the previous page protection
        VirtualProtect pFuncTmp, 4, lOldProtect, lOldProtect 'Optional
        
    End If
    
    'return address of original proc
    pvReplaceVTableEntry = pFuncOld
    
End Function



Public Sub InitializeVerbs(ByVal Obj As IOleVerbCallback)
Dim oOleObject As IOleObject
Dim oEnumOV As IEnumOLEVERB, OV As OLEVERB
   
   ' Get the IOleObject interface
   ' of the control
   Set oOleObject = Obj

   ' Replace the vtable only in
   ' the first instance
   If m_lCount = 0 Then
              
      ' Enumerate verbs that
      ' VB adds to the control
      ' (Properties and Edit)
      oOleObject.EnumVerbs oEnumOV
      
   Else
   
      ' Replace IOleObject.EnumVerbs
      ' with the original function
      m_lEnumVerbs = pvReplaceVTableEntry(ObjPtr(oOleObject), 13, m_lEnumVerbs)
      
      ' Get the original object
      oOleObject.EnumVerbs oEnumOV
      
      ' Replace IOleObject.EnumVerbs
      m_lEnumVerbs = pvReplaceVTableEntry(ObjPtr(oOleObject), 13, AddressOf IOleObject_EnumVerbs)

   End If
   
   Do While oEnumOV.Next(1, OV) = 0
   
      ' Add the verb to the array
      Obj.AddMenuItem pvPtr2Str(OV.lpszVerbName), OV.lVerb, OV.fuFlags
      
      ' Release the string pointer
      CoTaskMemFree OV.lpszVerbName
      
   Loop
      
   Set oEnumOV = Nothing
      
   If m_lCount = 0 Then
      
      ' Replace IOleObject.EnumVerbs
      m_lEnumVerbs = pvReplaceVTableEntry(ObjPtr(oOleObject), 13, AddressOf IOleObject_EnumVerbs)
      
      ' Replace IOleObject.DoVerb
      m_lDoVerb = pvReplaceVTableEntry(ObjPtr(oOleObject), 12, AddressOf IOleObject_DoVerb)
      
   End If
   
   m_lCount = m_lCount + 1

End Sub

Public Sub TerminateVerbs(ByVal Obj As IOleVerbCallback)
Dim oOleObject As IOleObject
   
   ' Decrement the control count
   m_lCount = m_lCount - 1
   
   ' Restore the original entries
   ' when the last control is
   ' destroyed
   If m_lCount <= 0 Then
   
      m_lCount = 0
      
      ' Get the IOleObject interface
      ' of the control
      Set oOleObject = Obj
      
      ' Restore original EnumVerbs method
      pvReplaceVTableEntry ObjPtr(oOleObject), 13, m_lEnumVerbs
      
      ' Restore original DoVerb method
      pvReplaceVTableEntry ObjPtr(oOleObject), 12, m_lDoVerb
            
   End If
   
End Sub
 
'*********************************************************************************************
' This function replaces the original
' DoVerb method
'*********************************************************************************************
Public Function IOleObject_DoVerb(ByVal This As IOleObject, ByVal iVerb As Long, lpmsg As MSG, ByVal pActiveSite As IOleClientSite, ByVal lindex As Long, ByVal hwndParent As Long, lprcPosRect As RECT) As Long
Dim oCallback As IOleVerbCallback
   
   On Error Resume Next
   
   Set oCallback = This
    
   If iVerb <= 2 Or oCallback Is Nothing Then
   
      ' The object doesn't implements
      ' IOleVerbCallback. Call the original
      ' function.
         
      ' Replace the pointer with the original function
      pvReplaceVTableEntry ObjPtr(This), 12, m_lDoVerb
      
      ' Call the original method
      This.DoVerb iVerb, lpmsg, pActiveSite, lindex, hwndParent, lprcPosRect
      
      ' Restore the new method
      pvReplaceVTableEntry ObjPtr(This), 12, AddressOf mdlCtxMenu.IOleObject_DoVerb
            
   Else
   
      ' Call the control function
      IOleObject_DoVerb = oCallback.DoVerb(iVerb, lpmsg, pActiveSite, lindex, hwndParent, lprcPosRect)
   
   End If
   
End Function

'*********************************************************************************************
' This function replaces the original
' EnumVerbs method
'*********************************************************************************************
Public Function IOleObject_EnumVerbs(ByVal This As IOleObject, IEnumVERB As IEnumOLEVERB) As Long

   On Error Resume Next
   
   ' Return a reference to
   ' the IEnumVERB. "This"
   ' points to the UserControl.
   Set IEnumVERB = This
   
   If IEnumVERB Is Nothing Then
       
      pvReplaceVTableEntry ObjPtr(This), 13, m_lEnumVerbs
      
      IOleObject_EnumVerbs = This.EnumVerbs(IEnumVERB)
      
      pvReplaceVTableEntry ObjPtr(This), 13, AddressOf mdlCtxMenu.IOleObject_EnumVerbs
   
   End If

End Function




