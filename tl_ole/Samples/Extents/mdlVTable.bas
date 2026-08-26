Attribute VB_Name = "mdlDeclarations"
'*********************************************************************************************
'
' Adding items to UserControl context menu
'
' Declarations module
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

Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Const PAGE_EXECUTE_READWRITE& = &H40&

Declare Function lstrcpyW Lib "kernel32" (ByVal P As Long, ByVal P1 As Long) As Long
Declare Function lstrlenW Lib "kernel32" (ByVal P1 As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    
Enum IOleObject_vtable_Indexes
   IDX_SetClientSite = 4
   IDX_GetClientSite
   IDX_SetHostNames
   IDX_Close
   IDX_SetMoniker
   IDX_GetMoniker
   IDX_InitFromData
   IDX_GetClipboardData
   IDX_DoVerb
   IDX_EnumVerbs
   IDX_Update
   IDX_IsUpToDate
   IDX_GetUserClassID
   IDX_GetUserType
   IDX_SetExtent
   IDX_GetExtent
   IDX_Advise
   IDX_Unadvise
   IDX_EnumAdvise
   IDX_GetMiscStatus
   IDX_SetColorScheme
End Enum
Public Function Ptr2Str(ByVal Ptr As Long) As String

   Ptr2Str = String$(lstrlenW(Ptr), 0)
   lstrcpyW StrPtr(Ptr2Str), Ptr
   
End Function


'
' Replaces an entry in a object v-table
'
Public Function ReplaceVTableEntry(ByVal oObject As Long, ByVal nEntry As Integer, ByVal pFunc As Long) As Long
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
    ReplaceVTableEntry = pFuncOld
    
End Function


