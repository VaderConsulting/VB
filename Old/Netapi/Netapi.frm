VERSION 5.00
Begin VB.Form frmNetAPI 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmNetAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
   Private Const GMEM_FIXED = &H0
   Private Const GMEM_ZEROINIT = &H40
   Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
 
   Private Declare Function GlobalAlloc Lib "KERNEL32" ( _
      ByVal wFlags As Long, ByVal dwBytes As Long) As Long
   Private Declare Function GlobalFree Lib "KERNEL32" ( _
      ByVal hMem As Long) As Long
 
   Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
      hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
 
   Private Declare Function CopyPointer2String Lib "KERNEL32" _
      Alias "lstrcpyA" ( _
      ByVal NewString As String, ByVal OldString As Long) As Long
 
   Private Sub Form_click()
      Dim hEnum As Long, lpBuff As Long, nr As NETRESOURCE
      Dim cbBuff As Long, cCount As Long
      Dim p As Long, res As Long, i As Long
 
      'Setup the NETRESOURCE input structure
      nr.dwUsage = RESOURCEUSAGE_CONTAINER
      nr.lpRemoteName = 0
      cbBuff = 1000
      cCount = &HFFFFFFFF
 
      'Open a Net enumeration operation handle: hEnum
      res = WNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_ANY, _
                         0, nr, hEnum)
      If res = 0 Then
         'Create a buffer large enough for the results.
         '1000 bytes should be sufficient
         lpBuff = GlobalAlloc(GPTR, cbBuff)
         'Call the enumeration function
         res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
         If res = 0 Then
            p = lpBuff
            Cls
            'WNetEnumResource fills the buffer with an array of NETRESOURCE
            'structures. Walk through the list and print each local and
            'remote name
            For i = 1 To cCount
               CopyMemory nr, ByVal p, LenB(nr)
               p = p + LenB(nr)
               Print PointerToString(nr.lpLocalName), _
                     PointerToString(nr.lpRemoteName)
            Next i
         Else
            MsgBox "Error: " & Err.LastDllError, vbOKOnly, _
                   "WNetEnumResources"
         End If
         If lpBuff <> 0 Then GlobalFree (lpBuff)
         WNetCloseEnum (hEnum) 'Close the enumeration
      Else
         MsgBox "Error: " & Err.LastDllError, vbOKOnly, "WNetOpenEnum"
      End If
   End Sub
 
   Private Function PointerToString(p As Long) As String
      'The values returned in the NETRESOURCE structures are pointers to
      'ANSI strings so they need to be converted to Visual Basic Strings
      Dim s As String
      s = String(255, Chr$(0))
      CopyPointer2String s, p
      PointerToString = Left(s, InStr(s, Chr$(0)) - 1)
   End Function


