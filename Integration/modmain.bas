Attribute VB_Name = "modMain"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : Main
' DateTime  : 31-03-2003 07:34
' Author    : Dave Robinson
' Purpose   : Turns DocsOpen integration on or off
'
'  V    Date        Author          History
' 1.0   31-03-2003  Dave Robinson   Initial Version
' 1.1   02-03-2003  Dave Robinson   Added HKLM Keys
'---------------------------------------------------------------------------------------

Sub Main()
    Dim strCMD As String
    Dim oRegistry As REGTool5.Registry
    
    On Error GoTo Main_Error

    Set oRegistry = CreateObject("REGTool5.Registry")
    
    strCMD = LCase(Command$)
    
    On Error Resume Next
    Select Case strCMD
        Case "on"
            FileCopy App.Path & "\Integration\WordInXP.dot", "C:\Program Files\Microsoft Office\Office10\Startup\WordInXP.dot"
            FileCopy App.Path & "\Integration\DocsXLXP.xla", "C:\Program Files\Microsoft Office\Office10\XLStart\DocsXLXP.xla"
            FileCopy App.Path & "\Integration\AcrodX32.api", "C:\Program Files\Adobe\Acrobat 5.0\Reader\Plug_Ins\AcrodX32.api"
            oRegistry.UpdateKey REGTool5.HKEY_CLASSES_ROOT, "MS WORD\ODMA32", "", "PCDOCS"
            oRegistry.UpdateKey REGTool5.HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Exchange\Client\Extensions", "ExchEIM.Module", "4.0;ExchEIM.dll;1"
            oRegistry.UpdateKey REGTool5.HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Exchange\Client\Extensions", "ExchINS.Module", "4.0;ExchINS.dll;1"
        Case "off"
            Kill "C:\Program Files\Microsoft Office\Office10\Startup\WordInXP.dot"
            Kill "C:\Program Files\Microsoft Office\Office10\XLStart\DocsXLXP.xla"
            Kill "C:\Program Files\Adobe\Acrobat 5.0\Reader\Plug_Ins\AcrodX32.api"
            oRegistry.UpdateKey REGTool5.HKEY_CLASSES_ROOT, "MS WORD\ODMA32", "", ""
            oRegistry.UpdateKey REGTool5.HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Exchange\Client\Extensions", "ExchEIM.Module", ""
            oRegistry.UpdateKey REGTool5.HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Exchange\Client\Extensions", "ExchINS.Module", ""
        Case Else
    End Select
    End

Main_Error:

    Err.Clear
    Resume Next
    
End Sub
