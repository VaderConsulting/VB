Attribute VB_Name = "modAPI"
Public Declare Function CopyFile Lib "KERNEL32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : APIFileCopy
' DateTime  : 20-04-2003 18:58
' Author    : Dave Robinson
' Purpose   : COPY FILES
'PARAMETERS: src: Source File (FullPath)
'           dest: Destination File (FullPath)
'           FailIfDestExists (Optional):
'                         Set to true if you don't want to overwrite the destination file if it exists

'           Returns (True if Successful, false otherwise)

'           EXAMPLE:
'           dim bSuccess as boolean
'           bSuccess = APIFileCopy ("C:\MyFile.txt", "D:\MyFile.txt")
'
'  V    Date    Author      History
' 1.0   20-04-2003  Dave Robinson   Initial Version - From http://www.freevbcode.com/ShowCode.Asp?ID=498
'---------------------------------------------------------------------------------------
Public Function APIFileCopy(src As String, dest As String, Optional FailIfDestExists As Boolean) As Boolean
    Dim lRet As Long
    
    lRet = CopyFile(src, dest, FailIfDestExists)
    APIFileCopy = (lRet > 0)
End Function
