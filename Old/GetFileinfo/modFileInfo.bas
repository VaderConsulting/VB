Attribute VB_Name = "modFileInfo"
Option Explicit
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
'Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Function GetFileinfo(FullFileName As String, ApplicationName As String, ApplicationVersion As String) As String
    Dim Buffer As String
    Dim rc As Long
    Dim FileVersionInfo As String
    
    Buffer = String(255, 0)
    'rc = GetSystemDirectory(Buffer, Len(Buffer))
    'Buffer = LCase$(Mid$(Buffer, 1, InStr(Buffer, Chr(0)) - 1))
    'FullFileName = Buffer & "\gdi32.dll"
    
    Dim lBufferLen As Long, lDummy As Long
    '*** Get size ****
    lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
    If lBufferLen < 1 Then
        'MsgBox "No Version Info available!"
        Exit Function
    End If
    
    Dim sBuffer()  As Byte
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
    If rc = 0 Then
        'MsgBox "GetFileVersionInfo failed."
        Exit Function
    End If
    
    Dim lVerPointer As Long
    
    rc = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
    
    If rc = 0 Then
        'MsgBox "VerQueryValue failed."
        Exit Function
    End If
    'lVerPointer is a pointer to four 4 bytes of Hex number,
    'first two bytes are language id, and last two bytes are code
    'page. However, Lang_Charset_String needs a  string of
    '4 hex digits, the first two characters correspond to the
    'language id and last two the last two character correspond
    'to the code page id.
    
    Dim bytebuffer(255) As Byte
    MoveMemory bytebuffer(0), lVerPointer, lBufferLen
    Dim Lang_Charset_String As String
    Dim HexNumber As Long
    
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
    Lang_Charset_String = Hex(HexNumber)
    'now we change the order of the language id and code page
    'and convert it into a string representation.
    'For example, it may look like 040904E4
    'Or to pull it all apart:
    '04------        = SUBLANG_ENGLISH_USA
    '--09----        = LANG_ENGLISH
    ' ----04E4 = 1252 = Codepage for Windows:Multilingual
    
    Do While Len(Lang_Charset_String) < 8
        Lang_Charset_String = "0" & Lang_Charset_String
    Loop
    
    'List1.Clear
    
    Dim strVersionInfo(7) As String
    strVersionInfo(0) = "CompanyName"
    strVersionInfo(1) = "FileDescription"
    strVersionInfo(2) = "FileVersion"
    strVersionInfo(3) = "InternalName"
    strVersionInfo(4) = "LegalCopyright"
    strVersionInfo(5) = "OriginalFileName"
    strVersionInfo(6) = "ProductName"
    strVersionInfo(7) = "ProductVersion"
    
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 7
        Buffer = String(255, 0)
        strTemp = "\StringFileInfo\" & Lang_Charset_String & "\" & strVersionInfo(i)
        rc = VerQueryValue(sBuffer(0), strTemp, lVerPointer, lBufferLen)
        
        If rc = 0 Then
            'MsgBox "VerQueryValue failed at" & i
        Exit Function
        End If
        
        lstrcpy Buffer, lVerPointer
        Buffer = Mid$(Buffer, 1, InStr(Buffer, Chr(0)) - 1)
        
        Select Case i
            Case 1, 3
                If FileVersionInfo <> "" Then
                    FileVersionInfo = FileVersionInfo & "," & Trim(Buffer)
                Else
                    FileVersionInfo = Trim(Buffer)
                End If
            Case 6
                If FileVersionInfo <> "" Then
                    FileVersionInfo = FileVersionInfo & "," & Trim(Buffer)
                Else
                    FileVersionInfo = Trim(Buffer)
                End If
                ApplicationName = Trim(Buffer)
            Case 7
                If FileVersionInfo <> "" Then
                    FileVersionInfo = FileVersionInfo & "," & Trim(Buffer)
                Else
                    FileVersionInfo = Trim(Buffer)
                End If
                ApplicationVersion = Trim(Buffer)
        End Select
        GetFileinfo = FileVersionInfo
    Next i
End Function

