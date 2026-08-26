Attribute VB_Name = "ModIniFiles"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As Any, _
   ByVal lpDefault As String, _
   ByVal lpReturnedString As String, _
   ByVal nSize As Long, _
   ByVal lpFileName As String) As Long
   
Private Declare Function WritePrivateProfileString Lib "kernel32" _
   Alias "WritePrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As Any, _
   ByVal lpString As Any, _
   ByVal lpFileName As String) As Long

Public Function GetIniValue(lpSectionName As String, _
                               lpKeyName As String, _
                               defaultValue As String, _
                               inifile As String) As String

'Retrieves a value from an ini file corresponding
'to the section and key name passed.

   Dim success As Long
   Dim nSize As Long
   Dim ret As String

  'call the API with the parameters passed.
  'The return value is the length of the string
  'in ret, including the terminating null. If a
  'default value was passed, and the section or
  'key name are not in the file, that value is
  'returned. If no default value was passed (""),
  'then success will = 0 if not found.

  'Pad a string large enough to hold the data.
   ret = Space$(2048)
   nSize = Len(ret)
   success = GetPrivateProfileString(lpSectionName, _
                                     lpKeyName, _
                                     defaultValue, _
                                     ret, _
                                     nSize, _
                                     inifile)

   If success Then
      GetIniValue = Left$(ret, success)
   End If

End Function

Public Sub SaveIniValue(lpSectionName As String, _
                           lpKeyName As String, _
                           lpValue As String, _
                           inifile As String)

'This function saves the passed value to the file,
'under the section and key names specified.
'If the ini file does not exist, it is created.
'If the section does not exist, it is created.
'If the key name does not exist, it is created.
'If the key name exists, it's value is replaced.

   Call WritePrivateProfileString(lpSectionName, _
                                  lpKeyName, _
                                  lpValue, _
                                  inifile)

End Sub

Public Sub DeleteIniSection(lpSectionName As String, _
                                inifile As String)

'this call will remove the entire section
'corresponding to lpSectionName. This is
'accomplished by passing vbNullString
'as both the lpKeyName and lpValue parameters.
'For example, assuming that an ini file had:
'  [Colours]
'  Colour1=Red
'  Colour2=Blue
'  Colour3=Green
'
'and this sub was called passing "Colours"
'as lpSectionName, the resulting Colours
'section in the ini file would be deleted.

   Call WritePrivateProfileString(lpSectionName, _
                                  vbNullString, _
                                  vbNullString, _
                                  inifile)

End Sub

Public Sub DeleteIniValue(lpSectionName As String, _
                             lpKeyName As String, _
                             inifile As String)

'this call will remove the keyname and its
'corresponding value from the section specified
'in lpSectionName. This is accomplished by passing
'vbNullString as the lpValue parameter. For example,
'assuming that an ini file had:
'  [Colours]
'  Colour1=Red
'  Colour2=Blue
'  Colour3=Green
'
'and this sub was called passing "Colour2"
'as lpKeyName, the resulting ini file
'would contain:
'  [Colours]
'  Colour1=Red
'  Colour3=Green

   Call WritePrivateProfileString(lpSectionName, _
                                  lpKeyName, _
                                  vbNullString, _
                                  inifile)

End Sub

Public Function StripNulls(startStrg As String) As String

'take a string separated by nulls,
'split off 1 item, and shorten the string
'so the next item is ready for removal.

'The passed string must have a terminating
'null for this function to work correctly.
'If you remain in a loop, check this first!

   Dim pos As Long
   Dim item As String

   pos = InStr(1, startStrg, Chr$(0))

   If pos Then

      item = Mid$(startStrg, 1, pos - 1)
      startStrg = Mid$(startStrg, pos + 1, Len(startStrg))
      StripNulls = item

   End If

End Function




