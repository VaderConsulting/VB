Attribute VB_Name = "Module1"
Option Explicit

Declare Function ExtractAssociatedIconA Lib "shell32.dll" _
                ( _
                 ByVal hInst As Long, _
                 ByVal lpIconPath As String, _
                 lpiIcon As Long _
                ) As Long

Declare Function DrawIcon Lib "user32" _
                ( _
                 ByVal hdc As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal hIcon As Long _
                ) As Long
Declare Function DestroyIcon Lib "user32" _
                ( _
                 ByVal hIcon As Long _
                ) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" _
                ( _
                 ByVal hInst As Long, _
                 ByVal lpszExeFileName As String, _
                 ByVal nIconIndex As Long _
                ) As Long

