Attribute VB_Name = "modForm1"
'************************************************************'
'Filename :       modForm1.bas                               '
'Author:          Coenraad Steenkamp                         '
'Created:         09/05/2000                                 '
'Operating System Windows 9x & NT ,2000 ,Me                  '
'                                                            '
'Description:                                                '
'In this module i declare al the variables for my Array and  '
'Recordings of the high score.There is also a few variables  '
'for the sound generator in the game.Other declarations      '
'include the use of a 3d Starfield                           '
'                                                            '
'************************************************************'
   'Let all array's start at 1
Option Base 1
Option Explicit
   'Make public type Recordings
Public Type Recordings
   StrName As String * 15
   iScore As Integer
End Type
   'Decalare Global Variables
   Public intRet As Integer
   Public gudtRecord As Recordings
   Public gudtRecordArray(15) As Recordings
   Public giNumRecords As Integer
   Public miFileNum As Integer
   Public gudRec As Recordings
   Public gOrb As Integer
   Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
   'Declare function to play music and sfx
   Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
   Public Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
   Public Const Music = "passport.mid"
   Public Const Music2 = "shot.wav"

'Declare the Variables for the Starfield
   Public iStarcolor As Integer
   Public iRadius As Integer
   Public iK As Integer
   Public iZoom As Integer
   Public iSpeed As Integer
