'
' Author:      mpryor@sprintmail.com
' script: Holiday_Reminder.vbs
' Description: a script called from the TskScheduler to make holiday announcements 
' keywords: agent speak text2speech task job VoiceText
' Date:    1/3/02 11:56AM
' 
' ------------------------------------------------

' --- revision history ---------------------------
' 
' 
'    
' --- end of revisions ---------------------------
Option Explicit
On error resume next
dim oSayIt, MyText, oArgs, myday, myintro, sMyDate
dim WSHShell, bResult
set WSHShell = CreateObject("WScript.Shell")
' clear the sound card if Winamp is running
bResult=WSHShell.AppActivate(" Winamp") ' V2.091 has a blank at the start!
WScript.echo bResult
WScript.sleep 500
if bResult then WSHShell.sendkeys "v"
' sound card is clear
myday = WeekDay(Date())
sMyDate = formatdatetime(Date(),vbLongDate) 
sMyDate = Replace(sMyDate," 0"," ") ' remove leading zero
myintro = "Today is " & sMyDate
set oSayIt = WScript.CreateObject("Speech.VoiceText")
'On error resume next
'Set oArgs = WScript.Arguments
'if oArgs.count = 0 then WScript.quit
WScript.echo sMyDate
select case sMyDate
case "Monday, January 21, 2002"
MyText = myintro & ". Today is the MLK Birthday, made possible by a monday federal holiday."
case "Monday, February 11, 2002"
MyText = myintro & ". Today is Lincoln's birthday. Don't forget to shave."
case "Monday, February 18, 2002"
MyText = myintrol & ". Today is President's day--actually Washington's birthday."
case "Wednesday, February 13, 2002"
MyText = myintro & ". Today is Ash wednesday--don't forget to clean your forehead after church."

case "Sunday, March 31, 2002"
MyText = myintro & ". Today is easter sunday--I hope you don't see ghosts. If you see one, you know who it is."

'if (myday = 7 or myday = 1 ) then MyText = " "
case else
MyText = myintro & ". Event not found--you need to edit the Holiday Reminder script."
'if (myday = 5 or myday = 6 ) then MyText = " "
end select
'oSayIt.Enabled 
oSayIt.Register "", "Holiday_Reminder.vbs"
oSayIt.Speak MyText,&h100
if err then
WScript.echo err.description
end if
while oSayIt.IsSpeaking
Wscript.Sleep 200
Wend
set oSayIt = nothing
' restart winamp
if bResult then WSHShell.sendkeys "x"
WScript.sleep 1000
WScript.quit(0)


