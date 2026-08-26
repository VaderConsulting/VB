'
' Author:      Mark Pryor
' script:	TSRuntimes_test.vbs
' Description: qc the registration of Scheduler.SchAgent and test getruntimes 
' keywords: mstask automate notepad registry
' Date:    10/22/01 5:17AM
' 

option explicit
dim WSHShell ' global
Set WSHShell = WScript.CreateObject("WScript.Shell")

'Tsk Objects and vars
dim oSch, tsk
dim sStr
const tfTriggerDisabled = 4
const tfHasEndDate = 1
const tfKillAtDurationEnd = 2
const tfHidden = &h200

set oSch = WScript.CreateObject("Scheduler.SchAgent")
oSch.TargetComputer = vbNullString
oSch.refresh

' activate your job
set tsk = oSch.Job("SchDay")

sStr = sStr & "flgs=" & tsk.status & "  ec=" & tsk.exitcode & "  nextrun=" & tsk.Nextruntime & vbCrLf & _
   tsk.triggers.item(1).flags & vbCrLf & vbCrLf


'///////////// Modify the trigger ///////////
'tsk.triggers.item(2).flags = 0  'tfTriggerDisabled
'tsk.triggers.item(2).update
'tsk.save
'///////////////////

sStr = sStr & tsk.triggers.item(1).text & vbCrLf & vbCrLf

dim dtArr, dtRT, sPipes

' try different date intervals, even into the past back to the start date
sPipes = tsk.Runtimes( #8/1/02#, #9/1/02# )
sStr = sStr & " " & sPipes & vbCrLf & vbCrLf

dtArr = split( sPipes, "|")
' get individual dates as dtArr(0)...

dim sJobKey
sJobKey = "HKEY_CLASSES_ROOT\Scheduler.Job\"
sStr = sStr & "  " & WSHShell.RegRead( sJobKey ) & vbCrLf

dim sClsidKey
sClsidKey = WSHShell.RegRead( sJobKey & "CLSID\" )
sStr = sStr & sClsidKey & vbCrLf

sClsidKey = "HKEY_CLASSES_ROOT\CLSID\" & sClsidKey

'WScript.echo sClsidKey & "\TypeLib" 
sStr = sStr & WSHShell.RegRead(  sClsidKey & "\TypeLib\" ) & vbCrLf

npdump( sStr )

' ******* I get these values ***********
'top level object for each task
'{6BB94505-F1B7-11D4-904B-444553540001}
'{02B16B5C-7232-11D4-904A-444553540001}


function npdump(sOut)
' helper objects and vars
dim  objFS, TempName, objTempFile, MyBase

Set objFS = WScript.CreateObject("Scripting.FileSystemObject")
' setup the temp file sent to Notepad later
TempName = WshShell.Environment("PROCESS").Item("TEMP")
MyBase = objFS.GetTempName
MyBase = objFS.GetBaseName(MyBase)
TempName = TempName & "\~" & MyBase & ".txt"
set objTempFile = objFS.CreateTextFile( TempName,  false,false)
objTempFile.WriteLine sOut
objTempFile.Close
WSHShell.Run "notepad " & TempName, 1, True
objFS.DeleteFile TempName, True 
end function