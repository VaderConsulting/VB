'=====================================
'
'  Description: Automate mstask with this TaskScheduler COM object
'
'         Name: Scheduler
'      Version: 1.0 [of Type Library]
'         GUID: {1DB362DD-A476-11D6-904C-444553540001}
'
'         File: C:\VBprojDev\typlib\TaskSch\ScriptInt\release\TaskScheduler.dll
'      Version: 4.0.0.0 [of file]
'
' Extracted on: 7/31/02 6:28:48 PM
'
'=====================================
' saErrCodes
'=====================================
Const ecACCESSDENIED = -2147024891
Const ecACCOUNT_DBASE_CORRUPT = -2147216623
Const ecACCOUNT_INFORMATION_NOT_SET = -2147216625
Const ecACCOUNT_NAME_NOT_FOUND = -2147216624
Const ecCANNOT_OPEN_TASK = -2147216627
Const ecINVALID_TASK = -2147216626
Const ecNO_SECURITY_SERVICES = -2147216622
Const ecSERVICE_NOT_INSTALLED = -2147216628
Const ecTASK_NOT_READY = -2147216630
Const ecTASK_NOT_RUNNING = -2147216629
Const ecTRIGGER_NOT_FOUND = -2147216631
Const ecUNKNOWN_OBJECT_VERSION = -2147216621
Const ecUNSUPPORTED_ACCOUNT_OPTION = -2147216620
'=====================================
' saPriority
'=====================================
Const HIGH_PRIORITY_CLASS = 128
Const IDLE_PRIORITY_CLASS = 64
Const NORMAL_PRIORITY_CLASS = 32
'=====================================
' saStatusCodes
'=====================================
Const scEVENT_TRIGGER = 267016
Const scTASK_DISABLED = 267010
Const scTASK_HAS_NOT_RUN = 267011
Const scTASK_NO_MORE_RUNS = 267012
Const scTASK_NO_VALID_TRIGGERS = 267015
Const scTASK_NOT_SCHEDULED = 267013
Const scTASK_READY = 267008
Const scTASK_RUNNING = 267009
Const scTASK_TERMINATED = 267014
'=====================================
' saTaskFlags
'=====================================
Const tfDeleteWhenDone = 2
Const tfDisabled = 4
Const tfDontStartIfOnBatteries = 64
Const tfHidden = 512
Const tfInteractive = 1
Const tfKillIfGoingOnBatteries = 128
Const tfKillOnIdleEnd = 32
Const tfRestartOnIdleResume = 2048
Const tfRunIfConnectedToInternet = 1024
Const tfRunOnlyIfDocked = 256
Const tfRunOnlyIfLoggedOn = 8192
Const tfStartOnlyIfIdle = 16
Const tfSystemRequired = 4096
'=====================================
' saTriggerDays
'=====================================
Const tdFriday = 32
Const tdMonday = 2
Const tdSaturday = 64
Const tdSunday = 1
Const tdThursday = 16
Const tdTuesday = 4
Const tdWednesday = 8
'=====================================
' saTriggerFlags
'=====================================
Const tfHasEndDate = 1
Const tfKillAtDurationEnd = 2
Const tfTriggerDisabled = 4
'=====================================
' saTriggerMonths
'=====================================
Const tmApril = 8
Const tmAugust = 128
Const tmDecember = 2048
Const tmFebruary = 2
Const tmJanuary = 1
Const tmJuly = 64
Const tmJune = 32
Const tmMarch = 4
Const tmMay = 16
Const tmNovember = 1024
Const tmOctober = 512
Const tmSeptember = 256
'=====================================
' saTriggerTypes
'=====================================
Const ttAtLogon = 7
Const ttAtSystemStart = 6
Const ttDaily = 1
Const ttMonthlyDate = 3
Const ttMonthlyWeekDay = 4
Const ttOnce = 0
Const ttOnIdle = 5
Const ttWeekly = 2
'=====================================
' saTriggerWeeks
'=====================================
Const twFirst = 1
Const twFourth = 4
Const twLast = 5
Const twSecond = 2
Const twThird = 3

       


