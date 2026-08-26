#!perl/bin/ 
# author="Mark Pryor mpryor@sprintmail.com"
# keywords="Sage Scheduler pl KnowledgeBase kb library tool utility"
# description=" builds an html page of all TS KB articles as of 1-31-01"
# script="MakeKB.pl"
#

# module from CPAN
use Text::ParseWords;

my ($FileName,$FilePath,$Title ,$Excerpts ,$Size ,$Time)="";
my $urlBase = 'http://support.microsoft.com/support/kb/articles/';

# add this line to use as CGI
print "CONTENT_TYPE: text/html", "\n\n";

# --- print the header ------
print <<HEADER;
<html>

<head>
<title>Knowledge Base for Task Scheduler</title>
<style type="text/css">
	BODY {
	font-family : Verdana, Arial, Helvetica ;	
 }
TD.main {
	 font-size : 75% ; 
	font-weight : bold;
	color: #00497A
}
</style>
<script type="text/javascript">
	function JumpTo( sPath ) {
	window.location.href = '$urlBase' + sPath;
	}
</script>
</head>

<body>
<h2> Knowledge Base for the Task Scheduler</h2>
<a href='http://home.sprintmail.com/~mpryor/c-frame.htm?tasksch.htm'> Home of this script </a><br>
<hr>
<TABLE BORDER='0' CELLPADDING='0' CELLSPACING='3' WIDTH='95%'>
HEADER
#  ------- resume code -----------

$_= <DATA>;
($FileName,$FilePath,$Title ,$Excerpts ,$Size ,$Time)= &parse_line(',', 1, $_);
print "$_\n";
$nCount = 0;
while (<DATA>) {
$nCount ++;
($FileName,$FilePath,$Title ,$Excerpts ,$Size ,$Time)= &parse_line(',', 0, $_);
#printf "Tile= $Title \n";
$FilePath=~ s/^\s+//;

print <<ROW
<TR WIDTH='100%'>
		<TD  WIDTH='20' VALIGN='top'>$nCount   </TD>
		<TD class='main'>
		<A href='javascript:JumpTo("$FilePath");'>$Title </A>
		</TD>
		</TR>
ROW
}
print <<FOOTER
</TABLE>
</body>
</html>
FOOTER
# 4.7 [en] (Win95; I) 
__END__;

"74.asp", "Q103/4/74.asp", "Q103474 - Implementing Scheduled Backups with Windows NT Backup", "Implementing Scheduled Backups with Windows NT Backup The information in this article applies to:       Microsoft Windows NT Server version  3.1      Microsoft Windows NT Workstation version  3.1      Microsoft Windows NT Advanced Server      Microsoft Windows NT Workstation versions  3.5, 3.51,", "8959", "10/26/2000 2:30:00 AM"
"28.asp", "Q109/2/28.asp", "Q109228 - SQL Server and Windows NT Thread Scheduling", "SQL Server for Windows NT is a Win32 application that runs on the Windows  NT or Windows NT Advanced Server operating systems. In some uncommon  situations, when SQL Server is running a compute-bound operation, the  interactive console respons", "18523", "10/26/2000 2:18:00 AM"
"18.asp", "Q112/2/18.asp", "Q112218 - Scheduled NTBackup Must Be Configured Using System Account", "If no users are currently logged on to Windows NT, a scheduled NTBACKUP  doesnt run unless the Scheduler service is configured to &quot;Log On As&quot; the  System Account. This may be a problem if you plan to implement user-level  security on director", "6020", "10/26/2000 2:40:00 AM"
"73.asp", "Q113/0/73.asp", "Q113073 - AT Command Fails to Run Some Applications", "If you have either a batch file or a command file that contains OS/2  commands and uses the Schedule service to run, OS/2 commands in the file  are not executed.", "6878", "10/26/2000 1:42:00 AM"
"92.asp", "Q115/6/92.asp", "Q115692 - BUG: No Video Refresh When .PIF Runs When Logged Off", "When you use the", "5741", "10/26/2000 1:05:00 AM"
"19.asp", "Q119/9/19.asp", "Q119919 - COPY Command Does Not Execute from the AT Command", "When you start the Schedule service and then enter an AT command to  instruct the COPY command to run at a specified time, the COPY command  does not run at the specified time and the following message appears in  Event Log:", "6987", "10/26/2000 5:59:00 AM"
"62.asp", "Q121/5/62.asp", "Q121562 - Applications Started with AT Are Not Interactive", "The Windows NT 3.5 AT command does not load applications interactively even  though the System Account Allow Service to Interact with Desktop check box  is selected for the Scheduler service. The AT command does load the  application in the ba", "7772", "10/26/2000 2:43:00 AM"
"27.asp", "Q125/6/27.asp", "Q125627 - INF: Starting SQL Server Remotely from the NT Command Prompt", "The SQL Service Manager provides a way to remotely start and stop SQL  Server. However, the SQL Service Manager cannot be used to start and stop  SQL Server from a batch file or Windows NT Command Prompt. The AT scheduler from Windows NT, the", "7489", "10/26/2000 3:37:00 AM"
"34.asp", "Q128/2/34.asp", "Q128234 - Scheduling Windows NT Backup Fails and Locks Backup Process", "When there is no tape in the drive and you use the AT command without the  interactive switch to run Windows NT Backup, the backup encounters an error  or is unable to accept the command line.  As a result, Windows NT Backup  stops responding ", "7034", "10/22/2000 1:44:00 AM"
"32.asp", "Q142/4/32.asp", "Q142432 - Problems Running Batch Files Through Scheduler Service", "When you run batch files through the Windows NT scheduler service, the  following symptoms occur:", "5839", "5/11/1999 12:38:00 PM"
"84.asp", "Q156/6/84.asp", "Q156684 - How to Use NLTEST to Force a New Secure Channel", "To validate access to resources in a trusting domain, the trusting  domains primary domain controller (PDC) establishes a secure channel  with a domain controller in the trusted domain. Pass-through  authentication then occurs over this secur", "9037", "1/22/1999 9:39:00 PM"
"43.asp", "Q161/0/43.asp", "Q161043 - SMS: Despooler Error: WIN32 ERROR = 32", "A single job may cause the distribution process to stop responding. The  job in question will report the following error in the Despooler log file:", "6792", "9/2/1999 1:50:00 PM"
"94.asp", "Q162/2/94.asp", "Q162294 - INFO: How to Schedule Backup &amp; DBCC Commands Using AT Scheduler", "There are times when you want jobs to be scheduled through the Windows NT  Server AT scheduler rather than scheduled through the task scheduling  manager used by the SQL Executive service. For example, you may want to  schedule a job to stop S", "9346", "10/26/2000 3:24:00 AM"
"55.ASP", "Q171/6/55.ASP", "Q171655 - Installing Task Scheduler Enables System Agent Tasks", "After you install the Task Scheduler component, scheduled tasks may  appear in the Scheduled Tasks folder.", "7181", "7/25/2000 2:14:00 PM"
"80.asp", "Q172/0/80.asp", "Q172080 - System Agent Does Not Function After Uninstalling Task Scheduler", "After you uninstall Task Scheduler, the System Agent icon does not appear  on the taskbar. If you attempt to start System Agent by using the Start  menu, you may receive the following error message:", "6659", "8/20/1999 10:47:00 PM"
"18.ASP", "Q174/4/18.ASP", "Q174418 - Internal Commands not Recognized with Schedule Service in Windows NT 4.0", "In Windows NT 3.51, output from a batch file run by the Schedule service  can be redirected to a log file. An example of this procedure is:", "6424", "1/16/1999 10:32:00 PM"
"24.ASP", "Q176/4/24.ASP", "Q176424 - Installing Task Scheduler Enables AT-Scheduled Tasks", "After you install the Task Scheduler component, scheduled tasks may appear  in the Scheduled Tasks folder.", "7632", "8/20/1999 8:13:00 PM"
"39.ASP", "Q177/0/39.ASP", "Q177039 - Unable to Change the Account Information for Scheduled Tasks", "When you attempt to change the account information for the Task Scheduler  service, the option to use This Account is unavailable. The only available  option under Log On As is the &quot;Allow Service to Interact with Desktop&quot;  check box.", "7629", "6/3/1999 8:07:00 PM"
"24.ASP", "Q177/1/24.ASP", "Q177124 - Settings Button Missing in Scheduled Tasks Tool", "When you view the properties for a ScanDisk for Windows or Disk  Defragmenter task in the Scheduled Tasks tool, the Settings button may be  missing.", "6489", "1/4/2001 6:51:00 PM"
"55.ASP", "Q178/5/55.ASP", "Q178555 - How to Move a Scheduled Job to Another NT Server", "This article describes how a Windows NT Administrator can move scheduled  jobs from one Windows NT 4.0 server to another. For example, a server with  several different batch jobs scheduled to run at various times throughout  the week needs to ", "8445", "1/29/1999 4:56:00 PM"
"91.ASP", "Q178/6/91.ASP", "Q178691 - Scheduled Program Does Not Start in Task Scheduler", "When Task Scheduler attempts to start a scheduled program, the program may  not start, and Could Not Start may appear as the status of the program.", "7944", "8/20/1999 7:34:00 AM"
"06.ASP", "Q178/7/06.ASP", "Q178706 - How to Schedule a Program Using Task Scheduler", "This article describes how to schedule a program by using Task Scheduler.", "6105", "1/8/2001 11:33:00 AM"
"06.ASP", "Q179/3/06.ASP", "Q179306 - How to Automate ScanDisk and Disk Defragmenter Using Task Scheduler Tool", "This article describes how to configure ScanDisk and Disk Defragmenter to  run without user intervention using the Task Scheduler tool.", "7674", "7/24/2000 6:59:00 PM"
"69.ASP", "Q179/3/69.ASP", "Q179369 - ScanDisk Settings Are Not Applied in Task Scheduler", "When a scheduled ScanDisk session is started by Task Scheduler, the  session may not use the settings you last selected in ScanDisk.", "5938", "1/8/2001 2:30:00 PM"
"62.ASP", "Q179/6/62.ASP", "Q179662 - Task Scheduler Restarts After You Disable It", "When you disable Task Scheduler, Task Scheduler may restart after you  start Microsoft WebTV for Windows.", "5798", "1/30/1999 9:57:00 PM"
"59.ASP", "Q180/0/59.ASP", "Q180059 - StarSight Loader Task Displays &quot;Could Not Start&quot; Status", "When you start Task Scheduler, you may see a task named StarSight Loader  with a status of &quot;Could not start.&quot;", "5773", "1/30/1999 11:32:00 PM"
"56.ASP", "Q184/7/56.ASP", "Q184756 - Backup Does Not Perform Unattended Scheduled Backups", "Backup Does Not Perform Unattended Scheduled Backups The information in this article applies to:       Microsoft Windows 98       SYMPTOM When you use Task Scheduler to start an unattended backup job in  Microsoft Backup, the backup job does not complete if it is left  unattended. This behavior", "5639", "2/2/1999 11:16:00 PM"
"24.ASP", "Q186/9/24.ASP", "Q186924 - Links LS Legends in Sport 1997 Edition Setup Hangs in Windows 98", "After you install Links LS Legends in Sport 1997 Edition, you may receive  the following message:", "6162", "2/4/1999 9:00:00 PM"
"39.ASP", "Q187/2/39.ASP", "Q187239 - Cannot View Schedlog.txt File Using Task Scheduler", "When you attempt to view the Schedlog.txt file by clicking View Log on the  Advanced menu in Task Scheduler, the Schedlog.txt file may not open, and  you may receive an error message similar to the following error message:", "7081", "2/4/1999 11:10:00 PM"
"69.ASP", "Q187/2/69.ASP", "Q187269 - Scheduled Tasks Do Not Start on Time", "When you use Task Scheduler to start a program and you configure the  program to run after the computer has been idle, the program may not start  at the scheduled time even though the computer has been idle for the specified time.", "5198", "2/4/1999 11:31:00 PM"
"80.ASP", "Q187/3/80.ASP", "Q187380 - Plus! 98: How to Specify a Rotation Schedule for Desktop Themes", "When you select a Desktop Theme in Plus! 98, you are able to rotate the  desktop theme on a monthly interval. To rotate the theme on a schedule you  specify, you can use the Task Scheduler included in Microsoft Windows 98.", "6861", "9/9/1999 3:13:00 PM"
"60.ASP", "Q187/6/60.ASP", "Q187660 - Changing Channel in WebTV Stops Program Guide Download", "If you change channels in Microsoft WebTV for Windows while the program  guide is being downloaded, the download is not completed.", "6105", "2/5/1999 11:13:00 PM"
"51.ASP", "Q187/8/51.ASP", "Q187851 - Plus! 98: Password Can Be Bypassed if Screen Saver Is Minimized", "When you password protect your screen saver, you may be able to bypass  the screen saver password if you minimize the screen saver.", "6738", "9/9/1999 3:31:00 PM"
"53.ASP", "Q187/8/53.ASP", "Q187853 - Task Scheduler Task Does Not Run When Copied to Another Computer", "When you copy a Task Scheduler task (.job) file from one Windows NT-based  computer to another Windows NT-based computer, the copied task may not  run.", "6518", "12/18/1999 12:07:00 AM"
"82.ASP", "Q188/1/82.ASP", "Q188182 - Disk Defragmenter Causes General Protection Fault in User.exe", "When you try to run Disk Defragmenter from System Agent or Windows Task  Scheduler, you may receive a general-protection (GP) fault in module  User.exe.", "5583", "12/14/2000 10:25:00 PM"
"33.ASP", "Q188/2/33.ASP", "Q188233 - Unfinished Scheduled Tasks Are Incorrectly Listed as Finished", "When you view the Task Scheduler tool, the Last Run Time column may  indicate that a scheduled task has completed even though the task has not  completed.", "5542", "2/9/1999 10:18:00 PM"
"24.ASP", "Q188/6/24.ASP", "Q188624 - &quot;Wake the Computer to Run This Task&quot; Feature May Not Work", "When you use the &quot;Wake the computer to run this task&quot; option for a task in  Task Scheduler, your computer may not resume from a suspended state to run  the scheduled task at the specified time.", "5760", "2/11/1999 12:10:00 AM"
"30.ASP", "Q189/0/30.ASP", "Q189030 - Plus! 98: Cannot Start McAfee Virus Scan Console", "When you attempt to start the McAfee Virus Scan Console to schedule a  virus scan, you receive the following error message:", "5597", "9/9/1999 3:40:00 PM"
"92.ASP", "Q189/3/92.ASP", "Q189392 - Plus! 98: Tasks Scheduled Using Task Scheduler No Longer Run", "When you remove Microsoft Plus! 98, some tasks or programs scheduled using  Task Scheduler may no longer run.", "6366", "9/9/1999 3:37:00 PM"
"89.ASP", "Q189/6/89.ASP", "Q189689 - &quot;Hot Corner&quot; Features Do Not Work in Windows 98", "When you attempt to use the &quot;hot corner&quot; features to activate your screen  saver by moving your mouse pointer to a corner of your desktop in Windows  98, your screen saver does not start.", "5445", "2/12/1999 6:16:00 PM"
"39.ASP", "Q190/1/39.ASP", "Q190139 - Unable to Find a Message for Task Exit Code in Schedlog.txt File", "When you view the Task Scheduler log file (Schedlog.txt), you may see the  following entry, or one similar to it:", "8218", "1/4/2001 6:50:00 PM"
"68.ASP", "Q190/5/68.ASP", "Q190568 - Scheduled Tasks Do Not Start", "When you use Task Scheduler to schedule a program, the program may not start. When this occurs, you may receive the following error message:", "5744", "3/20/2000 10:37:00 AM"
"27.ASP", "Q191/4/27.ASP", "Q191427 - Unable to Delete Text in the Task Scheduler Log File", "When you try to delete text from the Task Scheduler log file  (schedlog.txt) by opening Task Scheduler, clicking Advanced, clicking  View Log, selecting the text that you want to delete and then pressing  DELETE, you may receive the following ", "6496", "2/13/1999 11:13:00 PM"
"52.ASP", "Q193/8/52.ASP", "Q193852 - Task Scheduler May Fail to Run Job When Scheduled", "When Task Scheduler is scheduled to perform a task, it may not perform  the task.", "6682", "1/4/2001 6:50:00 PM"
"14.ASP", "Q193/9/14.ASP", "Q193914 - Task Scheduler Is Paused While Downloading Components", "When you download a component from the Microsoft Windows Update site, Task  Scheduler is paused and no scheduled tasks are run until the component is  installed.", "5266", "1/4/2001 6:50:00 PM"
"96.ASP", "Q195/8/96.ASP", "Q195896 - Forward Slash in Scheduled Task Command Changes to Backslash", "If you create a scheduled task in which the command in the Run box is  enclosed in quotation marks and contains a command-line switch using a  forward slash, the forward slash becomes a backslash. For example, the  command in the Run box chang", "6314", "2/2/1999 2:48:00 PM"
"33.ASP", "Q195/9/33.ASP", "Q195933 - Cannot Disable Task Scheduler", "When you click the Stop Using Task Scheduler command on the Advanced menu  in Task Scheduler, the Task Scheduler icon may still appear on the taskbar  after you restart your computer.", "6363", "10/21/2000 1:58:00 AM"
"20.ASP", "Q214/4/20.ASP", "Q214420 - AT Scheduler and Task Scheduler Display Scheduled Task Incorrectly", "After you install the Scheduled Tasks component from Microsoft Internet Explorer, tasks that are scheduled to run today may be reported as being scheduled to run next month.", "5921", "2/27/1999 10:41:00 PM"
"37.ASP", "Q215/9/37.ASP", "Q215937 - Task Scheduler Service Does Not Start", "When you start Windows 98 or Windows Millennium Edition (Me), Task Scheduler may not start. When this occurs, you may see the following information in the Task Scheduler log file:", "6004", "10/21/2000 1:54:00 AM"
"06.ASP", "Q217/0/06.ASP", "Q217006 - How to Distribute Task Scheduler Tasks to Multiple Users", "This article describes how to distribute a Microsoft Task Scheduler task to multiple users participating on a network.", "7480", "1/4/2001 6:48:00 PM"
"66.ASP", "Q218/0/66.ASP", "Q218066 - The Server Status Tool Does Not Work Without an E-mail Client", "When the Server Status tool (also called the Value-Added Provider Reporting tool, or Vaprpt.exe), tries to send e-mail notification, it   may generate the following error message:", "7637", "12/3/1999 5:30:00 PM"
"49.ASP", "Q220/1/49.ASP", "Q220149 - AT Tasks Cannot Be Viewed Using the Task Scheduler Tool", "When you schedule a task using the AT tool at a command prompt, the task appears as being scheduled in the Task Scheduler tool. When you schedule a task using the Task Scheduler tool, the task does not appear when you run the the AT tool at", "5883", "12/29/1999 5:58:00 PM"
"68.ASP", "Q221/0/68.ASP", "Q221068 - The Dr. Watson Icon Is Missing From the Taskbar", "If you use Task Scheduler to schedule the Dr. Watson tool to run automatically when you start your computer, the Dr. Watson icon may be missing from the right-side of the taskbar.", "5865", "5/11/1999 10:26:00 AM"
"70.ASP", "Q223/1/70.ASP", "Q223170 - Task Scheduler Service Must Be Started with System Account", "When you attempt to start Task Scheduler, you may receive the one of the following error messages:", "7458", "10/21/2000 6:34:00 AM"
"20.ASP", "Q224/4/20.ASP", "Q224420 - Description of the Windows Critical Update Notification Tool", "This article describes the Microsoft Windows Critical Update Notification tool that is included with Windows 98, Windows 98 Second Edition, and Windows 2000.", "11325", "12/15/2000 2:50:00 PM"
"62.ASP", "Q226/2/62.ASP", "Q226262 - Windows 2000 Does Not Have Any Default Tasks in the Scheduled Tasks Folder", "After you install Windows 2000 Professional or Windows 2000 Server, you do not have any default tasks in the Scheduled Tasks folder.", "5737", "12/29/1999 5:08:00 PM"
"70.ASP", "Q226/3/70.ASP", "Q226370 - Browsing Local Network Slow After Installing Internet Explorer", "After you install Internet Explorer version 4.x or 5 on your Microsoft Windows NT 4.0 Workstation-based computer, you may experience long wait times when trying to browse your local network. Also, you may notice a slow down when you browse ", "6961", "9/3/1999 10:54:00 AM"
"95.ASP", "Q226/7/95.ASP", "Q226795 - How to Modify a Scheduled Task to Repeat By Minutes or Hours", "The Task Scheduler tool in Control Panel does not have an option to schedule a task every", "6001", "12/29/1999 5:01:00 PM"
"77.ASP", "Q227/0/77.ASP", "Q227077 - Err Msg: An Error Occurred While Accessing Task Scheduler", "When you start the Desktop Themes tool from either Control Panel or the Start menu, you may receive the follow error message:", "6676", "6/11/1999 7:56:00 AM"
"63.ASP", "Q227/4/63.ASP", "Q227463 - Windows 2000 Disk Defragmenter Limitations", "The Windows 2000 Disk Defragmenter tool is based on Executive Softwares full retail version of Diskeeper. The version included with Windows 2000 provides limited functionality in maintaining disk performance by defragmenting volumes that u", "6206", "5/2/2000 11:19:00 PM"
"17.ASP", "Q230/5/17.ASP", "Q230517 - Scheduled Tasks May Not Run After Viewing Schedule in Item Properties", "After you install an Active Desktop item and view the schedule in the properties of that item, the items that are scheduled to run using Task Scheduler may not run when they are scheduled to run. Also, if Microsoft Windows Critical Update N", "5906", "1/4/2001 6:45:00 PM"
"03.ASP", "Q233/2/03.ASP", "Q233203 - QoS Traffic Control in Windows 2000", "Traffic control services in Windows 2000 are used to manage traffic flow for QoS-aware and non QoS-aware programs. For programs that are not QoS-aware, traffic that they generate uses the traffic control API (TCI).  This traffic is consider", "7525", "12/30/1999 3:16:00 PM"
"32.ASP", "Q234/6/32.ASP", "Q234632 - Task Scheduler Service Starts Automatically After Upgrade", "After you upgrade your Microsoft Windows NT 4.0-based computer to Windows 2000, the Task Scheduler service starts automatically, and this occurs   even if you previously disabled this feature.", "6233", "10/20/2000 7:19:00 PM"
"36.ASP", "Q235/5/36.ASP", "Q235536 - Task Scheduler Service on Windows NT", "The installation of the Offline Browsing Pack, a component of Internet Explorer 5, on a computer running Windows NT 4.0 replaces the Schedule service with the Task Scheduler service. This article explains the differences between the two ser", "7762", "10/26/2000 5:35:00 AM GMT" 
"34.ASP", "Q235/6/34.ASP", "Q235634 - Task Scheduler Add-on for Internet Explorer 4 Does Not Schedule EVERY:SUNDAY Events", "When you use the &quot;AT&quot; scheduling mechanism within Windows NT 4.0 or the WinAT tool from the Microsoft Windows NT 4.0 Resource Kit, you can schedule events to occur at any time and have them occur on a reoccurring basis (such as EVERY:SUNDAY", "5947", "11/15/2000 1:25:00 PM"
"73.ASP", "Q236/7/73.ASP", "Q236773 - Internet Explorer Replaces Atsvc.exe Tool with Mstask.exe Tool", "After you upgrade to Microsoft Internet Explorer version 5, all of   your existing Atsvc.exe jobs are converted to Scheduled Tasks.", "6759", "2/29/2000 10:54:00 PM"
"40.ASP", "Q237/8/40.ASP", "Q237840 - Soon.exe Schedules Tasks for Next Day Instead of Current Day", "When you schedule a task using the Soon tool (Soon.exe) included with the Windows NT 4.0 Resource Kit, the task may be scheduled for the next day instead of the current day.", "7231", "1/5/2001 3:12:00 PM"
"62.ASP", "Q241/1/62.ASP", "Q241162 - How to Save Backup Report Logs to an Alternate Location", "Windows 2000 Backup (Ntbackup.exe) does not have a command-line parameter to specify the location to which reports are saved after a backup operation is finished. The backup report is saved in the profiles folder of the user who performed t", "10888", "2/22/2000 11:19:00 PM"
"60.ASP", "Q243/2/60.ASP", "Q243260 - Changed Command Parameters for Scheduled Backup Job May Not Be Saved", "When you edit a previously scheduled backup job in Windows 2000 Backup, the setting changes are not saved, and Backup does not use the additional settings that you made. If you open the scheduled backup job again, the settings that you made", "8597", "4/25/2000 4:16:00 PM"
"55.ASP", "Q243/9/55.ASP", "Q243955 - How to Remove Task Scheduler from Internet Explorer 5 If the Option Is Not Available in Add/Remove Programs", "This article describes how to remove Task Scheduler if it is not listed in the Add/Remove Programs tool in Control Panel.", "6816", "10/25/1999 4:04:00 PM"
"68.ASP", "Q244/3/68.ASP", "Q244368 - How to Optimize Active Directory Replication in a Large Network", "This article describes how to optimize Active Directory replication in large network configurations.", "36824", "12/29/2000 10:50:00 AM"
"83.ASP", "Q246/1/83.ASP", "Q244368 - Error Message 0x80090016 or 0x8009000f When Trying to Schedule a Task (Q246183)", "When you attempt to schedule a task using the Scheduled Task tool in Control Panel, you may receive either of the following error messages:", "6884", "9/12/2001 3:04:00 AM"
"46.ASP", "Q246/3/46.ASP", "Q246346 - Internet Explorer 5.01 Readme.txt File", "This article contains a copy of the Readme.txt file included with Internet Explorer 5.01.", "30041", "1/4/2001 6:43:00 PM"
"72.ASP", "Q246/9/72.ASP", "Q246972 - Internet Explorer 5 Task Scheduler Allows Privilege Elevation on Windows NT", "When you use the schedule feature for updating Web pages that is included with Internet Explorer version 5 to schedule jobs to run at a designated time, it may be possible for a malicious user to obtain elevated privileges on your computer ", "7236", "12/2/1999 4:09:00 PM"
"46.ASP", "Q247/8/46.ASP", "Q247846 - Data in Remote Storage Value Becomes Inaccurate Over Time", "Over a variable period of time, the", "6896", "1/13/2000 11:50:00 PM"
"60.ASP", "Q249/8/60.ASP", "Q249860 - Kernel32 Error Message in Task Scheduler (Mstask.exe) with Terminal Server Edition", "When you use the Task Scheduler tool (Mstask.exe) that is included in Internet Explorer 5 with Microsoft Windows NT Server 4.0, Terminal Server Edition, and a scheduled task is started either manually (by using the", "6425", "11/15/2000 1:22:00 PM"
"39.ASP", "Q250/0/39.ASP", "Q250039 - Error Message: The New Task Has Been Created, But May Not Run Because the Account Information Could Not Be Set", "When you are scheduling a task with Task Scheduler, you may receive the following error message:", "7907", "1/19/2000 3:43:00 PM"
"81.ASP", "Q250/5/81.ASP", "Q250581 - Windows 2000 Keywords for Searching the Microsoft Knowledge Base", "This article lists keywords that are used in Windows 2000 articles in the Microsoft Knowledge Base to help you find articles about specific topics. Use these keywords as search words when you are researching a specific topic.", "19523", "10/18/2000 6:56:00 PM"
"19.ASP", "Q251/0/19.ASP", "Q251019 - How to Schedule VIACompact to Run on a Routine Basis", "For large Microsoft Metadirectory Services (MMS) server databases, it may be necessary to compact the database at a predetermined interval. You can do this by using scripts and one of the scheduling services included with Microsoft Windows ", "12026", "12/6/2000 11:49:00 PM"
"78.ASP", "Q257/9/78.ASP", "Q257978 - List of Issues Fixed in Internet Explorer 5.01 Service Pack 1", "This article provides a listing of article numbers for issues that were fixed in Internet Explorer 5.01 Service Pack 1. Service packs are cumulative. This means that the issues that were fixed in a service pack are also fixed in subsequent ", "13033", "1/17/2001 3:20:00 PM"
"07.ASP", "Q259/9/07.ASP", "Q259907 - Access Denied When Attempting to Start the Task Scheduler", "When you attempt to start the Task Scheduler with the Services program under Administrative Tools, you may receive an &quot;Access denied&quot; error message.", "6758", "11/15/2000 1:19:00 PM"
"01.ASP", "Q263/2/01.ASP", "Q263201 - Default Processes in Windows 2000", "This article describes the processes which run by default in Microsoft Windows 2000. These processes can be viewed using Task Manager.", "9899", "10/25/2000 11:49:00 PM"
"02.ASP", "Q268/9/02.ASP", "Q268902 - Task Scheduler and AT Command Display Inaccurate Information for Pending Jobs", "After you schedule a job by using the", "7174", "11/15/2000 1:16:00 PM"
"90.ASP", "Q269/5/90.ASP", "Q269590 - Long Delay in Opening Computer Folder if Computer Is a NetWare Server", "When you try to open a computer folder on a NetWare-based (or compatible) network, you may experience a delay of 20-30 seconds.", "7581", "1/8/2001 9:16:00 PM"
"22.ASP", "Q272/0/22.ASP", "Q272022 - Mstask.exe Does Not Run Scheduled Jobs", "Intermittently, the Mstask.exe process may stop running scheduled jobs. To cause Mstask.exe to run the jobs again, you must stop and restart the Task Scheduler service. Additionally, a job status may be displayed as &quot;Running&quot; even though th", "8084", "1/6/2001 8:01:00 AM"
"94.ASP", "Q282/4/94.ASP", "Q282494 - SMS: Courier Sender Does Not Work When Sent to Grandchild Site", "When a package is configured and sent from the central site to a grandchild site by using the courier sender, and there is no direct LAN-sender address from the central site to the grandchild site, the package distribution does not work. In", "6766", "1/19/2001 8:42:00 PM"
"60.asp", "q300/1/60.asp", "Q300160 - HOW TO: Schedule a Server Process in Windows 2000 ", "This step-by-step article describes how to schedule a program to automatically start at a pre-determined interval.", "9219", "4/9/2002 8:15:00 PM"

