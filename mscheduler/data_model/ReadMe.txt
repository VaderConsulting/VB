About
------
    SchAgent_Creation_Template.xml
	From- http://home.sprintmail.com/~mpryor/mscheduler.zip
	
Contact
--------
	Author- Mark Pryor mpryor@sprintmail.com	

Purpose
-------
    This is a TEMPLATE for JOB creation scripts. At this stage of 
    development the XML is intended
    as a guide to show you the important properties of the JOB
    and Trigger objects and how certain flags and properties pair up.
    Looking at all the Trigger Nodes, you will notice that there is one for
    each TriggerType. This works fine as we intend to enable 
    only those used-- see the TriggerFlags.

    Eventually there will be built-in or Script logic to build creation
    script from the XML, but not yet. You are free to tackle this
    yourself if you like. The basic Object Model will not change once
    I get this feature working.
    
	Configure your scheduled tasks from the a command-line
	script without using the Mstask interface to the Shell.
	Once you have correct logic in your script, you can see
	changes from the Mstask interface.
	
January 6, 2002
msp
	