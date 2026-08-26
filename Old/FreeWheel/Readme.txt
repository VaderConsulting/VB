FreeWheel 2.0
-------------

FreeWheel README.TXT (June 28, 1997)


What is FreeWheel?
------------------

FreeWheel provides mouse wheel support for applications which do not directly 
support the IntelliPoint mouse wheel. It runs as a tray icon, which you can 
hide if you dislike clutter in your tray. Windows NT 3.51 users get a 
minimized program icon instead of a tray icon. FreeWheel should work on 
Windows NT 3.51 but it has not been properly tested on this platform. 
FreeWheel even runs on Win32s 1.3 but this is completely pointless, as there 
are no IntelliPoint drivers for Windows 3.x.

FreeWheel supports the 'default' mode of operation, where rotating the wheel 
causes the active (focus) window to scroll, and an alternative mode where the 
window directly underneath the mouse pointer is scrolled. You can switch 
between documents in an application by holding down Ctrl+Shift while rotating 
the mouse wheel, and you can switch between applications by holding down Alt 
and rotating the wheel.

FreeWheel caters for the possibility of conflicts with applications that 
directly support the wheel by allowing you to specify a list of applications. 
FreeWheel can either include or exclude applications which appear in this list.

FreeWheel is free, so there are no irritating nag screens, no timeouts, and 
no guilt associated with failing to register!


Installing FreeWheel
--------------------

There are two versions of FreeWheel. The compact version weighs in at 
49KB, and requires MSVCRT40.DLL and MFC40.DLL. These libraries are fairly 
ubiquitous and will hopefully be available from the same place that you 
got FreeWheel. If not, there is a stand-alone version which is somewhat 
larger at about 172KB.

Run Setup.exe to install FreeWheel to the folder of your choice. When the setup 
program first runs FreeWheel, a shortcut will be created in your StartUp menu 
so that FreeWheel will run automatically each time you start Windows.

Using FreeWheel
---------------

FreeWheel should start up initially with some sensible default settings. 
To customize the settings, double-click on its tray icon, or select 
'Properties' from its context menu. You will be presented with a dialog 
containing several property pages.

The 'General' page reports information about your Windows version and 
the level of mouse wheel support that your system has.

The 'Settings' page allows you to the way FreeWheel works. You can choose 
whether to scroll the focus window or the window under the mouse pointer. 
You can enable alternate scrolling, where holding down Ctrl+Alt temporarily 
causes page scrolling if the default is line scrolling, and vice versa. 
You may also hide FreeWheel's tray icon.

The 'Switching' page allows to enable or disable FreeWheel's window switching 
features.

The 'Compatibility' page allows you to disable FreeWheel's scrolling 
support in selected applications. Alternatively you can disable FreeWheel 
for all applications except those that you specify.

The 'About' page contains a pretty icon and some blurb.

FreeWheel does not insert any extra pages in the system 'Mouse Properties' 
dialog, as it is just as easy to get to the FreeWheel's properties from its
tray icon.


Support
-------

FreeWheel is free software so, as such, there is no support. 
Either it works for you, or it doesn't. However, if you have any bug 
reports or suggestions, I will be happy to hear about them.


Compatibility
-------------

FreeWheel works fine with most applications. One 'problem app' seems 
to be the Windows 95 Explorer (though not the NT Explorer, curiously), 
where each notch on the wheel scrolls twice as many lines as it should, 
or both panes scroll at once! This can be cured by putting Explorer 
(e.g. c:\windows\explorer.exe) into the exclusion list in the 
'Compatibility' page. Please feel free to report any other applications 
which are problematic.


Uninstalling FreeWheel
----------------------

Launch the "Add/Remove Programs" icon from the Control Panel and select 
FreeWheel from the list of applications that can be removed.


History
-------

Date          | Version | Comments
-----------------------------------------------------------
Jun 28, 1997  | 2.0     | Major enhancements in scroll code
              |         | Improved compatibility with apps, including Word97
              |         | Added horizontal scrolling
              |         | Added combo box scrolling (Win95 only)
              |         | Added Alt+Wheel app switching
              |         | Added 'hide tray icon' feature
Jun 13, 1997  | 1.2     | Fixed silly bug in 'scroll window under mouse' code
Jun 12, 1997  | 1.1     | Added installer/uninstaller
Jun 09, 1997  | 1.0     | Proof of pudding


Trademarks
----------

Microsoft is a registered trademark, and IntelliMouse is a trademark
of Microsoft Corp. But you knew that already.


License Agreement
-----------------

FreeWheel is free software. That means that my liability shall not exceed
the amount paid to me for use of the software!


Updated June 28th, 1997

Jim Barry <Jim.Barry@ilp.com>
