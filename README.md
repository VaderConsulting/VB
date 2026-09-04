# VB

VB is Dave Robinson's Visual Basic 6 working folder from OneDrive Historical Dev: Windows IT-admin utilities (ADSI ACL editors, PC audit, bandwidth monitor, NAT discovery, GPO HTML/XML dumps) plus samples and a large `Old/` archive of earlier VB6 projects. A handful of VB.NET solutions sit under `Old/DOT NET` (Balloon, TaskbarApp, WindowsService1, WinForms-Calc). Password-related tools `AdminSetPassword` and `AccessPwdBreak` are included with usernames, emails, and passwords redacted.

**Source last updated:** 2019-02-10 · **Language:** VB6 / VB.NET · **Target:** VB6 Win32 (most projects) / VS.NET 2002 `.NET` (Old/DOT NET) · **Output:** WinForms exes, ActiveX OCX/DLL, GPO HTML/XML dumps

## Solution structure

There is no single `.sln` for the folder. Notable projects on disk:

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `AccessPwdBreak` | VB6 | WinForms exe (`Breaker.exe`) | Third-party "KNR's Access 97 Password Breaker" (`VersionCompanyName` smart software). Sample `Dummy.mdb` is gitignored. Author email redacted. |
| `ACL` / `ACL2` / `ACL3` | VB6 | WinForms exe | ADSI / file ACL experiments (`ADS_RIGHTS_ENUM`, `IADsSecurityDescriptor`). ACL2 writes a test ACE; ACL3 uses Win32 memory APIs for DACL work. |
| `ADMExport` | VB6 | WinForms exe (`prjADMExport.exe`) | Exports ADM (Group Policy administrative template) files. |
| `AdminSetPassword` | VB6 | WinForms exe | Domain admin password-set utility (Tusk/Freelance). Reads `setup.ini` and `userlist.txt` - originals gitignored; use the `.example` files. |
| `Audit` | VB6 | WinForms exe (`Audit.exe`) | PC Audit: computer name, IP/MAC, OS, hotfixes, installed apps, local accounts, services. Includes a sample `LAPTOP.txt` run. |
| `Bandwidth` | VB6 | WinForms exe (`CS Bandwidth Monitor.exe`) | Systray/desktop bandwidth monitor (IP Helper / netstat). |
| `Barcodes` | VB6 | WinForms exe (`Barcode.exe`) | Code 39 barcode generator (Allen Allegretto). |
| `DomainInfo` | VB6 | WinForms exe (`DomInfo.exe`) | Domain Info. |
| `EWoW Protocol` | VB6 | WinForms exe | EWoW-C protocol tester (Realtag). |
| `GPOAudit` | HTML/XML | document dump | Exported Group Policy reports (Citrix baselines, domain controllers, site/user OU policies). No `.vbp`. |
| `LoggedOnUsers` | VB6 | WinForms exe (`LoggedOn.exe`) | Lists logged-on users. |
| `Logon34` | VB6 | WinForms exe (`Logon34.exe`) | Logon UI (`VersionCompanyName` Computer Sciences Corporation and Tusk Technologies). |
| `Nat` | VB6 | several WinForms exes | NAT suite by Dave Robinson: Discover, Discover2, Broker, Ping, Collect, Reader. |
| `QueryTool` | VB6 | WinForms exe | Policy Query Tool (`VersionCompanyName` Rio Tinto). |
| `ServiceMan` | VB6 | WinForms exe | Windows service manager. |
| `USB` | VB6 | WinForms exe (`Media.exe`) | Media Detect. Nested `usb.zip` also present. |
| `Wol` | VB6 | WinForms exe (`MiniDisc.exe`) | Mini Discover (Dave Robinson). |
| `VB Accelerator` | VB6 | OCX/DLL/exe samples | Steve McMahon vbAccelerator controls (S-Grid, Image List, List Bar, popup menu, SysTray, journal hook, icon extractor). |
| `xml0800` | VB6 | WinForms exe (`SaxTest.exe`) | Microsoft SAX Workbench. |
| `Old` | VB6 / VB.NET | archive | ~199 earlier projects (NT admin, SMS design docs, NWN tools, audio, personnel, etc.). `Old/NWN Resource Viewer` was not in the OneDrive zip. VB.NET: Balloon, TaskbarApp, WindowsService1, WinForms-Calc. |

Other top-level VB6 projects on disk (open their `.vbp`): `CaptureMouseEvents`, `cbGPS`, `Conmon`, `FileSplitter`, `FileTransfer`, `findandreplace`, `Fixurl`, `GUID Creator`, `Guiddll`, `Helper`, `Idle Time`, `ini_editor`, `INIEditor`, `inifiletotreeview`, `Integration`, `Lasertag`, `Linker`, `Mgecomp`, `MigrateProfile`, `ModGroups`, `ModifyEA`, `MouseHook`, `MP4 Display`, `MXXMLWriterSample`, `MyMonitor`, `Mywsh`, `Netinfo`, `netserv`, `NewConn`, `Nicinfo`, `Ntinfo`, `PakConfig`, `Photos`, `PICPic`, `PrinterTest`, `Regwrite`, `ReleaseRenewIP`, `RemText`, `Repl20Code`, `ReplaceData`, `ReplaceLines`, `RunRemote`, `schedule`, `ScriptingDemo`, `SecSearch`, `SendKeys`, `setdefprint`, `SetLibrary`, `SetRegPerms`, `ShellExecute`, `Shelllnk`, `SwapLibrary`, `SwapServer`, `Tcp`, `TCPClient`, `TCPServer`, `Temp`, `Thumbs`, `ttclass`, `UpdateNetwork`, `UpdateTrendGUID`, `Vb4kix`, `VBsax2jumpstart`, `Wait`, `watchdir`, `WinLogon Notification`, `Word`, `WordCounter`, `XP Style`.

## How to open

There is no folder-level solution. Open a project's `.vbp` in Visual Basic 6, for example:

- `AdminSetPassword/AdminSetPassword.vbp`
- `Audit/Audit.vbp`
- `Bandwidth/Bandwidth.vbp`
- `ACL3/ACL3.vbp`
- `Nat/Discover/Discover.vbp`
- `QueryTool/QueryTool.vbp`
- `AccessPwdBreak/Project1.vbp`

VB.NET (Visual Studio .NET 2002, solution format 7.00):

- `Old/DOT NET/Balloon.sln`
- `Old/DOT NET/TaskbarApp/TaskbarApp.sln`
- `Old/DOT NET/WindowsService1/WindowsService1.sln`
- `Old/DOT NET/Samples/Vb.net/WinForms-Calc/Calc/Calc.sln`

## Requirements

- Visual Basic 6.0 IDE

## Attribution and provenance

Working copy from Dave Robinson's OneDrive Historical Dev folder `VB`. Company names in `.vbp` files include Freelance, Dave Robinson, Empired Limited, Realtag, Rio Tinto, Computer Sciences Corporation and Tusk Technologies, plus third-party vendors listed in `THIRD_PARTY_NOTICES.md`. OneDrive skipped `VB/Old` as a folder item and `VB/Old/NWN Resource Viewer`; other `Old/` projects did extract.

## License

MIT © 2026 VaderConsulting for Dave Robinson's code. See `LICENSE`. Third-party trees (vbAccelerator, KNR AccessPwdBreak, Microsoft SAX samples, Edanmo, TheScarms, and others) keep their original terms; see `THIRD_PARTY_NOTICES.md`.
