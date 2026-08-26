The first thing to note when enumerating the processes of the operating system is that the necessary API functions are completely different under Windows 95/98 and Windows NT. Under Windows 95/98, functions from the ToolHelp32 group of APIs are used. Under Windows NT, functions from PSAPI.DLL are used.

Windows 95/98/2000 and ToolHelp32
----------------------------

The Toolhelp32 APIs used under Windows 95 and Windows 98 reside in the KERNEL32.DLL. These API functions are available only under Windows 95 and Windows 98. The following ToolHelp32 functions allow you to enumerate processes in the system, as well as get memory and module information:

   CreateToolhelp32Snapshot() 
   Process32First() 
   Process32Next() 
   Module32First() 
   Module32Next()

The first step is to create a "snapshot" of the information in the system using the CreateToolhelp32Snapshot() function. This function allows you to choose what type of information is stored in the snapshot. The ModList sample initially specifies the TH32CS_SNAPPROCESS flag because we are interested in process information. This function returns a handle to a PROCESSENTRY32 structure, and it is important to remember to pass the handle to CloseHandle() when processing is complete.
 
To iterate through the list of processes in the snapshot, call Process32First once, followed by repeated calls to Process32Next, until one of these functions returns FALSE. Both of these functions take the handle to the snapshot and a pointer to a PROCESSENTRY32 structure as parameters. Process32First and Process32Next fill a PROCESSENTRY32 structure with useful information about a process in the system. 

The process ID is in the th32ProcessID member of the structure. The process' executable file and path are stored in the szExeFile member of the structure. Other useful information is also available in the structure. The ModList sample only retrieves the EXE name and adds it to a listbox. The process ID can be passed to the OpenProcess() API to get a handle to the process. 

With a valid process id, we can again call the CreateToolHelp32Snapshot() to retrieve module information for the process. This secondary call to CreateToolHelp32Snapshot() passes the TH32CS_SNAPMODULE flag and a pointer to the process ID (th32ProcessID value of the PROCESSENTRY32 structure) as the parameters, and creates a snapshot of the modules in that process. 

Again, this function returns a HANDLE and it is important to remember to close the handle, using CloseHandle(), after retrieving information about the process.

In a manner similar to retrieving process information, Module32First is called once and Module32Next is called as many times as required to iterate through the modules information for the process. 

Windows NT and the PSAPI.DLL
----------------------------

The Windows NT approach to creating a list of processes and modules uses functions from the PSAPI.DLL. The PSAPI.DLL file is distributed with the Platform SDK, available at http://www.microsoft.com/msdn/sdk. 

Like the ToolHelp32 functions, the PSAPI.DLL also contains a variety of useful functions. However, this article only discusses those functions relevant to enumerating processes and modules: 

   EnumProcesses()
   EnumProcessModules()
   GetModuleFileNameExA()

First, a call is made to EnumProcesses() to fill an array of process IDs. The ModList sample code also includes a method of calculating the number of processes returned. 

Next, OpenProcess() is called for each of the process IDs to retrieve a handle to the process, if the handle is valid, then call EnumProcessModules() to enumerate the modules of the process. EnumProcessModules() fills an array passed as a parameter, with the 

module handles associated with the process. 

GetModuleFileNameExA() is used to retrieve the name of the module using the process handle and module handle as parameters. The module name would be the path and file name of the dll, ocx, etc. that the process has loaded. 

In ModList, a module name is indented in the listbox to show it to be a "child" under the process it is associated with.

Additional notes
----------------

The name of a process may also display in the list of modules for that process. If this is not desirable, simply compare the module name to the process name before adding it to the list.

In Windows 95, 16-bit applications have process IDs etc., just like Win32 applications; therefore they are reported equally by Toolhelp32 functions. However, this is not the case under Windows NT. 16-Bit applications running under Windows NT run in what is called a Virtual Dos Machine (VDM). EnumProcesses will not recognize any 16-bit applications in the system, however it will return the 32-bit NTVDM processes under which the 16-bit 
applications are running. To enumerate 16-bit applications under Windows NT you must use a function called VDMEnumTaskWOWEx(). 