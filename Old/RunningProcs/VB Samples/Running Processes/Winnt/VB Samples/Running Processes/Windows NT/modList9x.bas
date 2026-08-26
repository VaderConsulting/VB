Attribute VB_Name = "modList9x"
Option Explicit
'
' WINDOWS 9x ONLY
'
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
   ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Public Declare Function Process32First Lib "kernel32" ( _
    ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
      
Public Declare Function Process32Next Lib "kernel32" ( _
   ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Public Declare Function Module32First Lib "kernel32" ( _
    ByVal hSnapshot As Long, lpme As MODULEENTRY32) As Long
    
Public Declare Function Module32Next Lib "kernel32" ( _
    ByVal hSnapshot As Long, lpme As MODULEENTRY32) As Long
    
Private Const MAX_MODULE_NAME32 As Integer = 255
Private Const MAX_MODULE_NAME32plus As Integer = MAX_MODULE_NAME32 + 1
Private Const MAX_PATH = 260
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const TH32CS_SNAPMODULE = &H8&

Public Type PROCESSENTRY32
   dwSize               As Long 'Specifies the length, in bytes, of the structure.
   cntUsage             As Long 'Number of references to the process.
   th32ProcessID        As Long 'Identifier of the process.
   th32DefaultHeapID    As Long 'Identifier of the default heap for the process.
   th32ModuleID         As Long 'Module identifier of the process. (Associated exe)
   cntThreads           As Long 'Number of execution threads started by the process.
   th32ParentProcessID  As Long 'Identifier of the process that created the process being examined.
   pcPriClassBase       As Long 'Base priority of any threads created by this process.
   dwFlags              As Long 'Reserved; do not use.
   szExeFile            As String * MAX_PATH 'Path and filename of the executable file for the process.
End Type

Public Type MODULEENTRY32
    dwSize          As Long 'Specifies the length, in bytes, of the structure.
    th32ModuleID    As Long 'Module identifier in the context of the owning process.
    th32ProcessID   As Long 'Identifier of the process being examined.
    GlblcntUsage    As Long 'Global usage count on the module.
    ProccntUsage    As Long 'Module usage count in the context of the owning process.
    modBaseAddr     As Long 'Base address of the module in the context of the owning process.
    modBaseSize     As Long 'Size, in bytes, of the module.
    hModule         As Long 'Handle to the module in the context of the owning process.
    szModule        As String * MAX_MODULE_NAME32plus 'String containing the module name.
    szExePath       As String * MAX_PATH 'String containing the location (path) of the module.
End Type

Public Function StrZToStr(s As String) As String
   StrZToStr = Left$(s, Len(s) - 1)
End Function

Public Function FillProcessList9x(lstBox As ListBox) As Long
'
' Clears the listbox and fills it with the
' processes and the modules used by each process.
'
Dim lReturnID     As Long
Dim hSnapProcess  As Long
Dim hSnapModule   As Long
Dim sName         As String
Dim proc          As PROCESSENTRY32
Dim module        As MODULEENTRY32
Dim iProcesses    As Integer
Dim iModules      As Integer
'
' Clear the listbox.
'
lstBox.Clear
'
'Get a snapshot of all the processes.
'
hSnapProcess = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    
If hSnapProcess = hNull Then
    '
    ' If the snapshot is empty, then exit.
    ' Return the number of Processes found.
    '
    FillProcessList9x = 0
Else
    '
    ' Initialize the processentry structure.
    '
    proc.dwSize = Len(proc)
    '
    ' Get first process.
    '
    lReturnID = Process32First(hSnapProcess, proc)
    '
    ' Iterate through each process with an ID that <> 0.
    '
    Do While lReturnID
        '
        ' Add the process to the listbox.
        '
        lstBox.AddItem StrZToStr(proc.szExeFile)
        '
        ' Increment the count of processes we've added.
        '
        iProcesses = iProcesses + 1
        '
        ' Get a snapshot of all the modules in this process.
        '
        hSnapModule = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, proc.th32ProcessID)
        '
        ' If the process has modules loaded, iterate through them.
        '
        If Not hSnapModule = hNull Then
            '
            ' Initialize the moduleentry structure.
            '
            module.dwSize = LenB(module) - 1
            '
            ' Get first module.
            '
            lReturnID = Module32First(hSnapModule, module)
            '
            ' Iterate through the modules with an ID that <> 0.
            '
            Do While lReturnID
                '
                ' If there is a module, add it to the list.
                '
                lstBox.AddItem "    " & StrZToStr(module.szModule)
                '
                ' Get next module.
                '
                lReturnID = Module32Next(hSnapModule, module)
            Loop
        End If
        '
        ' Close the module snapshot handle.
        '
        Call CloseHandle(hSnapModule)
        '
        ' Get next process.
        '
        lReturnID = Process32Next(hSnapProcess, proc)
    Loop
    '
    ' Close the Process snapshot handle.
    '
    Call CloseHandle(hSnapProcess)
    '
    ' Return the number of Processes found.
    '
    FillProcessList9x = iProcesses
End If
End Function
