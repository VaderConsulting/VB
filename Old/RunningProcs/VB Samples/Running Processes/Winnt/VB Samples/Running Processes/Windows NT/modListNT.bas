Attribute VB_Name = "modListNT"
Option Explicit
'
' WINDOWS NT ONLY
'
' Remember to distribute the PSAPI.DLL file.
'
Public Declare Function CloseHandle Lib "Kernel32.dll" _
   (ByVal Handle As Long) As Long

Public Declare Function OpenProcess Lib "Kernel32.dll" _
  (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
      ByVal dwProcId As Long) As Long

Public Declare Function EnumProcesses Lib "psapi.dll" _
   (ByRef lpidProcess As Long, ByVal cb As Long, _
      ByRef cbNeeded As Long) As Long

Public Declare Function GetModuleFileNameExA Lib "psapi.dll" _
   (ByVal hProcess As Long, ByVal hModule As Long, _
      ByVal ModuleName As String, ByVal nSize As Long) As Long

Public Declare Function EnumProcessModules Lib "psapi.dll" _
   (ByVal hProcess As Long, ByRef lphModule As Long, _
      ByVal cb As Long, ByRef cbNeeded As Long) As Long

Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Public Const PROCESS_ALL_ACCESS = &H1F0FFF


Public Function FillProcessListNT(lstBox As ListBox) As Long
'
' Clears the listbox and fill it with the
' processes and the modules used by each process.
'
Dim cb                As Long
Dim cbNeeded          As Long
Dim NumElements       As Long
Dim ProcessIDs()      As Long
Dim cbNeeded2         As Long
Dim NumElements2      As Long
Dim Modules(1 To 200) As Long
Dim lRet              As Long
Dim ModuleName        As String
Dim nSize             As Long
Dim hProcess          As Long
Dim i                 As Long
Dim sModName          As String
Dim sChildModName     As String
Dim iModDlls          As Long
Dim iProcesses        As Integer
    
lstBox.Clear
'
' Get the array containing the process id's for each process object.
'
cb = 8
cbNeeded = 96
'
' There is no way to find out how big the passed in array
' must be. EnumProcesses() will never return a value in
' cbNeeded that is larger than the size of array value
' that you passed in the cb parameter.
'
' If cbNeeded == cb upon return, allocate a larger array
' and try again until cbNeeded is smaller than cb.
'
Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
Loop
'
' Calculate how many process IDs were returned.
'
NumElements = cbNeeded / 4
    
For i = 1 To NumElements
    '
    ' Get a handle to the Process.
    '
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
            Or PROCESS_VM_READ, 0, ProcessIDs(i))
    '
    ' Iterate through each process with an ID that <> 0.
    '
    If hProcess Then
        '
        ' Get an array of the module handles for the specified process.
        '
        lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
        '
        ' If the Module Array is retrieved, Get the ModuleFileName.
        '
        If lRet <> 0 Then
            '
            ' Fill the ModuleName buffer with spaces.
            '
            ModuleName = Space(MAX_PATH)
            '
            ' Preset buffer size.
            '
            nSize = 500
            '
            ' Get the module file name.
            '
            lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
            '
            ' Get the module file name out of the buffer, lRet is how
            ' many characters the string is, the rest of the buffer is spaces.
            '
            sModName = Left$(ModuleName, lRet)
            '
            ' Add the process to the listbox.
            '
            lstBox.AddItem sModName
            '
            ' Increment the count of processes we've added.
            '
            iProcesses = iProcesses + 1
                
            iModDlls = 1
            Do
                iModDlls = iModDlls + 1
                '
                ' Fill the ModuleName buffer with spaces.
                '
                ModuleName = Space(MAX_PATH)
                '
                ' Preset buffer size.
                '
                nSize = 500
                '
                ' Get the module file name out of the buffer, lRet is how
                ' many characters the string is, the rest of the buffer is spaces.
                '
                lRet = GetModuleFileNameExA(hProcess, Modules(iModDlls), ModuleName, nSize)
                sChildModName = Left$(ModuleName, lRet)
                    
                If sChildModName = sModName Then Exit Do
                If Trim(sChildModName) <> "" Then lstBox.AddItem "    " & sChildModName
            Loop
        End If
    Else
        '
        ' Return the number of Processes found.
        '
        FillProcessListNT = 0
    End If
    '
    ' Close the handle to the process.
    '
    lRet = CloseHandle(hProcess)
Next
'
' Return the number of Processes found.
'
FillProcessListNT = iProcesses
End Function



