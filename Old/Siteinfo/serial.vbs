ComputerName = InputBox("Enter the name of the computer you wish to query")

winmgmt1 = "winmgmts:{impersonationLevel=impersonate}!//"& ComputerName &""

'WScript.Echo winmgmt1

Set SNSet = GetObject( winmgmt1 ).InstancesOf ("Win32_BIOS")

for each SN in SNSet
	MsgBox "The serial number for the specified computer is: " & SN.SerialNumber
Next

