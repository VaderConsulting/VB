	Set objConn = WScript.CreateObject("ADODB.Connection")
	objConn.Open "<<database>>", "<<UserID>>", "<<password>>"
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	SQLStmt = "Select distinct Site, Server from ExchangeUpload"
	Set objSites = WScript.CreateObject("ADODB.Recordset")
	objSites.Open SQLStmt, objConn, 0, 1, 1
	Set objBatFile = objFileSystem.CreateTextFile("ExImport.cmd",True)
		
	Do While Not objSites.EOF
		objBatFile.WriteLine("admin /i " & objSites.Fields("Site") & ".csv /d " & objSites.Fields("Server") & " /n")
		Set objFile = objFileSystem.CreateTextFile(objSites.Fields("Site") & ".csv",True)
		objFile.WriteLine("Obj-Class,Directory Name,Display Name,Employee Number,Manager,Cost Center,Department,Title,Office,Company,Phone number,Fax number,Obj-Container,Hide From AB")
		
		SQLStmt = "Select * from ExchangeUpload where Site='" & objSites.Fields("Site") & "' Order by Container, DisplayName"
		Set objRst = WScript.CreateObject("ADODB.Recordset")
		objRst.Open SQLStmt, objConn, 0, 1, 1
		
		Do While Not objRst.EOF
			outText = objRst.Fields("ObjClass") & "," 
			outText = outText & chr(34) & objRst.Fields("DirectoryName") & chr(34) & ","
			outText = outText & chr(34) & objRst.Fields("DisplayName") & chr(34) & ","
			outText = outText & objRst.Fields("EmployeeNumber") & ","
			outText = outText & objRst.Fields("Manager") & ","
			outText = outText & chr(34) & objRst.Fields("CostCenter") & chr(34) & ","
			outText = outText & chr(34) & objRst.Fields("Department") & chr(34) & ","
			outText = outText & chr(34) & objRst.Fields("Title") & chr(34) & ","
			outText = outText & chr(34) & objRst.Fields("Office") & chr(34) & ","
			outText = outText & chr(34) & "Glenayre Electronics" & chr(34) & ","
			if (Right(objRst.Fields("Phone"),4) = Left(objRst.Fields("Extension"),4)) OR (UCase(Left(objRst.Fields("Extension"),4)) = "NONE") Then
				pPhone = Trim(objRst.Fields("Phone"))
			else
				pPhone = Trim(objRst.Fields("Phone")) & " Ext. " & Trim(objRst.Fields("Extension"))
			end if
			outText = outText & chr(34) & pPhone & chr(34) & ","
			outText = outText & chr(34) & Left(Trim(objRst.Fields("Extension")),4) & chr(34) & ","
			outText = outText & chr(34) & objRst.Fields("ObjContainer") & chr(34) & ","
			outText = outText & objRst.Fields("HideFromAB")
			objFile.WriteLine(outText)
			objRst.MoveNext
		Loop
		
		objFile.Close
		objRst.Close
		Set objRst = Nothing	
		Set objFile = Nothing
	
		objSites.MoveNext
	Loop
	
	objBatFile.Close
	objSites.Close
	Set objSites = Nothing
	objConn.Close
	Set objConn = Nothing
	Set objFileSystem = Nothing