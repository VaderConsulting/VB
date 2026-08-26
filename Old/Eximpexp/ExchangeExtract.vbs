const adBSTR = 8
const adParamInput = 1


const CdoPR_ADDRTYPE                             = &H3002001E
const CdoPR_EMAIL_ADDRESS                        = &H3003001E
const CdoPR_GIVEN_NAME                           = &H3A06001E
const CdoPR_INITIALS                             = &H3A0A001E
const CdoPR_SURNAME                              = &H3A11001E
const CdoPR_TITLE                                = &H3A17001E
const CdoPR_DEPARTMENT_NAME                      = &H3A18001E
const CdoPR_MANAGER_NAME                         = &H3A4E001E
const PR_EMS_AB_HIERARCHY_PATH                   = &HFFF9001E
const PR_EMS_AB_OBJ_DIST_NAME                    = &H803C001E
const PR_EMS_AB_EXTENSION_ATTRIBUTE_1            = &H802D001E
const PR_EMS_AB_EXTENSION_ATTRIBUTE_2            = &H802E001E
const PR_EMS_AB_EXTENSION_ATTRIBUTE_9            = &H8035001E
const PR_EMS_AB_ASSOC_NT_ACCOUNT                 = &H80270102
const PR_LAST_MODIFICATION_TIME                  = &H30080040
const PR_EMS_AB_USN_CHANGED                      = &H80290003
const PR_EMS_AB_MANAGER                          = &H8005000D
const PR_EMS_AB_MANAGER_T                        = &H8005001E
const PR_EMS_AB_ALIASED_OBJECT_NAME              = &H804D001E
const PR_EMS_AB_DISPLAY_NAME_PRINTABLE           = &H39FF001E
const PR_EMS_AD_DISPLAY_NAME_OVERRIDE            = &H8001000B
const PR_EMS_AB_HIDE_FROM_ADDRESS_BOOK           = &H80B9000B
const ActMsgPR_EMS_AB_PROXY_ADDRESSES            = &H800F101E




	function OrgFromHierarchyPath(txtString)
		Dim token1 
		Dim token2 
		Dim result
		
		token1 = InStr(1,txtString,"\")
		token2 = InStr(token1+1,txtString,"\")
		if token2 > 0 Then
			result = Mid(txtString,token1+1, token2-(token1+1))
		else
			result = ""
		end if
		
		OrgFromHierarchyPath = result           
	end function
	
	function SiteFromHierarchyPath(txtString)
		Dim token1 
		Dim token2 
		Dim token3
		Dim result 
		
		token1 = InStr(1,txtString,"\")
		token2 = InStr(token1+1,txtString,"\")
		token3 = InStr(token2+1,txtString,"\")
		
		if token3 > 0 Then
			result = Mid(txtString,token2+1, token3-(token2+1))
		else
			result = ""
		end if
		
		SiteFromHierarchyPath = result
	end function
	
	function ContainerFromHierarchyPath(txtString) 
		Dim token1 
		Dim token2
		Dim token3
		Dim result
		
		token1 = InStr(1,txtString,"\")
		token2 = InStr(token1+1,txtString,"\")
		token3 = InStr(token2+1,txtString,"\")

		if token3 > 0 Then
			result = Right(txtString,Len(txtString)-token3)
		else
			result = ""
		end if
		
		ContainerFromHierarchyPath = result
	end function

	Dim vbLF
	Dim strUser
	Dim strPassword
	Dim strServer
	Dim strMailbox
	Dim strProfileInfo
	
	Dim objSession
	Dim objInbox
	
	Dim pObjClass
	
	vbLF = chr(10)
	strUser="ssmith"
	strPassword="<<password>>"
	strServer="<<servername>>"
	strMailbox="<<mailboxname>>"
	
	strProfileInfo = strServer + vbLF + strMailbox
	
	Set objSession = WScript.CreateObject("MAPI.Session")
	objSession.Logon "", "", False, True, 0, True, strProfileInfo
	'Set objInbox = objSession.Inbox
	
       'On Error Resume Next

	' This should occur at the top of the script as well.  If the AM session does not
	' exist and the URL does not contain server,mailbox & alias data, a logon form will be
	' emited and returned for the user to fill out.
	'CheckAMSession
	' After calling CheckAMSession, the AM session can be retrieved from 
	' Session("AMSession").  CheckAMSession may have failed to create a session, though,
	' so check whether objAMSession is Nothing and if so cleanup, inform the user, etc.
	set objAMSession= objSession

	if objAMSession Is Nothing Then
		' CheckSession was unable to retrieve or create a session
		Wscript.Echo "GetAMSession returned nothing!"
	End If
	
	'Get the pointer to the list of Exchange Address Lists
	Set AddrLists = objAMSession.AddressLists
	
	'Get connection to SQL Server database
	Set objConn = WScript.CreateObject("ADODB.Connection")
	objConn.Open "UserInformation", "<<username>>", "<<password>>"
	'Delete existing SQL Server Exchange data.  Data is old and this prevents key violations.
				SQLStmt = "Insert into ExchangeMailboxes "
				SQLStmt = SQLStmt & "(ObjDistinguishedName,"
				SQLStmt = SQLStmt & "Alias,"
				SQLStmt = SQLStmt & "DisplayName,"
				SQLStmt = SQLStmt & "Organization,"
				SQLStmt = SQLStmt & "Site,"
				SQLStmt = SQLStmt & "Container,"
				SQLStmt = SQLStmt & "FirstName,"
				SQLStmt = SQLStmt & "MiddleInitial,"
				SQLStmt = SQLStmt & "LastName,"
				SQLStmt = SQLStmt & "EmployeeNumber,"
				SQLStmt = SQLStmt & "CostCenter,"
				SQLStmt = SQLStmt & "EMailAddress,"
				SQLStmt = SQLStmt & "ObjClass,"
				SQLStmt = SQLStmt & "PrimaryWindowsNTAccount,"
				SQLStmt = SQLStmt & "Title,"
				SQLStmt = SQLStmt & "Department,"
				SQLStmt = SQLStmt & "Manager,"
				SQLStmt = SQLStmt & "ManagerEmployeeNumber,"
				SQLStmt = SQLStmt & "LastModifiedDate,"
				SQLStmt = SQLStmt & "ImportDate,"
				SQLStmt = SQLStmt & "HideFromAB)"
				SQLStmt = SQLStmt & " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
				
	objConn.Execute("Delete from ExchangeContainers")
	objConn.Execute("Delete from ExchangeMailboxes")
	
	Set objCm = WScript.CreateObject("ADODB.Command")
	Set objCm.ActiveConnection = objConn
	objCm.CommandText = SQLStmt
	objCm.Prepared = True
	objCm.Parameters.Append objCm.CreateParameter("ObjDistinguishedName",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("Alias",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("DisplayName",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("Organization",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("Site",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("Container",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("FirstName",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("MiddleInitial",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("LastName",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("EmployeeNumber",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("CostCenter",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("EMailAddress",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("ObjClass",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("PrimaryWindowsNTAccount",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("Title",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("Department",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("Manager",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("ManagerEmployeeNumber",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("LastModifiedDate",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("ImportDate",adBSTR,adParamInput)
	objCm.Parameters.Append objCm.CreateParameter("HideFromAB",adBSTR,adParamInput)
	
	SQLStmt = "Insert into ExchangeContainers (ObjDistinguishedName,Organization, Site, Container) "
	SQLStmt = SQLStmt & "values (?,?,?,?)"

	Set objCm2 = WScript.CreateObject("ADODB.Command")
	Set objCm2.ActiveConnection = objConn
	objCm2.CommandText=SQLStmt
	objCm2.Prepared = True
	objCm2.Parameters.Append objCm2.CreateParameter("ObjDistinguishedName",adBSTR,adParamInput)
	objCm2.Parameters.Append objCm2.CreateParameter("Organization",adBSTR,adParamInput)
	objCm2.Parameters.Append objCm2.CreateParameter("Site",adBSTR,adParamInput)
	objCm2.Parameters.Append objCm2.CreateParameter("Container",adBSTR,adParamInput)
	
	'NOTE: Starting at 2 SKIPS the global Address list
	for counter = 2 to AddrLists.Count
		On Error Resume Next
		'Get Address List Information
		pOrg = OrgFromHierarchyPath(AddrLists(counter).Fields(PR_EMS_AB_HIERARCHY_PATH))
		pSite = SiteFromHierarchyPath(AddrLists(counter).Fields(PR_EMS_AB_HIERARCHY_PATH))
		pContainer = ContainerFromHierarchyPath(AddrLists(counter).Fields(PR_EMS_AB_HIERARCHY_PATH))
		pDistName = AddrLists(counter).Fields(PR_EMS_AB_OBJ_DIST_NAME)
		'Update SQL Server's ExchangeContainers tables with Address List information
		On Error Goto 0
		objCm2("ObjDistinguishedName") = pDistName
		objCm2("Organization") = pOrg
		objCm2("Site") = pSite
		objCm2("Container") = pContainer
		objCm2.Execute

		WScript.Echo AddrLists(counter).Fields(PR_EMS_AB_HIERARCHY_PATH) & " - " & counter 
			'Get Address Entry Information (Users)
			Set AddrEntries = AddrLists(counter).AddressEntries
			for count = 1 to AddrEntries.Count
				'===============================
				' Initialize Variables to Nothing
				pDisplayName=""
				pAlias=""
				pHideFromAB=""
				pDistName=""
				pFirstName=""
				pMiddleInitial=""
				pLastName=""
				pEmpNum=""
				pCostCenter=""
				pSMTPAddress=""
				pObjClass=0
				pNTAccount=""
				pTitle=""
				pDepartment=""
				pManager=""
				pManagerEmpID = ""
				pModifyDate=Null
				'==============================
				' Get values from Exchange Directory
				On Error Resume Next
				pDisplayName = AddrEntries(count).Name
				pObjClass = AddrEntries(count).DisplayType
				pDistName = AddrEntries(count).Fields(PR_EMS_AB_OBJ_DIST_NAME)
				pAlias = AddrEntries(count).Fields(PR_EMS_AB_DISPLAY_NAME_PRINTABLE)
				pHideFromAB = AddrEntries(count).Fields(PR_EMS_AB_HIDE_FROM_ADDRESS_BOOK)
				pFirstName = AddrEntries(count).Fields(CdoPR_GIVEN_NAME)
				pMiddleInitial = AddrEntries(count).Fields(CdoPR_INITIALS)
				pLastName = AddrEntries(count).Fields(CdoPR_SURNAME)
				pEmpNum = AddrEntries(count).Fields(PR_EMS_AB_EXTENSION_ATTRIBUTE_1)
				pCostCenter = AddrEntries(count).Fields(PR_EMS_AB_EXTENSION_ATTRIBUTE_2)
				EMailAddresses = AddrEntries(count).Fields(ActMsgPR_EMS_AB_PROXY_ADDRESSES)
				proxyCount = UBound(EMailAddresses)
				for proxycounter = LBound(EMailAddresses) to proxyCount
					if Left(EMailAddresses(proxycounter),5) = "SMTP:" Then
						pSMTPAddress = EMailAddresses(proxycounter)
					end if
				next
				'On Error Goto 0
				
				'On Error Resume Next
				pNTAccount = AddrEntries(count).Fields(PR_EMS_AB_ASSOC_NT_ACCOUNT)
				pTitle = AddrEntries(count).Fields(CdoPR_TITLE)
				pDepartment = AddrEntries(count).Fields(CdoPR_DEPARTMENT_NAME)
				pManager = AddrEntries(count).Fields(CdoPR_MANAGER_NAME)
				WScript.Echo pManager
				pManager = AddrEntries(count).Manager.Fields(PR_EMS_AB_OBJ_DIST_NAME)
				pManagerEmpID = AddrEntries(count).Fields(PR_EMS_AB_EXTENSION_ATTRIBUTE_9)
				pModifyDate = AddrEntries(count).LastModified
				pModifyDate = AddrEntries(count).Fields(PR_LAST_MODIFICATION_TIME)
				'==============================
				On Error Goto 0
				' Validate data and set blanks to '*'
				if Trim(pFirstName) = "" Then
					pFirstName = "*"
				end if
				if Trim(pMiddleInitial) = "" Then
					pMiddleInitial = "*"
				end if
				if Trim(pLastName) = "" Then
					pLastName = "*"
				end if
				if Trim(pEmpNum) = "" Then
					pEmpNum = "*"
				end if
				if Trim(pCostCenter) = "" Then
					pCostCenter = "*"
				end if
				if Trim(pSMTPAddress) = "" Then
					pSMTPAddress = "*"
				end if
				if Trim(pObjClass) = "" Then
					pObjClass = "*"
				end if
				if Trim(pNTAccount) = "" Then
					pNTAccount = "*"
				end if
				if Trim(pTitle) = "" Then
					pTitle = "*"
				end if
				if Trim(pDepartment) = "" Then
					pDepartment = "*"
				end if
				if Trim(pManager) = "" Then
					pManager = "*"
				end if
				if Trim(pManagerEmpID) = "" Then
					pManagerEmpID = "*"
				end if
				'=======================================
				' Insert Exchange data in to SQL Server Table
				WScript.Echo pLastName & ", " & pFirstName & " - " & pManager 
				objConn.Errors.Clear

				objCm("ObjDistinguishedName")           = pDistName
				objCm("Alias")                          = pAlias
				objCm("DisplayName")                    = pDisplayName
				objCm("Organization")                   = pOrg
				objCm("Site")                           = pSite
				objCm("Container")                      = pContainer
				objCm("FirstName")                      = pFirstName
				objCm("MiddleInitial")                  = pMiddleInitial
				objCm("LastName")                       = pLastName
				objCm("EmployeeNumber")                 = pEmpNum
				objCm("CostCenter")                     = pCostCenter
				objCm("EMailAddress")                   = pSMTPAddress
				objCm("ObjClass")                       = CStr(pObjClass)
				objCm("PrimaryWindowsNTAccount")        = pNTAccount
				objCm("Title")                          = pTitle
				objCm("Department")                     = pDepartment
				objCm("Manager")                        = pManager
				objCm("ManagerEmployeeNumber")          = pManagerEmpID
				objCm("LastModifiedDate")               = FormatDateTime(pModifyDate,0)
				objCm("ImportDate")                     = FormatDateTime(Date,0)
				objCm("HideFromAB")                     = pHideFromAB
				
				objCm.Execute

				if objConn.Errors.Count > 0 Then
					For i = 0 to objConn.Errors.Count - 1
						WScript.Echo "Error:" & i & " - Error Code:" & objConn.Errors(i).Number & " - Error Text:" & objConn.Errors(i).Description
					next
					WScript.Echo "SQL:" & SQLStmt
					objConn.Errors.Clear
				end if
			next            
	next
	objConn.Execute("Update ExchangeMailBoxes Set HideFromAB='0'")        
	objConn.Close
	Set objConn = Nothing
	Set AddrLists = Nothing
