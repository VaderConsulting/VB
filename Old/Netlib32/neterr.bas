'Name:		NetErr.Bas
'Description:	Network API Error Handler for Windows NT
'Dependencies:	NetDecs.BAS
'------------------------------------------------------------------------

Public Sub NetReturnError(Reply&, entity$)
'THIS FUNCTION IS A GENERIC FUNCTION TO RETURN WIN 32 NET API ERROR CODES FOR MOST THINGS
    If Reply& = 0 Then Exit Sub
    TripError = True
    e$ = "While working with " & entity$ & ", "
    Select Case Reply&
        Case 5, 65
            e$ = e$ + "you had insufficient privileges and the action failed."
        Case 50
            e$ = e$ + "a network request was not supported."
        Case 53
            e$ = e$ + "a network path was not found and the action failed."
        Case 86
            e$ = e$ + "an insufficient old password was used and the action failed." + Chr$(13) + _
            "More than 2 more failed attempts to change password will cause an" + Chr$(13) + _
            "account lockout for 5 minutes."
        Case 87
            e$ = e$ + "an invalied parameter was specified and the action failed."
        Case 124
            e$ = e$ + "an invalid parameter was attempted to be passed and failed."
        Case 234, 2123
            e$ = e$ + "too little memory was allocated to perform the action and failed."
        Case 2102
            e$ = e$ + "a device driver was not installed to conduct the action."
        Case 2106
            e$ = e$ + "an action was attempted that can take place only from an NT Server."
        Case 2138
            e$ = e$ + "NT Workstation services have not been started and the action failed."
        Case 2141
            e$ = e$ + "either the NT Server is not configured for this transaction or IPC$ is not shared."
        Case 2202
            e$ = e$ + "an invalid NT user name was specified and the action failed."
        Case 2220
            e$ = e$ + "the NT group name does not exist."
        Case 2221
            e$ = e$ + "the NT user name was not found."
        Case 2223
            e$ = e$ + "the NT group name already exists."
        Case 2224
            e$ = e$ + "the NT user name already exists."
        Case 2226
            e$ = e$ + "the specified NT Server is not the Primary Domain Controller and the action failed."
        Case 2227
            e$ = e$ + "the NT Server is not running in user-level security."
        Case 2228
            e$ = e$ + "the accounts database became full."
        Case 2229
            e$ = e$ + "an unknown error (code 2229) occurred in accessing the NT accounts database."
        Case 2245
            e$ = e$ + "the password is too short."
        Case 2247
            e$ = e$ + "the NT accounts database file is corrupted."
        Case 2351
            e$ = e$ + "an invalid NT computer name was specified and the action failed."
        Case 2456
            e$ = e$ + "the NT user accounts database cannot be enlarged because the NT Server's hard disk is full."
        Case Else
            e$ = e$ + "an error code of " + Str(Reply&) + " was generated."
    End Select 'Reply&
    Beep
    Call MsgBox(e$, vbExclamation, "NT User Manager Transaction Error")
End Sub
