Attribute VB_Name = "modMain"
Public Sub GetInfo(ServerName As String, DomainName, EnumType As Long, List As ListBox)
    Dim pszTemp As String, pszServer As String, pszDomain As String
    Dim nLevel As Long, i As Long, BufPtr As Long, TempBufPtr As Long
    Dim nPrefMaxLen As Long, nEntriesRead As Long, nTotalEntries As Long
    Dim nServerType As Long, nResumeHandle As Long, nRes As Long
    Dim ServerInfo As SERVER_INFO_101
    
    ' Get the server name. It can be a null string
    pszTemp = Chr(0)
    pszTemp = ServerName
    
    If Left(pszTemp, 2) <> "\\" Then pszTemp = "\\" & pszTemp
    
    If Len(pszTemp) = 0 Then
        pszServer = vbNullString
    Else
        pszServer = StrConv(pszTemp, vbUnicode)
    End If
    
    ' Get the domain name. It can be a null string
    pszTemp = Chr(0)
    pszTemp = DomainName
    If Len(pszTemp) = 0 Then
        pszDomain = vbNullString
    Else
        pszDomain = StrConv(pszTemp, vbUnicode)
    End If
    
    List.Clear
    
    nLevel = 101
    BufPtr = 0
    nPrefMaxLen = &HFFFFFFFF
    nEntriesRead = 0
    nTotalEntries = 0
    nServerType = EnumType
    nResumeHandle = 0
    
    Do
        nRes = NetServerEnum(pszServer, nLevel, BufPtr, _
        nPrefMaxLen, nEntriesRead, nTotalEntries, _
        nServerType, pszDomain, nResumeHandle)
        If ((nRes = ERROR_SUCCESS) Or (nRes = ERROR_MORE_DATA)) And _
            (nEntriesRead > 0) Then
            TempBufPtr = BufPtr
            For i = 1 To nEntriesRead
                RtlMoveMemory ServerInfo, TempBufPtr, SIZE_SI_101
                List.AddItem PointerToString(ServerInfo.lpszServerName)
                TempBufPtr = TempBufPtr + SIZE_SI_101
            Next i
        Else
            'MsgBox "NetServerEnum failed: " & nRes
        End If
        NetApiBufferFree (BufPtr)
    Loop While nEntriesRead < nTotalEntries
    
End Sub
