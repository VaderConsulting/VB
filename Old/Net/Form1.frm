VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Shutdown"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enum Servers"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub Command1_Click()
    Dim pszTemp As String, pszServer As String, pszDomain As String
    Dim nLevel As Long, i As Long, BufPtr As Long, TempBufPtr As Long
    Dim nPrefMaxLen As Long, nEntriesRead As Long, nTotalEntries As Long
    Dim nServerType As Long, nResumeHandle As Long, nRes As Long
    Dim ServerInfo As SERVER_INFO_101
    
    ' Get the server name. It can be a null string
    pszTemp = Chr(0)
    pszTemp = InputBox("Enter server name:", "Server Name", "server")
    
    If Left(pszTemp, 2) <> "\\" Then pszTemp = "\\" & pszTemp
    
    If Len(pszTemp) = 0 Then
        pszServer = vbNullString
    Else
        pszServer = StrConv(pszTemp, vbUnicode)
    End If
    
    ' Get the domain name. It can be a null string
    pszTemp = Chr(0)
    pszTemp = InputBox("Enter domain name:", "Domain Name", "warnbro")
    If Len(pszTemp) = 0 Then
        pszDomain = vbNullString
    Else
        pszDomain = StrConv(pszTemp, vbUnicode)
    End If
    
    List1.Clear
    
    nLevel = 101
    BufPtr = 0
    nPrefMaxLen = &HFFFFFFFF
    nEntriesRead = 0
    nTotalEntries = 0
    nServerType = SV_TYPE_WORKSTATION
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
                List1.AddItem PointerToString(ServerInfo.lpszServerName)
                TempBufPtr = TempBufPtr + SIZE_SI_101
            Next i
        Else
            MsgBox "NetServerEnum failed: " & nRes
        End If
        NetApiBufferFree (BufPtr)
    Loop While nEntriesRead < nTotalEntries
    
End Sub
    
Private Sub Command2_Click()
    Unload Me
End Sub
    
' Works only on Win NT
Public Function GetPrimaryDCName(ByVal DName As String) As String
    
    Dim DCName As String, DCNPtr As Long
    Dim DNArray() As Byte, DCNArray(100) As Byte
    Dim result As Long
    DNArray = DName & vbNullChar
    ' Lookup the Primary Domain Controller
    result = NetGetDCName(0&, DNArray(0), DCNPtr)
    
    If result <> 0 Then
      Err.Raise vbObjectError + 4000, "CNetworkInfo", result
      Exit Function
    End If
     
    lstrcpyW DCNArray(0), DCNPtr
    result = NetApiBufferFree(DCNPtr)
    DCName = DCNArray()
     
    GetPrimaryDCName = Left(DCName, InStr(DCName, Chr(0)) - 1)

End Function

Private Function EnumDomains() As Variant
    
    Dim p_lngRtn As Long
    Dim p_lngEnumHwnd As Long
    Dim p_lngCount As Long
    Dim p_lngLoop As Long
    Dim p_lngBufSize As Long
    Dim p_astrDomainNames() As String
    Dim p_atypNetAPI(0 To MAX_RESOURCES) As NETRESOURCE
    
    ' ------------------------------------------
    ' First time thru, we are just getting the root level
    ' ------------------------------------------
    p_lngEnumHwnd = 0&
    p_lngRtn = WNetOpenEnum(dwScope:=RESOURCE_GLOBALNET, _
    dwType:=RESOURCETYPE_ANY, _
    dwUsage:=RESOURCEUSAGE_ALL, _
    lpNetResource:=ByVal 0&, _
    lppEnumHwnd:=p_lngEnumHwnd)
    
    If p_lngRtn = NO_ERROR Then
        p_lngCount = RESOURCE_ENUM_ALL
        
        p_lngBufSize = UBound(p_atypNetAPI) * Len(p_atypNetAPI(0))
        p_lngRtn = WNetEnumResource(pEnumHwnd:=p_lngEnumHwnd, _
        lpcCount:=p_lngCount, _
        lpBuffer:=p_atypNetAPI(0), _
        lpBufferSize:=p_lngBufSize)
        
        If p_lngCount > 0 Then
            For p_lngLoop = 0 To p_lngCount - 1
                Debug.Print PointerToAsciiStr(p_atypNetAPI(p_lngLoop).pRemoteName)
            Next p_lngLoop
        End If
        
    End If
    
    If p_lngEnumHwnd <> 0 Then
        Call WNetCloseEnum(p_lngEnumHwnd)
    End If
    
    ' ------------------------------------------
    ' Now we are going for the second level,
    ' which should contain the domain names
    ' ------------------------------------------
    p_lngRtn = WNetOpenEnum(dwScope:=RESOURCE_GLOBALNET, _
    dwType:=RESOURCETYPE_ANY, _
    dwUsage:=RESOURCEUSAGE_ALL, _
    lpNetResource:=p_atypNetAPI(0), _
    lppEnumHwnd:=p_lngEnumHwnd)
    
    If p_lngRtn = NO_ERROR Then
        p_lngCount = RESOURCE_ENUM_ALL
        
        p_lngBufSize = UBound(p_atypNetAPI) * Len(p_atypNetAPI(0))
        p_lngRtn = WNetEnumResource(pEnumHwnd:=p_lngEnumHwnd, _
        lpcCount:=p_lngCount, _
        lpBuffer:=p_atypNetAPI(0), _
        lpBufferSize:=p_lngBufSize)
        
        If p_lngCount > 0 Then
            ReDim p_astrDomainNames(1 To p_lngCount) As String
            For p_lngLoop = 0 To p_lngCount - 1
                p_astrDomainNames(p_lngLoop + 1) = PointerToAsciiStr(p_atypNetAPI(p_lngLoop).pRemoteName)
            Next p_lngLoop
        End If
    End If
    
    If p_lngEnumHwnd <> 0 Then
        Call WNetCloseEnum(p_lngEnumHwnd)
    End If
    
    ' ------------------------------------------
    ' Set the return value
    ' ------------------------------------------
    
    EnumDomains = p_astrDomainNames
End Function

Private Function PointerToAsciiStr(ByVal xi_lngPtrToString As Long) As String
    
    On Error Resume Next ' Don't accept an error here
    Dim p_lngLen As Long
    Dim p_strStringValue As String
    Dim p_lngNullPos As Long
    Dim p_lngRtn As Long
    
    p_lngLen = StrLenA(xi_lngPtrToString)
    
    If xi_lngPtrToString > 0 And p_lngLen > 0 Then
        p_strStringValue = Space$(p_lngLen + 1)
        p_lngRtn = StrCopyA(p_strStringValue, xi_lngPtrToString)
        p_lngNullPos = InStr(p_strStringValue, Chr$(0))
        If p_lngNullPos > 0 Then
            'Lose the null terminator...
            PointerToAsciiStr = Left$(p_strStringValue, p_lngNullPos - 1)
        Else
            PointerToAsciiStr = p_strStringValue 'Just pass the string...
        End If
    Else
        PointerToAsciiStr = ""
    End If

End Function

Public Sub GetDomains(lst As Object)

    Dim p_avntDomains As Variant
    Dim p_lngLoop As Long
    Dim p_lngNumItems As Long
        
    List1.Clear
        
    p_avntDomains = EnumDomains()
     
    On Error Resume Next
    p_lngNumItems = UBound(p_avntDomains)
    On Error GoTo 0
     
    If p_lngNumItems > 0 Then
        For p_lngLoop = 1 To p_lngNumItems
        List1.AddItem p_avntDomains(p_lngLoop)
        Next p_lngLoop
    End If
 
End Sub

Private Sub Command3_Click()
    frmShutdown.Show
End Sub
