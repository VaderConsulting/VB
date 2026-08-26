Attribute VB_Name = "modMain"
Option Explicit

Public Type RASIPADDR
    a As Byte
    b As Byte
    c As Byte
    d As Byte
End Type

Public Enum RasEntryOptions
    RASEO_UseCountryAndAreaCodes = &H1
    RASEO_SpecificIpAddr = &H2
    RASEO_SpecificNameServers = &H4
    RASEO_IpHeaderCompression = &H8
    RASEO_RemoteDefaultGateway = &H10
    RASEO_DisableLcpExtensions = &H20
    RASEO_TerminalBeforeDial = &H40
    RASEO_TerminalAfterDial = &H80
    RASEO_ModemLights = &H100
    RASEO_SwCompression = &H200
    RASEO_RequireEncryptedPw = &H400
    RASEO_RequireMsEncryptedPw = &H800
    RASEO_RequireDataEncryption = &H1000
    RASEO_NetworkLogon = &H2000
    RASEO_UseLogonCredentials = &H4000
    RASEO_PromoteAlternates = &H8000
    RASEO_SecureLocalFiles = &H10000
    RASEO_RequireEAP = &H20000
    RASEO_RequirePAP = &H40000
    RASEO_RequireSPAP = &H80000
    RASEO_Custom = &H100000
    RASEO_PreviewPhoneNumber = &H200000
    RASEO_SharedPhoneNumbers = &H800000
    RASEO_PreviewUserPw = &H1000000
    RASEO_PreviewDomain = &H2000000
    RASEO_ShowDialingProgress = &H4000000
    RASEO_RequireCHAP = &H8000000
    RASEO_RequireMsCHAP = &H10000000
    RASEO_RequireMsCHAP2 = &H20000000
    RASEO_RequireW95MSCHAP = &H40000000
    RASEO_CustomScript = &H80000000
End Enum

Public Enum RASNetProtocols
    RASNP_NetBEUI = &H1
    RASNP_Ipx = &H2
    RASNP_Ip = &H4
End Enum

Public Enum RasFramingProtocols
    RASFP_Ppp = &H1
    RASFP_Slip = &H2
    RASFP_Ras = &H4
End Enum

Public Type VBRasEntry
    Options As RasEntryOptions
    CountryID As Long
    CountryCode As Long
    AreaCode As String
    LocalPhoneNumber As String
    AlternateNumbers As String
    ipAddr As RASIPADDR
    ipAddrDns As RASIPADDR
    ipAddrDnsAlt As RASIPADDR
    ipAddrWins As RASIPADDR
    ipAddrWinsAlt As RASIPADDR
    FrameSize As Long
    fNetProtocols As RASNetProtocols
    FramingProtocol As RasFramingProtocols
    ScriptName As String
    AutodialDll As String
    AutodialFunc As String
    DeviceType As String
    DeviceName As String
    X25PadType As String
    X25Address As String
    X25Facilities As String
    X25UserData As String
    Channels As Long
    NT4En_SubEntries As Long
    NT4En_DialMode As Long
    NT4En_DialExtraPercent As Long
    NT4En_DialExtraSampleSeconds As Long
    NT4En_HangUpExtraPercent As Long
    NT4En_HangUpExtraSampleSeconds As Long
    NT4En_IdleDisconnectSeconds As Long
    Win2000_Type As Long
    Win2000_EncryptionType As Long
    Win2000_CustomAuthKey As Long
    Win2000_guidId(0 To 15) As Byte
    Win2000_CustomDialDll As String
    Win2000_VpnStrategy As Long
End Type

Public Type VBRASDEVINFO
    DeviceType As String
    DeviceName As String
End Type

Type VBRASCTRYINFO
    CountryCode As Long
    CountryID As Long
    CountryName As String
    NextCountryID As Long
End Type


Public Declare Function RasEnumDevices Lib "rasapi32.dll" Alias "RasEnumDevicesA" (lpRasDevInfo As Any, lpCb As Long, lpCDevices As Long) As Long
Public Declare Function RasGetEntryProperties Lib "rasapi32.dll" Alias "RasGetEntryPropertiesA" (ByVal lpszPhonebook As String, ByVal lpszEntry As String, lpRasEntry As Any, lpdwEntryInfoSize As Long, lpbDeviceInfo As Any, lpdwDeviceInfoSize As Long) As Long
Public Declare Function RasSetEntryProperties Lib "rasapi32.dll" Alias "RasSetEntryPropertiesA" (ByVal lpszPhonebook As String, ByVal lpszEntry As String, lpRasEntry As Any, ByVal dwEntryInfoSize As Long, lpbDeviceInfo As Any, ByVal dwDeviceInfoSize As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function RasGetErrorString Lib "rasapi32.dll" Alias "RasGetErrorStringA" (ByVal uErrorValue As Long, ByVal lpszErrorString As String, cBufSize As Long) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Declare Function RasGetCountryInfo Lib "rasapi32.dll" Alias "RasGetCountryInfoA" (lpRasCtryInfo As Any, lpdwSize As Long) As Long

Function VBRASErrorHandler(rtn As Long) As String
    Dim strError As String, i As Long
    
    strError = String(512, 0)
    If rtn > 600 Then
        RasGetErrorString rtn, strError, 512&
    Else
        FormatMessage &H1000, ByVal 0&, rtn, 0&, strError, 512, ByVal 0&
    End If
    i = InStr(strError, Chr$(0))
    If i > 1 Then VBRASErrorHandler = Left$(strError, i - 1)
End Function

Function VBRasSetEntryProperties(strEntryName As String, clsRasEntry As VBRasEntry, Optional strPhoneBook As String) As Long
    Dim rtn As Long, lngCb As Long, lngBuffLen As Long
    Dim b() As Byte
    Dim lngPos As Long, lngStrLen As Long
    
    rtn = RasGetEntryProperties(vbNullString, vbNullString, ByVal 0&, lngCb, ByVal 0&, ByVal 0&)
    
    If rtn <> 603 Then
        VBRasSetEntryProperties = rtn
        Exit Function
    End If
    
    lngStrLen = Len(clsRasEntry.AlternateNumbers)
    lngBuffLen = lngCb + lngStrLen + 1
    ReDim b(lngBuffLen)
    
    CopyMemory b(0), lngCb, 4
    CopyMemory b(4), clsRasEntry.Options, 4
    CopyMemory b(8), clsRasEntry.CountryID, 4
    CopyMemory b(12), clsRasEntry.CountryCode, 4
    CopyStringToByte b(16), clsRasEntry.AreaCode, 11
    CopyStringToByte b(27), clsRasEntry.LocalPhoneNumber, 129
    
    If lngStrLen > 0 Then
        CopyMemory b(lngCb), ByVal clsRasEntry.AlternateNumbers, lngStrLen
        CopyMemory b(156), lngCb, 4
    End If
    
    CopyMemory b(160), clsRasEntry.ipAddr, 4
    CopyMemory b(164), clsRasEntry.ipAddrDns, 4
    CopyMemory b(168), clsRasEntry.ipAddrDnsAlt, 4
    CopyMemory b(172), clsRasEntry.ipAddrWins, 4
    CopyMemory b(176), clsRasEntry.ipAddrWinsAlt, 4
    CopyMemory b(180), clsRasEntry.FrameSize, 4
    CopyMemory b(184), clsRasEntry.fNetProtocols, 4
    CopyMemory b(188), clsRasEntry.FramingProtocol, 4
    CopyStringToByte b(192), clsRasEntry.ScriptName, 260
    CopyStringToByte b(452), clsRasEntry.AutodialDll, 260
    CopyStringToByte b(712), clsRasEntry.AutodialFunc, 260
    CopyStringToByte b(972), clsRasEntry.DeviceType, 17
    If lngCb = 1672& Then
        lngStrLen = 33
    Else
        lngStrLen = 129
    End If
    CopyStringToByte b(989), clsRasEntry.DeviceName, lngStrLen
    lngPos = 989 + lngStrLen
    CopyStringToByte b(lngPos), clsRasEntry.X25PadType, 33
    lngPos = lngPos + 33
    CopyStringToByte b(lngPos), clsRasEntry.X25Address, 201
    lngPos = lngPos + 201
    CopyStringToByte b(lngPos), clsRasEntry.X25Facilities, 201
    lngPos = lngPos + 201
    CopyStringToByte b(lngPos), clsRasEntry.X25UserData, 201
    lngPos = lngPos + 203
    CopyMemory b(lngPos), clsRasEntry.Channels, 4
   
    If lngCb > 1768 Then 'NT4 Enhancements & Win2000
        CopyMemory b(1768), clsRasEntry.NT4En_SubEntries, 4
        CopyMemory b(1772), clsRasEntry.NT4En_DialMode, 4
        CopyMemory b(1776), clsRasEntry.NT4En_DialExtraPercent, 4
        CopyMemory b(1780), clsRasEntry.NT4En_DialExtraSampleSeconds, 4
        CopyMemory b(1784), clsRasEntry.NT4En_HangUpExtraPercent, 4
        CopyMemory b(1788), clsRasEntry.NT4En_HangUpExtraSampleSeconds, 4
        CopyMemory b(1792), clsRasEntry.NT4En_IdleDisconnectSeconds, 4
        
        If lngCb > 1796 Then ' Win2000
            CopyMemory b(1796), clsRasEntry.Win2000_Type, 4
            CopyMemory b(1800), clsRasEntry.Win2000_EncryptionType, 4
            CopyMemory b(1804), clsRasEntry.Win2000_CustomAuthKey, 4
            CopyMemory b(1808), clsRasEntry.Win2000_guidId(0), 16
            CopyStringToByte b(1824), clsRasEntry.Win2000_CustomDialDll, 260
            CopyMemory b(2084), clsRasEntry.Win2000_VpnStrategy, 4
        End If
    End If
    rtn = RasSetEntryProperties(strPhoneBook, strEntryName, b(0), lngCb, ByVal 0&, ByVal 0&)
    
    VBRasSetEntryProperties = rtn
End Function

Function VBRasGetEntryProperties(strEntryName As String, clsRasEntry As VBRasEntry, Optional strPhoneBook As String) As Long
    Dim rtn As Long, lngCb As Long, lngBuffLen As Long
    Dim b() As Byte
    Dim lngPos As Long, lngStrLen As Long
    
    rtn = RasGetEntryProperties(vbNullString, vbNullString, ByVal 0&, lngCb, ByVal 0&, ByVal 0&)
    
    rtn = RasGetEntryProperties(strPhoneBook, strEntryName, ByVal 0&, lngBuffLen, ByVal 0&, ByVal 0&)
    
    If rtn <> 603 Then VBRasGetEntryProperties = rtn: Exit Function
    
    ReDim b(lngBuffLen - 1)
    CopyMemory b(0), lngCb, 4
    
    rtn = RasGetEntryProperties(strPhoneBook, strEntryName, b(0), lngBuffLen, ByVal 0&, ByVal 0&)
    
    VBRasGetEntryProperties = rtn
    If rtn <> 0 Then Exit Function
    
    CopyMemory clsRasEntry.Options, b(4), 4
    CopyMemory clsRasEntry.CountryID, b(8), 4
    CopyMemory clsRasEntry.CountryCode, b(12), 4
    CopyByteToTrimmedString clsRasEntry.AreaCode, b(16), 11
    CopyByteToTrimmedString clsRasEntry.LocalPhoneNumber, b(27), 129
    
    CopyMemory lngPos, b(156), 4
    If lngPos <> 0 Then
        lngStrLen = lngBuffLen - lngPos
        clsRasEntry.AlternateNumbers = String(lngStrLen, 0)
        CopyMemory ByVal clsRasEntry.AlternateNumbers, b(lngPos), lngStrLen
    End If
    
    CopyMemory clsRasEntry.ipAddr, b(160), 4
    CopyMemory clsRasEntry.ipAddrDns, b(164), 4
    CopyMemory clsRasEntry.ipAddrDnsAlt, b(168), 4
    CopyMemory clsRasEntry.ipAddrWins, b(172), 4
    CopyMemory clsRasEntry.ipAddrWinsAlt, b(176), 4
    CopyMemory clsRasEntry.FrameSize, b(180), 4
    CopyMemory clsRasEntry.fNetProtocols, b(184), 4
    CopyMemory clsRasEntry.FramingProtocol, b(188), 4
    CopyByteToTrimmedString clsRasEntry.ScriptName, b(192), 260
    CopyByteToTrimmedString clsRasEntry.AutodialDll, b(452), 260
    CopyByteToTrimmedString clsRasEntry.AutodialFunc, b(712), 260
    CopyByteToTrimmedString clsRasEntry.DeviceType, b(972), 17
    If lngCb = 1672& Then
        lngStrLen = 33
    Else
        lngStrLen = 129
    End If
    CopyByteToTrimmedString clsRasEntry.DeviceName, b(989), lngStrLen
    lngPos = 989 + lngStrLen
    CopyByteToTrimmedString clsRasEntry.X25PadType, b(lngPos), 33
    lngPos = lngPos + 33
    CopyByteToTrimmedString clsRasEntry.X25Address, b(lngPos), 201
    lngPos = lngPos + 201
    CopyByteToTrimmedString clsRasEntry.X25Facilities, b(lngPos), 201
    lngPos = lngPos + 201
    CopyByteToTrimmedString clsRasEntry.X25UserData, b(lngPos), 201
    lngPos = lngPos + 203
    CopyMemory clsRasEntry.Channels, b(lngPos), 4
    
    If lngCb > 1768 Then 'NT4 Enhancements & Win2000
        CopyMemory clsRasEntry.NT4En_SubEntries, b(1768), 4
        CopyMemory clsRasEntry.NT4En_DialMode, b(1772), 4
        CopyMemory clsRasEntry.NT4En_DialExtraPercent, b(1776), 4
        CopyMemory clsRasEntry.NT4En_DialExtraSampleSeconds, b(1780), 4
        CopyMemory clsRasEntry.NT4En_HangUpExtraPercent, b(1784), 4
        CopyMemory clsRasEntry.NT4En_HangUpExtraSampleSeconds, b(1788), 4
        CopyMemory clsRasEntry.NT4En_IdleDisconnectSeconds, b(1792), 4
        
        If lngCb > 1796 Then ' Win2000
            CopyMemory clsRasEntry.Win2000_Type, b(1796), 4
            CopyMemory clsRasEntry.Win2000_EncryptionType, b(1800), 4
            CopyMemory clsRasEntry.Win2000_CustomAuthKey, b(1804), 4
            CopyMemory clsRasEntry.Win2000_guidId(0), b(1808), 16
            CopyByteToTrimmedString clsRasEntry.Win2000_CustomDialDll, b(1824), 260
            CopyMemory clsRasEntry.Win2000_VpnStrategy, b(2084), 4
        End If
    End If
        
End Function

Function VBRasEnumDevices(clsVBRasDevInfo() As VBRASDEVINFO) As Long
    Dim rtn As Long, i As Long
    Dim lpCb As Long, lpCDevices As Long
    Dim b() As Byte
    Dim dwSize As Long
    
    rtn = RasEnumDevices(ByVal 0&, lpCb, lpCDevices)
    
    If lpCDevices = 0 Then Exit Function
    
    dwSize = lpCb \ lpCDevices
    
    ReDim b(lpCb - 1)
    
    CopyMemory b(0), dwSize, 4
    
    rtn = RasEnumDevices(b(0), lpCb, lpCDevices)
    
    If lpCDevices = 0 Then Exit Function
    
    ReDim clsVBRasDevInfo(lpCDevices - 1)
    
    For i = 0 To lpCDevices - 1
        CopyByteToTrimmedString clsVBRasDevInfo(i).DeviceType, b((i * dwSize) + 4), 17
        CopyByteToTrimmedString clsVBRasDevInfo(i).DeviceName, b((i * dwSize) + 21), dwSize - 21
    Next i
    
    VBRasEnumDevices = lpCDevices
    
End Function

Function VBRasGetCountryInfo(clsCountryInfo As VBRASCTRYINFO) As Long
    Dim b(511) As Byte, lpSize As Long, rtn As Long
    Dim lPos As Long, strTemp As String, lngLen As Long
    
    b(0) = 20
    CopyMemory b(4), clsCountryInfo.CountryID, 4
    lpSize = 512
    
    rtn = RasGetCountryInfo(b(0), lpSize)
    
    VBRasGetCountryInfo = rtn
    If rtn <> 0 Then Exit Function
    
    CopyMemory clsCountryInfo.NextCountryID, b(8), 4
    CopyMemory clsCountryInfo.CountryCode, b(12), 4
    
    CopyMemory lPos, b(16), 4
    lngLen = lpSize - lPos - 2
    
    If lngLen > 0 Then
        strTemp = String(lngLen, 0)
        CopyMemory ByVal strTemp, b(lPos), lngLen
    End If
    
    clsCountryInfo.CountryName = strTemp
End Function


Sub CopyByteToTrimmedString(strToCopyTo As String, bPos As Byte, lngMaxLen As Long)
    Dim strTemp As String, lngLen As Long
    
    strTemp = String(lngMaxLen + 1, 0)
    CopyMemory ByVal strTemp, bPos, lngMaxLen
    lngLen = InStr(strTemp, Chr$(0)) - 1
    strToCopyTo = Left$(strTemp, lngLen)
End Sub


Sub CopyStringToByte(bPos As Byte, strToCopy As String, lngMaxLen As Long)
    Dim lngLen As Long
    
    lngLen = Len(strToCopy)
    If lngLen = 0 Then
        Exit Sub
    ElseIf lngLen > lngMaxLen Then
        lngLen = lngMaxLen
    End If
    CopyMemory bPos, ByVal strToCopy, lngLen
End Sub

Public Function EnumCountries()
    Dim lngID As Long, rtn As Long
    Dim MyCountry As VBRASCTRYINFO
    
    lngID = 1
    
    Do
       MyCountry.CountryID = lngID
       rtn = VBRasGetCountryInfo(MyCountry)
    
       With MyCountry
            Debug.Print .CountryCode, .CountryName
        End With
    
        If rtn <> 0 Then Exit Do
        lngID = MyCountry.NextCountryID
    Loop While lngID <> 0
End Function
