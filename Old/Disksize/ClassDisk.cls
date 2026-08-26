VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
    (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
 
Private m_DriveName As String
Private m_FreeSpaceMB As Double
Private m_CapacityMB As Double
Private m_Serial As String
Private m_VolumeLabel As String
Private m_FileSystem As String
Private m_UsedspaceMB As Double
Private m_VolumeCompressed As Boolean
Private m_FilesCompressed As Boolean
Private m_CaseSensitiveSearch As Boolean
Private m_CasePreserved As Boolean
Private m_UniCodeOnDisk As Boolean


Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long


Const WN_Success = &H0
Const WN_Not_Supported = &H1
Const WN_Net_Error = &H2
Const WN_Bad_Pointer = &H4
Const WN_Bad_NetName = &H32
Const WN_Bad_Password = &H6
Const WN_Bad_Localname = &H33
Const WN_Access_Denied = &H7
Const WN_Out_Of_Memory = &HB
Const WN_Already_Connected = &H34
 '-- Error number and message
Public ErrorNum     As Long
Public ErrorMsg     As String
Public rc As Long
Private Const ERROR_NO_CONNECTION = 8
Private Const ERROR_NO_DISCONNECT = 9

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Sub Connect(sDrive, sService As String)
On Error GoTo Err_Connect
    Me.ErrorNum = 0
    Me.ErrorMsg = ""
 '   rc = WNetAddConnection(sService & Chr(0), "" & Chr(0), sDrive & Chr(0))
 rc = WNetAddConnection(sService & Chr(0), vbNullString, sDrive & Chr(0))
    If rc <> 0 Then GoTo Err_Connect
    Exit Sub
Err_Connect:
    Me.ErrorNum = rc
    Me.ErrorMsg = WnetError(rc)
End Sub

Public Sub DisConnect(sDrive As String)
On Error GoTo Err_DisConnect
    Me.ErrorNum = 0
    Me.ErrorMsg = ""
    rc = WNetCancelConnection(sDrive + Chr(0), 0)
 '   If rc <> 0 Then Error rc
    Exit Sub
Err_DisConnect:
    Me.ErrorNum = rc
    Me.ErrorMsg = WnetError(rc)
End Sub

Private Function WnetError(Errcode As Long) As String

Select Case Errcode
         Case WN_Not_Supported:
            WnetError = "Function is not supported."
         Case WN_Out_Of_Memory:
            WnetError = "Out of Memory."
         Case WN_Net_Error:
            WnetError = "An error occurred on the network."
         Case WN_Bad_Pointer:
            WnetError = "The Pointer was Invalid."
         Case WN_Bad_NetName:
            WnetError = "Invalid Network Resource Name."
         Case WN_Bad_Password:
            WnetError = "The Password was Invalid."
         Case WN_Bad_Localname:
            WnetError = "The local device name was invalid."
         Case WN_Access_Denied:
            WnetError = "A security violation occurred."
         Case WN_Already_Connected:
            WnetError = "The local device was connected to a remote resource."
         Case Else:
            WnetError = "Unrecognized Error " + Str(Errcode) + "."
      End Select
End Function

Public Property Get UsedSpaceMB() As Double
  UsedSpaceMB = m_UsedspaceMB
End Property

Public Property Get Filesystem() As String
  Filesystem = m_FileSystem
End Property

Public Property Get VolumeLabel() As String
  VolumeLabel = m_VolumeLabel
End Property

Public Property Get Serial() As String
  Serial = m_Serial
End Property

Public Property Get CapacityMB() As Double
  CapacityMB = m_CapacityMB
End Property

Public Property Get FreeSpaceMB() As Double
  FreeSpaceMB = m_FreeSpaceMB
End Property

Public Property Let DriveName(ByVal New_DriveName As String)
  m_DriveName = New_DriveName
  Call GetInformation(m_DriveName)
End Property

Public Function GetNextFreedrive()

      Dim DriveNum As Integer
      Dim FirstFreeDrive As String
      Dim FirstDrive As Integer
      Dim NextDrive As String
      DriveNum = -1
      Do
         DriveNum = DriveNum + 1
           NextDrive = Chr(DriveNum + 67) + ":\"
           FirstDrive = GetDriveType(NextDrive)
         Loop Until FirstDrive = 1
      FirstFreeDrive = Chr(DriveNum + 67) + ":"
      GetNextFreedrive = FirstFreeDrive
End Function

Private Sub GetInformation(sConnectDrive As String)

'-- Declarations
    Dim rc As Long                                              ' Return code
    Dim VolumeLabel As String
    Dim Filesystem As String
    Dim nSerial As Long
    Dim nMaxFileName As Long
    Dim nFileSystemFlags As Long
    
    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberOfFreeClusters As Long
    Dim TotalNumberOfClusters As Long
    Dim DiskCapacity As Double
    Dim FreeSpace As Double

'-- Initialize
    sConnectDrive = sConnectDrive & "\"
    VolumeLabel = Space(255)
    Filesystem = Space(255)
'--
    rc = GetVolumeInformation(sConnectDrive, VolumeLabel, Len(VolumeLabel), nSerial, nMaxFileName, nFileSystemFlags, Filesystem, Len(Filesystem))
    VolumeLabel = VolumeLabel & Chr(0)
    m_VolumeLabel = Left(VolumeLabel, InStr(VolumeLabel, Chr(0)) - 1)
    If m_VolumeLabel = "" Then m_VolumeLabel = "No label"
   
    Filesystem = Filesystem & Chr(0)
    m_FileSystem = Left(Filesystem, InStr(Filesystem, Chr(0)) - 1)
       
    m_Serial = nSerial
      
    rc = GetDiskFreeSpace(sConnectDrive, SectorsPerCluster, BytesPerSector, NumberOfFreeClusters, TotalNumberOfClusters)
    DiskCapacity = CDbl(BytesPerSector) * CDbl(SectorsPerCluster) * CDbl(TotalNumberOfClusters)  ' Convert to double to prevent overflow
    m_CapacityMB = Int(DiskCapacity / (1024# * 1024#))
    FreeSpace = CDbl(BytesPerSector) * CDbl(SectorsPerCluster) * CDbl(NumberOfFreeClusters) ' Convert to double to prevent overflow
    m_FreeSpaceMB = Int(FreeSpace / (1024# * 1024#))
    m_UsedspaceMB = m_CapacityMB - m_FreeSpaceMB
           
End Sub
