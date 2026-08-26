VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "drNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    WLX_NOTIFICATION_INFO
'    This structure stores information about a Winlogon event.
'
'    typedef struct _WLX_NOTIFICATION_INFO {
'                                           ULONG Size;
'                                           ULONG Flags;
'                                           PWSTR UserName;
'                                           PWSTR Domain;
'                                           PWSTR WindowStation;
'                                           HANDLE hToken;
'                                           HDESK hDesktop;
'                                           PFNMSGECALLBACK pStatusCallback;
'                                           } WLX_NOTIFICATION_INFO, *PWLX_NOTIFICATION_INFO;
'
'    Members

'    Size
'    Indicates the size of the structure, in bytes. Your application can check this value against
'    the structure size to validate the structure.

'    Flags
'    Reserved for internal use.

'    UserName
'    String that specifies the name of the user currently logged on to the system. If the event
'    occurs before a user logs on, this value is NULL.

'    Domain
'    String that specifies the name of the domain the user is currently logged on to. If the event
'    occurs before a user logs on, this value is NULL.

'    WindowStation
'    Specifies the name of the window station the user is currently logged on to. If the event occurs
'    before a user logs on, this value is NULL. Note that most configurations use a single, default
'    window station. Some applications, such as Terminal Services, use multiple window stations.

'    hToken
'    Handle to the user's token. This value is NULL if the event occurs before a user logs on.

'    hDesktop
'    Handle to the desktop that is currently active.

'    pStatusCallback
'    Reserved for internal use.

Public Type WLX_NOTIFICATION_INFO
    Size As Long
    Flags As Long
    UserName As String ' UNICODE
    Domain As String ' UNICODE
    WindowStation As String ' UNICODE
    hToken As Long
    hDesktop As Long
    pStatusCallback As Variant ' PFNMSGE
End Type


