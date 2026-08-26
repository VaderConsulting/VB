
Option Explicit

' --------------------------------------------------------------------------------------
' Info gathered from MSDN, Platform SDK (January 2000), various .h files and internet.
' Last updated: February 03, 2000
' --------------------------------------------------------------------------------------

' some constants to get started
Public Const MAXLONG = &H7FFFFFFF
Public Const PARMNUM_BASE_INFOLEVEL = 1000


Public Enum SidTypes
   SidTypeUser = 1&           ' Indicates a user SID.
   SidTypeGroup = 2&          ' Indicates a group SID.
   SidTypeDomain = 3&         ' Indicates a domain SID.
   SidTypeAlias = 4&          ' Indicates an alias SID.
   SidTypeWellKnownGroup = 5& ' Indicates an SID for a well-known group. (such as Everyone)
   SidTypeDeletedAccount = 6& ' Indicates an SID for a deleted account.
   SidTypeInvalid = 7&        ' Indicates an invalid SID.
   SidTypeUnknown = 8&        ' Indicates an unknown SID type.
   SidTypeComputer = 9&       ' Indicates a SID for a computer.
End Enum

' Predefined identifier authority constants.
' The first four values are used with universal well-known SIDs;
' the last value is used with Windows NT well-known SIDs.
Public Enum SECURITY_SIDs
   SECURITY_NULL_SID_AUTHORITY = 0
   SECURITY_WORLD_SID_AUTHORITY = 1
   SECURITY_LOCAL_SID_AUTHORITY = 2
   SECURITY_CREATOR_SID_AUTHORITY = 3
   SECURITY_NT_AUTHORITY = 5
End Enum

Public Enum TokenTypes
   TokenUser = 1
   TokenGroups = 2
   TokenPrivileges = 3
   TokenOwner = 4
   TokenPrimaryGroup = 5
   TokenDefaultDacl = 6
   TokenSource = 7
   TokenType = 8
   TokenImpersonationLevel = 9
   TokenStatistics = 10
End Enum

Public Enum TOKEN_TYPE
   TokenPrimary = 1
   TokenImpersonation
End Enum

Public Enum SECURITY_IMPERSONATION_LEVEL
   SecurityAnonymous = 0
   SecurityIdentification
   SecurityImpersonation
   SecurityDelegation
End Enum



' Predefined Resource Types
Public Enum ResourceTypes
   RT_CURSOR = 1&
   RT_BITMAP = 2&
   RT_ICON = 3&
   RT_MENU = 4&
   RT_DIALOG = 5&
   RT_STRING = 6&
   RT_FONTDIR = 7&
   RT_FONT = 8&
   RT_ACCELERATOR = 9&
   RT_RCDATA = 10&
End Enum

'  ControlKeyState flags
Public Enum ControlKeyStateFlags
   RIGHT_ALT_PRESSED = &H1     '  the right alt key is pressed.
   LEFT_ALT_PRESSED = &H2     '  the left alt key is pressed.
   RIGHT_CTRL_PRESSED = &H4     '  the right ctrl key is pressed.
   LEFT_CTRL_PRESSED = &H8     '  the left ctrl key is pressed.
   SHIFT_PRESSED = &H10    '  the shift key is pressed.
   NUMLOCK_ON = &H20    '  the numlock light is on.
   SCROLLLOCK_ON = &H40    '  the scrolllock light is on.
   CAPSLOCK_ON = &H80    '  the capslock light is on.
   ENHANCED_KEY = &H100   '  the key is enhanced.
End Enum

' ButtonState flags
Public Enum ButtonStateFlags
   FROM_LEFT_1ST_BUTTON_PRESSED = &H1
   RIGHTMOST_BUTTON_PRESSED = &H2
   FROM_LEFT_2ND_BUTTON_PRESSED = &H4
   FROM_LEFT_3RD_BUTTON_PRESSED = &H8
   FROM_LEFT_4TH_BUTTON_PRESSED = &H10
End Enum

' EventType flags:
Public Enum EventTypeFlags
   KEY_EVENT = &H1     '  Event contains key event record
   mouse_eventC = &H2     '  Event contains mouse event record
   WINDOW_BUFFER_SIZE_EVENT = &H4     '  Event contains window change event record
   MENU_EVENT = &H8     '  Event contains menu event record
   FOCUS_EVENT = &H10    '  event contains focus change
End Enum

'  Attributes flags:
Public Enum AttributeFlags
   FOREGROUND_BLUE = &H1     '  text color contains blue.
   FOREGROUND_GREEN = &H2     '  text color contains green.
   FOREGROUND_RED = &H4     '  text color contains red.
   FOREGROUND_INTENSITY = &H8     '  text color is intensified.
   BACKGROUND_BLUE = &H10    '  background color contains blue.
   BACKGROUND_GREEN = &H20    '  background color contains green.
   BACKGROUND_RED = &H40    '  background color contains red.
   BACKGROUND_INTENSITY = &H80    '  background color is intensified.
End Enum

Public Enum ConnectFlags
   CONNECT_UPDATE_PROFILE = &H1
   CONNECT_UPDATE_RECENT = &H2
   CONNECT_TEMPORARY = &H4
   CONNECT_INTERACTIVE = &H8
   CONNECT_PROMPT = &H10
   CONNECT_NEED_DRIVE = &H20
   CONNECT_REFCOUNT = &H40
   CONNECT_REDIRECT = &H80
   CONNECT_LOCALDRIVE = &H100
   CONNECT_CURRENT_MEDIA = &H200
   CONNECT_DEFERRED = &H400
   CONNECT_RESERVED = &HFF000000
End Enum

Public Enum DropEffects
   DROPEFFECT_NONE = (0)
   DROPEFFECT_COPY = (1)
   DROPEFFECT_MOVE = (2)
   DROPEFFECT_LINK = (4)
   DROPEFFECT_SCROLL = (&H80000000)
End Enum

Public Enum SFGAO_Flags
   SFGAO_HIDDEN = &H8000                  ' hidden object
   
   ' capability flags
   SFGAO_CANCOPY = DROPEFFECT_COPY        ' Objects can be copied
   SFGAO_CANMOVE = DROPEFFECT_MOVE        ' Objects can be moved
   SFGAO_CANLINK = DROPEFFECT_LINK        ' Objects can be linked
   SFGAO_CANRENAME = &H1                  ' Objects can be renamed
   SFGAO_CANDELETE = &H2                  ' Objects can be deleted
   SFGAO_HASPROPSHEET = &H4               ' Objects have property sheets
   SFGAO_DROPTARGET = &H10                ' Objects are drop target
   SFGAO_CAPABILITYMASK = &H177
   
   ' display attributes
   SFGAO_LINK = &H1000                    ' Shortcut (link)
   SFGAO_SHARE = &H2000                   ' shared
   SFGAO_READONLY = &H4000                ' read-only
   SFGAO_GHOSTED = &H8000                 ' ghosted icon
   SFGAO_DISPLAYATTRMASK = &HF000         ' Mask for the display attributes
      
   ' contents flags
   SFGAO_HASSUBFOLDER = &H8000000         ' Expandable in the map pane
   SFGAO_CONTENTSMASK = &H8000000         ' Mask for the contents attributes.
   
   ' miscellaneous attributes
   SFGAO_NONENUMERATED = &H10000          ' is a non-enumerated object
   SFGAO_NEWCONTENT = &H20000             ' should show bold in explorer tree
   SFGAO_VALIDATE = &H100000              ' validate cached information
   SFGAO_REMOVABLE = &H200000             ' is this removeable media?
   SFGAO_COMPRESSED = &H400000            ' Object is compressed (use alt color)
   SFGAO_BROWSABLE = &H800000             ' is in-place browsable
   SFGAO_FILESYSANCESTOR = &H1000000      ' It contains file system folder
   SFGAO_FOLDER = &H2000000               ' It's a folder.
   SFGAO_FILESYSTEM = &H4000000           ' is a file system thing (file/folder/root)
End Enum

' NCB Command codes
Public Enum NCBCommandCodes
   NCBCALL = &H10             '  NCB CALL
   NCBLISTEN = &H11           '  NCB LISTEN
   NCBHANGUP = &H12           '  NCB HANG UP
   NCBSEND = &H14             '  NCB SEND
   NCBRECV = &H15             '  NCB RECEIVE
   NCBRECVANY = &H16          '  NCB RECEIVE ANY
   NCBCHAINSEND = &H17        '  NCB CHAIN SEND
   NCBDGSEND = &H20           '  NCB SEND DATAGRAM
   NCBDGRECV = &H21           '  NCB RECEIVE DATAGRAM
   NCBDGSENDBC = &H22         '  NCB SEND BROADCAST DATAGRAM
   NCBDGRECVBC = &H23         '  NCB RECEIVE BROADCAST DATAGRAM
   NCBADDNAME = &H30          '  NCB ADD NAME
   NCBDELNAME = &H31          '  NCB DELETE NAME
   NCBRESET = &H32            '  NCB RESET
   NCBASTAT = &H33            '  NCB ADAPTER STATUS
   NCBSSTAT = &H34            '  NCB SESSION STATUS
   NCBCANCEL = &H35           '  NCB CANCEL
   NCBADDGRNAME = &H36        '  NCB ADD GROUP NAME
   NCBENUM = &H37             '  NCB ENUMERATE LANA NUMBERS
   NCBUNLINK = &H70           '  NCB UNLINK
   NCBSENDNA = &H71           '  NCB SEND NO ACK
   NCBCHAINSENDNA = &H72      '  NCB CHAIN SEND NO ACK
   NCBLANSTALERT = &H73       '  NCB LAN STATUS ALERT
   NCBACTION = &H77           '  NCB ACTION
   NCBFINDNAME = &H78         '  NCB FIND NAME
   NCBTRACE = &H79            '  NCB TRACE

   ASYNCH = &H80  '  high bit set == ASYNCHronous
End Enum

Public Enum PrinterChanges
   PRINTER_CHANGE_ADD_PRINTER = &H1                      ' A printer was added to the server.
   PRINTER_CHANGE_SET_PRINTER = &H2                      ' A printer was set.
   PRINTER_CHANGE_DELETE_PRINTER = &H4                   ' A printer was deleted.
   PRINTER_CHANGE_FAILED_CONNECTION_PRINTER = &H8        ' A printer connection has failed.
   PRINTER_CHANGE_PRINTER = &HFF                         '
   PRINTER_CHANGE_ADD_JOB = &H100                        ' A print job was sent to the printer.
   PRINTER_CHANGE_SET_JOB = &H200                        ' A job was set.
   PRINTER_CHANGE_DELETE_JOB = &H400                     ' A job was deleted.
   PRINTER_CHANGE_WRITE_JOB = &H800                      ' Job data was written.
   PRINTER_CHANGE_JOB = &HFF00                           '
   PRINTER_CHANGE_ADD_FORM = &H10000                     ' A form was added to the server.
   PRINTER_CHANGE_SET_FORM = &H20000                     ' A form was set on the server.
   PRINTER_CHANGE_DELETE_FORM = &H40000                  ' A form was deleted from the server.
   PRINTER_CHANGE_FORM = &H70000                         '
   PRINTER_CHANGE_ADD_PORT = &H100000                    ' A port or monitor was added to the server.
   PRINTER_CHANGE_CONFIGURE_PORT = &H200000              ' A port was configured on the server.
   PRINTER_CHANGE_DELETE_PORT = &H400000                 ' A port or monitor was deleted from the server.
   PRINTER_CHANGE_PORT = &H700000                        '
   PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000        ' A print processor was added to the server.
   PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000     ' A print processor was deleted from the server.
   PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000            '
   PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000        ' A printer driver was added to the server.
   PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000     ' A printer driver was deleted from the server.
   PRINTER_CHANGE_PRINTER_DRIVER = &H70000000            '
   PRINTER_CHANGE_TIMEOUT = &H80000000                   ' The job timed out.
   PRINTER_CHANGE_ALL = &H7777FFFF                       '
End Enum

Public Enum PrinterFlags
   PRINTER_ENUM_DEFAULT = &H1          ' Windows 95: The function returns information about the default printer.
   PRINTER_ENUM_LOCAL = &H2            ' Ignores the Name parameter, and enumerates the locally installed printers.  Windows 95: The function will also enumerate network printers because they are handled by the local print provider.
   PRINTER_ENUM_CONNECTIONS = &H4      ' Windows NT/2000: The function enumerates the list of printers to which the user has made previous connections.
   PRINTER_ENUM_FAVORITE = &H4         ' Enumerates a list of favorite printers. This is essentially the list of printers that the user has made previous connections to.
   PRINTER_ENUM_NAME = &H8             ' Enumerates the printer identified by Name. This can be a server, a domain, or a print provider. If Name is NULL, the function enumerates available print providers.
   PRINTER_ENUM_REMOTE = &H10          ' Windows NT/2000: The function enumerates network printers and print servers in the computer's domain. This value is valid only if Level is 1.
   PRINTER_ENUM_SHARED = &H20          ' Enumerates printers that have the shared attribute. Cannot be used in isolation; use an OR operation to combine with another PRINTER_ENUM type.
   PRINTER_ENUM_NETWORK = &H40         ' Windows NT/2000: The function enumerates network printers in the computer's domain. This value is valid only if Level is 1.

   PRINTER_ENUM_EXPAND = &H4000        ' A print provider can set this flag as a hint to a calling application to enumerate this object further if default expansion is enabled. For example, when domains are enumerated, a print provider might indicate the user's domain by setting this flag
   PRINTER_ENUM_CONTAINER = &H8000     ' If this flag is set, the printer object may contain enumerable objects. For example, the object may be a print server that contains printers.

   PRINTER_ENUM_ICONMASK = &HFF0000    '
   PRINTER_ENUM_ICON1 = &H10000        ' Indicates that, where appropriate, an application should display an icon identifying the object as a top-level network name, such as Microsoft Windows Network.
   PRINTER_ENUM_ICON2 = &H20000        ' Indicates that, where appropriate, an application should display an icon that identifies the object as a network domain.
   PRINTER_ENUM_ICON3 = &H40000        ' Indicates that, where appropriate, an application should display an icon that identifies the object as a print server.
   PRINTER_ENUM_ICON4 = &H80000        ' Reserved for future use.
   PRINTER_ENUM_ICON5 = &H100000       ' Reserved for future use.
   PRINTER_ENUM_ICON6 = &H200000       ' Reserved for future use.
   PRINTER_ENUM_ICON7 = &H400000       ' Reserved for future use.
   PRINTER_ENUM_ICON8 = &H800000       ' Indicates that, where appropriate, an application should display an icon that identifies the object as a printer.
   PRINTER_ENUM_HIDE = &H1000000
End Enum

Public Enum PrinterStatus
   PRINTER_STATUS_PAUSED = &H1                     ' The printer is paused.
   PRINTER_STATUS_ERROR = &H2                      ' The printer is in an error state.
   PRINTER_STATUS_PENDING_DELETION = &H4           ' The printer is deleting a print job.
   PRINTER_STATUS_PAPER_JAM = &H8                  ' Paper is jammed in the printer
   PRINTER_STATUS_PAPER_OUT = &H10                 ' The printer is out of paper.
   PRINTER_STATUS_MANUAL_FEED = &H20               ' The printer is in a manual feed state.
   PRINTER_STATUS_PAPER_PROBLEM = &H40             ' The printer has a paper problem.
   PRINTER_STATUS_OFFLINE = &H80                   ' The printer is offline.
   PRINTER_STATUS_IO_ACTIVE = &H100                ' The printer is in an active input/output state
   PRINTER_STATUS_BUSY = &H200                     ' The printer is busy.
   PRINTER_STATUS_PRINTING = &H400                 ' The printer is printing.
   PRINTER_STATUS_OUTPUT_BIN_FULL = &H800          ' The printer's output bin is full.
   PRINTER_STATUS_NOT_AVAILABLE = &H1000           ' The printer is not available for printing.
   PRINTER_STATUS_WAITING = &H2000                 ' The printer is waiting.
   PRINTER_STATUS_PROCESSING = &H4000              ' The printer is processing a print job.
   PRINTER_STATUS_INITIALIZING = &H8000            ' The printer is initializing.
   PRINTER_STATUS_WARMING_UP = &H10000             ' The printer is warming up.
   PRINTER_STATUS_TONER_LOW = &H20000              ' The printer is low on toner.
   PRINTER_STATUS_NO_TONER = &H40000               ' The printer is out of toner.
   PRINTER_STATUS_PAGE_PUNT = &H80000              ' The printer cannot print the current page. Windows 95: Indicates the page is being "punted" (that is, not printed) because it is too complex for the printer to print.
   PRINTER_STATUS_USER_INTERVENTION = &H100000     ' The user needs to do something to the printer.
   PRINTER_STATUS_OUT_OF_MEMORY = &H200000         ' The printer has run out of memory.
   PRINTER_STATUS_DOOR_OPEN = &H400000             ' The printer door is open.
   PRINTER_STATUS_SERVER_UNKNOWN = &H800000        ' The printer status is unknown.
   PRINTER_STATUS_POWER_SAVE = &H1000000           ' The printer is in power-save mode
End Enum

Public Enum PrinterAttributes
   PRINTER_ATTRIBUTE_QUEUED = &H1                  ' If set, the printer spools and starts printing after the last page is spooled. If not set and PRINTER_ATTRIBUTE_DIRECT is not set, the printer spools and prints while spooling.
   PRINTER_ATTRIBUTE_DIRECT = &H2                  ' Job is sent directly to the printer (it is not spooled).
   PRINTER_ATTRIBUTE_DEFAULT = &H4                 ' Windows 95: Indicates the printer is the default printer in the system.
   PRINTER_ATTRIBUTE_SHARED = &H8                  ' Printer is shared.
   PRINTER_ATTRIBUTE_NETWORK = &H10                ' Printer is a network printer connection.
   PRINTER_ATTRIBUTE_HIDDEN = &H20                 ' Reserved.
   PRINTER_ATTRIBUTE_LOCAL = &H40                  ' Printer is a local printer.

   PRINTER_ATTRIBUTE_ENABLE_DEVQ = &H80            ' If set, DevQueryPrint is called. DevQueryPrint may fail if the document and printer setups do not match. Setting this flag causes mismatched documents to be held in the queue.
   PRINTER_ATTRIBUTE_KEEPPRINTEDJOBS = &H100       ' If set, jobs are kept after they are printed. If unset, jobs are deleted.
   PRINTER_ATTRIBUTE_DO_COMPLETE_FIRST = &H200     ' If set and printer is set for print-while-spooling, any jobs that have completed spooling are scheduled to print before jobs that have not completed spooling.
   
   PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400          ' Windows 95: Indicates whether the printer is currently connected. If the printer is not currently connected, print jobs will continue to spool.
   PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800           ' Windows 95: Indicates whether bi-directional communications are enabled for the printer.
   PRINTER_ATTRIBUTE_RAW_ONLY = &H1000             ' Indicates that only raw data type print jobs can be spooled.
   PRINTER_ATTRIBUTE_PUBLISHED = &H2000            ' Windows 2000: Indicates whether the printer is published in the directory service.
End Enum

Public Enum PrinterControls
   PRINTER_CONTROL_PAUSE = 1&                      ' Pauses the printer.
   PRINTER_CONTROL_RESUME = 2&                     ' Deletes all print jobs in the printer.
   PRINTER_CONTROL_PURGE = 3&                      ' Resumes a paused printer.
   PRINTER_CONTROL_SET_STATUS = 4&                 ' Sets the printer status. The pPrinter parameter is a pointer to a DWORD that specifies the new printer status.
End Enum

'  conversation status bits (fsStatus)
Public Enum ConversationStatusBits
   ST_CONNECTED = &H1
   ST_ADVISE = &H2
   ST_ISLOCAL = &H4
   ST_BLOCKED = &H8
   ST_CLIENT = &H10
   ST_TERMINATED = &H20
   ST_INLIST = &H40
   ST_BLOCKNEXT = &H80
   ST_ISSELF = &H100
End Enum

'  conversation states (usState)
Public Enum ConversationStates
   XST_NULL = 0            '  quiescent states
   XST_INCOMPLETE = 1
   XST_CONNECTED = 2
   XST_INIT1 = 3           '  mid-initiation states
   XST_INIT2 = 4
   XST_REQSENT = 5         '  active conversation states
   XST_DATARCVD = 6
   XST_POKESENT = 7
   XST_POKEACKRCVD = 8
   XST_EXECSENT = 9
   XST_EXECACKRCVD = 10
   XST_ADVSENT = 11
   XST_UNADVSENT = 12
   XST_ADVACKRCVD = 13
   XST_UNADVACKRCVD = 14
   XST_ADVDATASENT = 15
   XST_ADVDATAACKRCVD = 16
End Enum

' notifications passed in low word of lParam on WM_COMMNOTIFY messages
Public Enum WM_COMMNOTIFYNotifications
   CN_RECEIVE = &H1
   CN_TRANSMIT = &H2
   CN_EVENT = &H4
End Enum

' WM_SYNCTASK Commands
Public Enum WM_SYNCTASKCommands
   ST_BEGINSWP = 0
   ST_ENDSWP = 1
End Enum

' Monitor Flags
' Callback filter flags for use with MONITOR apps - 0 implies no monitor callbacks
Public Enum MonitorFlags
   MF_HSZ_INFO = &H1000000
   MF_SENDMSGS = &H2000000
   MF_POSTMSGS = &H4000000
   MF_CALLBACKS = &H8000000
   MF_ERRORS = &H10000000
   MF_LINKS = &H20000000
   MF_CONV = &H40000000
   MF_MASK = &HFF000000
End Enum

' Callback Filter Flags for use with standard apps.
Public Enum CallbackFilterFlags
   CBF_FAIL_SELFCONNECTIONS = &H1000
   CBF_FAIL_CONNECTIONS = &H2000
   CBF_FAIL_ADVISES = &H4000
   CBF_FAIL_EXECUTES = &H8000
   CBF_FAIL_POKES = &H10000
   CBF_FAIL_REQUESTS = &H20000
   CBF_FAIL_ALLSVRXACTIONS = &H3F000

   CBF_SKIP_CONNECT_CONFIRMS = &H40000
   CBF_SKIP_REGISTRATIONS = &H80000
   CBF_SKIP_UNREGISTRATIONS = &H100000
   CBF_SKIP_DISCONNECTS = &H200000
   CBF_SKIP_ALLNOTIFICATIONS = &H3C0000
End Enum

Public Enum RegistryKeys
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA = &H80000004
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum

Public Enum ControlIdentifiers
   ctlFirst = &H400
   ctlLast = &H4FF
   ' Push buttons
   psh1 = &H400
   psh2 = &H401
   psh3 = &H402
   psh4 = &H403
   psh5 = &H404
   psh6 = &H405
   psh7 = &H406
   psh8 = &H407
   psh9 = &H408
   psh10 = &H409
   psh11 = &H40A
   psh12 = &H40B
   psh13 = &H40C
   psh14 = &H40D
   psh15 = &H40E
   pshHelp = psh15
   psh16 = &H40F
   ' Checkboxes
   chx1 = &H410
   chx2 = &H411
   chx3 = &H412
   chx4 = &H413
   chx5 = &H414
   chx6 = &H415
   chx7 = &H416
   chx8 = &H417
   chx9 = &H418
   chx10 = &H419
   chx11 = &H41A
   chx12 = &H41B
   chx13 = &H41C
   chx14 = &H41D
   chx15 = &H41E
   chx16 = &H41D
   ' Radio buttons
   rad1 = &H420
   rad2 = &H421
   rad3 = &H422
   rad4 = &H423
   rad5 = &H424
   rad6 = &H425
   rad7 = &H426
   rad8 = &H427
   rad9 = &H428
   rad10 = &H429
   rad11 = &H42A
   rad12 = &H42B
   rad13 = &H42C
   rad14 = &H42D
   rad15 = &H42E
   rad16 = &H42F
   '  Groups, frames, rectangles, and icons
   grp1 = &H430
   grp2 = &H431
   grp3 = &H432
   grp4 = &H433
   
   frm1 = &H434
   frm2 = &H435
   frm3 = &H436
   frm4 = &H437
   
   rct1 = &H438
   rct2 = &H439
   rct3 = &H43A
   rct4 = &H43B
   
   ico1 = &H43C
   ico2 = &H43D
   ico3 = &H43E
   ico4 = &H43F
   ' Static text
   stc1 = &H440
   stc2 = &H441
   stc3 = &H442
   stc4 = &H443
   stc5 = &H444
   stc6 = &H445
   stc7 = &H446
   stc8 = &H447
   stc9 = &H448
   stc10 = &H449
   stc11 = &H44A
   stc12 = &H44B
   stc13 = &H44C
   stc14 = &H44D
   stc15 = &H44E
   stc16 = &H44F
   stc17 = &H450
   stc18 = &H451
   stc19 = &H452
   stc20 = &H453
   stc21 = &H454
   stc22 = &H455
   stc23 = &H456
   stc24 = &H457
   stc25 = &H458
   stc26 = &H459
   stc27 = &H45A
   stc28 = &H45B
   stc29 = &H45C
   stc30 = &H45D
   stc31 = &H45E
   stc32 = &H45F
   ' Listboxes
   lst1 = &H460
   lst2 = &H461
   lst3 = &H462
   lst4 = &H463
   lst5 = &H464
   lst6 = &H465
   lst7 = &H466
   lst8 = &H467
   lst9 = &H468
   lst10 = &H469
   lst11 = &H46A
   lst12 = &H46B
   lst13 = &H46C
   lst14 = &H46D
   lst15 = &H46E
   lst16 = &H46F
   ' Combo boxes
   cmb1 = &H470
   cmb2 = &H471
   cmb3 = &H472
   cmb4 = &H473
   cmb5 = &H474
   cmb6 = &H475
   cmb7 = &H476
   cmb8 = &H477
   cmb9 = &H478
   cmb10 = &H479
   cmb11 = &H47A
   cmb12 = &H47B
   cmb13 = &H47C
   cmb14 = &H47D
   cmb15 = &H47E
   cmb16 = &H47F
   ' Edit controls
   edt1 = &H480
   edt2 = &H481
   edt3 = &H482
   edt4 = &H483
   edt5 = &H484
   edt6 = &H485
   edt7 = &H486
   edt8 = &H487
   edt9 = &H488
   edt10 = &H489
   edt11 = &H48A
   edt12 = &H48B
   edt13 = &H48C
   edt14 = &H48D
   edt15 = &H48E
   edt16 = &H48F
   ' Scroll bars
   scr1 = &H490
   scr2 = &H491
   scr3 = &H492
   scr4 = &H493
   scr5 = &H494
   scr6 = &H495
   scr7 = &H496
   scr8 = &H497
End Enum


' WININET ENUMS

' HTTP Response Status Codes:
Public Enum HTTPResponses
   HTTP_STATUS_CONTINUE = 100           ' OK to continue with request
   HTTP_STATUS_SWITCH_PROTOCOLS = 101   ' server has switched protocols in upgrade header
   
   HTTP_STATUS_OK = 200                 ' request completed
   HTTP_STATUS_CREATED = 201            ' object created, reason = new URI
   HTTP_STATUS_ACCEPTED = 202           ' async completion (TBS)
   HTTP_STATUS_PARTIAL = 203            ' partial completion
   HTTP_STATUS_NO_CONTENT = 204         ' no info to return
   HTTP_STATUS_RESET_CONTENT = 205      ' request completed, but clear form
   HTTP_STATUS_PARTIAL_CONTENT = 206    ' partial GET furfilled
   
   HTTP_STATUS_AMBIGUOUS = 300          ' server couldn't decide what to return
   HTTP_STATUS_MOVED = 301              ' object permanently moved
   HTTP_STATUS_REDIRECT = 302           ' object temporarily moved
   HTTP_STATUS_REDIRECT_METHOD = 303    ' redirection w/ new access method
   HTTP_STATUS_NOT_MODIFIED = 304       ' if-modified-since was not modified
   HTTP_STATUS_USE_PROXY = 305          ' redirection to proxy, location header specifies proxy to use
   HTTP_STATUS_REDIRECT_KEEP_VERB = 307 ' HTTP/1.1: keep same verb
   
   HTTP_STATUS_BAD_REQUEST = 400        ' invalid syntax
   HTTP_STATUS_DENIED = 401             ' access denied
   HTTP_STATUS_PAYMENT_REQ = 402        ' payment required
   HTTP_STATUS_FORBIDDEN = 403          ' request forbidden
   HTTP_STATUS_NOT_FOUND = 404          ' object not found
   HTTP_STATUS_BAD_METHOD = 405         ' method is not allowed
   HTTP_STATUS_NONE_ACCEPTABLE = 406    ' no response acceptable to client found
   HTTP_STATUS_PROXY_AUTH_REQ = 407     ' proxy authentication required
   HTTP_STATUS_REQUEST_TIMEOUT = 408    ' server timed out waiting for request
   HTTP_STATUS_CONFLICT = 409           ' user should resubmit with more info
   HTTP_STATUS_GONE = 410               ' the resource is no longer available
   HTTP_STATUS_LENGTH_REQUIRED = 411    ' the server refused to accept request w/o a length
   HTTP_STATUS_PRECOND_FAILED = 412     ' precondition given in request failed
   HTTP_STATUS_REQUEST_TOO_LARGE = 413  ' request entity was too large
   HTTP_STATUS_URI_TOO_LONG = 414       ' request URI too long
   HTTP_STATUS_UNSUPPORTED_MEDIA = 415  ' unsupported media type
   HTTP_STATUS_RETRY_WITH = 449         ' retry after doing the appropriate action.
   
   HTTP_STATUS_SERVER_ERROR = 500       ' internal server error
   HTTP_STATUS_NOT_SUPPORTED = 501      ' required not supported
   HTTP_STATUS_BAD_GATEWAY = 502        ' error response received from gateway
   HTTP_STATUS_SERVICE_UNAVAIL = 503    ' temporarily overloaded
   HTTP_STATUS_GATEWAY_TIMEOUT = 504    ' timed out waiting for gateway
   HTTP_STATUS_VERSION_NOT_SUP = 505    ' HTTP version not supported
   
   HTTP_STATUS_FIRST = HTTP_STATUS_CONTINUE
   HTTP_STATUS_LAST = HTTP_STATUS_VERSION_NOT_SUP
End Enum

Public Enum InternetScheme
   INTERNET_SCHEME_PARTIAL = -2
   INTERNET_SCHEME_UNKNOWN = -1
   INTERNET_SCHEME_DEFAULT = 0
   INTERNET_SCHEME_FTP = 1
   INTERNET_SCHEME_GOPHER = 2
   INTERNET_SCHEME_HTTP = 3
   INTERNET_SCHEME_HTTPS = 4
   INTERNET_SCHEME_FILE = 5
   INTERNET_SCHEME_NEWS = 6
   INTERNET_SCHEME_MAILTO = 7
   INTERNET_SCHEME_FIRST = INTERNET_SCHEME_FTP
   INTERNET_SCHEME_LAST = INTERNET_SCHEME_MAILTO
End Enum

Public Enum CacheEntryTypeFlags
   NORMAL_CACHE_ENTRY = &H1
   STICKY_CACHE_ENTRY = &H4
   EDITED_CACHE_ENTRY = &H8
   TRACK_OFFLINE_CACHE_ENTRY = &H10
   TRACK_ONLINE_CACHE_ENTRY = &H20
   SPARSE_CACHE_ENTRY = &H10000
   COOKIE_CACHE_ENTRY = &H100000
   URLHISTORY_CACHE_ENTRY = &H200000
End Enum

'flags for InternetCanonicalizeUrl() and InternetCombineUrl()
Public Enum URLFlags
   ICU_NO_ENCODE = &H20000000             ' Don't convert unsafe characters to escape sequence
   ICU_DECODE = &H10000000                ' Convert %XX escape sequences to characters
   ICU_NO_META = &H8000000                ' Don't convert .. etc. meta path sequences
   ICU_ENCODE_SPACES_ONLY = &H4000000     ' Encode spaces only
   ICU_BROWSER_MODE = &H2000000           ' Special encode/decode rules for browser
   ICU_ENCODE_PERCENT = &H1000            ' Encode any percent (ASCII25) signs encountered, default is to not encode percent. IE5 Bug#21680

   ' flags for InternetCrackUrl() and InternetCreateUrl()
   ICU_ESCAPE = &H80000000      ' (un)escape URL characters
   ICU_USERNAME = &H40000000    ' use internal username & password
End Enum

' manifests for GopherType
Public Enum GopherTypes
   GOPHER_TYPE_TEXT_FILE = &H1
   GOPHER_TYPE_DIRECTORY = &H2
   GOPHER_TYPE_CSO = &H4
   GOPHER_TYPE_ERROR = &H8
   GOPHER_TYPE_MAC_BINHEX = &H10
   GOPHER_TYPE_DOS_ARCHIVE = &H20
   GOPHER_TYPE_UNIX_UUENCODED = &H40
   GOPHER_TYPE_INDEX_SERVER = &H80
   GOPHER_TYPE_TELNET = &H100
   GOPHER_TYPE_BINARY = &H200
   GOPHER_TYPE_REDUNDANT = &H400
   GOPHER_TYPE_TN3270 = &H800
   GOPHER_TYPE_GIF = &H1000
   GOPHER_TYPE_IMAGE = &H2000
   GOPHER_TYPE_BITMAP = &H4000
   GOPHER_TYPE_MOVIE = &H8000
   GOPHER_TYPE_SOUND = &H10000
   GOPHER_TYPE_HTML = &H20000
   GOPHER_TYPE_PDF = &H40000
   GOPHER_TYPE_CALENDAR = &H80000
   GOPHER_TYPE_INLINE = &H100000
   GOPHER_TYPE_UNKNOWN = &H20000000
   GOPHER_TYPE_ASK = &H40000000
   GOPHER_TYPE_GOPHER_PLUS = &H80000000
End Enum

Public Enum HttpAddRequestHeaderFlags
' values for dwModifiers parameter of HttpAddRequestHeaders()
   HTTP_ADDREQ_INDEX_MASK = &HFFFF
   HTTP_ADDREQ_FLAGS_MASK = &HFFFF0000

' HTTP_ADDREQ_FLAG_ADD_IF_NEW - the header will only be added if it doesn't
' already exist
   HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' HTTP_ADDREQ_FLAG_ADD - if HTTP_ADDREQ_FLAG_REPLACE is set but the header is
' not found then if this flag is set, the header is added anyway, so long as
' there is a valid header-value
   HTTP_ADDREQ_FLAG_ADD = &H20000000

' HTTP_ADDREQ_FLAG_COALESCE - coalesce headers with same name. e.g.
' "Accept: text/*" and "Accept: audio/*" with this flag results in a single
' header: "Accept: text/*, audio/*"
   HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA = &H40000000
   HTTP_ADDREQ_FLAG_COALESCE_WITH_SEMICOLON = &H1000000
   HTTP_ADDREQ_FLAG_COALESCE = HTTP_ADDREQ_FLAG_COALESCE_WITH_COMMA

' HTTP_ADDREQ_FLAG_REPLACE - replaces the specified header. Only one header can
' be supplied in the buffer. If the header to be replaced is not the first
' in a list of headers with the same name, then the relative index should be
' supplied in the low 8 bits of the dwModifiers parameter. If the header-value
' part is missing, then the header is removed
   HTTP_ADDREQ_FLAG_REPLACE = &H80000000
End Enum


' common per-API flags (new APIs)
Public Enum WinInetFlags
   WININET_API_FLAG_ASYNC = &H1                  ' force async operation
   WININET_API_FLAG_SYNC = &H4                   ' force sync operation
   WININET_API_FLAG_USE_CONTEXT = &H8            ' use value supplied in dwContext (even if 0)
End Enum

' flags for HttpSendRequestEx(), HttpEndRequest()
Public Enum HttpEndRequestFlags
   HSR_ASYNC = WININET_API_FLAG_ASYNC               ' force async
   HSR_SYNC = WININET_API_FLAG_SYNC                 ' force sync
   HSR_USE_CONTEXT = WININET_API_FLAG_USE_CONTEXT   ' use dwContext value
   HSR_INITIATE = &H8                              ' iterative operation (completed by HttpEndRequest)
   HSR_DOWNLOAD = &H10                             ' download to file
   HSR_CHUNKED = &H20                              ' operation is send of chunked data
End Enum

' HttpQueryInfo info levels. Generally, there is one info level
' for each potential RFC822/HTTP/MIME header that an HTTP server
' may send as part of a request response.
'
' The HTTP_QUERY_RAW_HEADERS info level is provided for clients
' that choose to perform their own header parsing.
Public Enum InfoLevels
   HTTP_QUERY_MIME_VERSION = 0
   HTTP_QUERY_CONTENT_TYPE = 1
   HTTP_QUERY_CONTENT_TRANSFER_ENCODING = 2
   HTTP_QUERY_CONTENT_ID = 3
   HTTP_QUERY_CONTENT_DESCRIPTION = 4
   HTTP_QUERY_CONTENT_LENGTH = 5
   HTTP_QUERY_CONTENT_LANGUAGE = 6
   HTTP_QUERY_ALLOW = 7
   HTTP_QUERY_PUBLIC = 8
   HTTP_QUERY_DATE = 9
   HTTP_QUERY_EXPIRES = 10
   HTTP_QUERY_LAST_MODIFIED = 11
   HTTP_QUERY_MESSAGE_ID = 12
   HTTP_QUERY_URI = 13
   HTTP_QUERY_DERIVED_FROM = 14
   HTTP_QUERY_COST = 15
   HTTP_QUERY_LINK = 16
   HTTP_QUERY_PRAGMA = 17
   HTTP_QUERY_VERSION = 18                      ' special: part of status line
   HTTP_QUERY_STATUS_CODE = 19                  ' special: part of status line
   HTTP_QUERY_STATUS_TEXT = 20                  ' special: part of status line
   HTTP_QUERY_RAW_HEADERS = 21                  ' special: all headers as ASCIIZ
   HTTP_QUERY_RAW_HEADERS_CRLF = 22             ' special: all headers
   HTTP_QUERY_CONNECTION = 23
   HTTP_QUERY_ACCEPT = 24
   HTTP_QUERY_ACCEPT_CHARSET = 25
   HTTP_QUERY_ACCEPT_ENCODING = 26
   HTTP_QUERY_ACCEPT_LANGUAGE = 27
   HTTP_QUERY_AUTHORIZATION = 28
   HTTP_QUERY_CONTENT_ENCODING = 29
   HTTP_QUERY_FORWARDED = 30
   HTTP_QUERY_FROM = 31
   HTTP_QUERY_IF_MODIFIED_SINCE = 32
   HTTP_QUERY_LOCATION = 33
   HTTP_QUERY_ORIG_URI = 34
   HTTP_QUERY_REFERER = 35
   HTTP_QUERY_RETRY_AFTER = 36
   HTTP_QUERY_SERVER = 37
   HTTP_QUERY_TITLE = 38
   HTTP_QUERY_USER_AGENT = 39
   HTTP_QUERY_WWW_AUTHENTICATE = 40
   HTTP_QUERY_PROXY_AUTHENTICATE = 41
   HTTP_QUERY_ACCEPT_RANGES = 42
   HTTP_QUERY_SET_COOKIE = 43
   HTTP_QUERY_COOKIE = 44
   HTTP_QUERY_REQUEST_METHOD = 45               ' special: GET/POST etc.
   HTTP_QUERY_REFRESH = 46
   HTTP_QUERY_CONTENT_DISPOSITION = 47

' HTTP 1.1 defined headers
   HTTP_QUERY_AGE = 48
   HTTP_QUERY_CACHE_CONTROL = 49
   HTTP_QUERY_CONTENT_BASE = 50
   HTTP_QUERY_CONTENT_LOCATION = 51
   HTTP_QUERY_CONTENT_MD5 = 52
   HTTP_QUERY_CONTENT_RANGE = 53
   HTTP_QUERY_ETAG = 54
   HTTP_QUERY_HOST = 55
   HTTP_QUERY_IF_MATCH = 56
   HTTP_QUERY_IF_NONE_MATCH = 57
   HTTP_QUERY_IF_RANGE = 58
   HTTP_QUERY_IF_UNMODIFIED_SINCE = 59
   HTTP_QUERY_MAX_FORWARDS = 60
   HTTP_QUERY_PROXY_AUTHORIZATION = 61
   HTTP_QUERY_RANGE = 62
   HTTP_QUERY_TRANSFER_ENCODING = 63
   HTTP_QUERY_UPGRADE = 64
   HTTP_QUERY_VARY = 65
   HTTP_QUERY_VIA = 66
   HTTP_QUERY_WARNING = 67
   HTTP_QUERY_EXPECT = 68
   HTTP_QUERY_PROXY_CONNECTION = 69
   HTTP_QUERY_UNLESS_MODIFIED_SINCE = 70
   HTTP_QUERY_ECHO_REQUEST = 71
   HTTP_QUERY_ECHO_REPLY = 72
   
   ' These are the set of headers that should be added back to a request when
   ' re-doing a request after a RETRY_WITH response.
   HTTP_QUERY_ECHO_HEADERS = 73
   HTTP_QUERY_ECHO_HEADERS_CRLF = 74
   HTTP_QUERY_MAX = 74

   ' HTTP_QUERY_CUSTOM - if this special value is supplied as the dwInfoLevel
   ' parameter of HttpQueryInfo() then the lpBuffer parameter contains the name
   ' of the header we are to query
   HTTP_QUERY_CUSTOM = 65535

   ' HTTP_QUERY_FLAG_REQUEST_HEADERS - if this bit is set in the dwInfoLevel
   ' parameter of HttpQueryInfo() then the request headers will be queried for the
   ' request information
   HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000

   ' HTTP_QUERY_FLAG_SYSTEMTIME - if this bit is set in the dwInfoLevel parameter
   ' of HttpQueryInfo() AND the header being queried contains date information,
   ' e.g. the "Expires:" header then lpBuffer will contain a SYSTEMTIME structure
   ' containing the date and time information converted from the header string
   HTTP_QUERY_FLAG_SYSTEMTIME = &H40000000

   ' HTTP_QUERY_FLAG_NUMBER - if this bit is set in the dwInfoLevel parameter of
   ' HttpQueryInfo(), then the value of the header will be converted to a number
   ' before being returned to the caller, if applicable
   HTTP_QUERY_FLAG_NUMBER = &H20000000

   ' HTTP_QUERY_FLAG_COALESCE - combine the values from several headers of the
   ' same name into the output buffer
   HTTP_QUERY_FLAG_COALESCE = &H10000000

   HTTP_QUERY_MODIFIER_FLAGS_MASK = HTTP_QUERY_FLAG_REQUEST_HEADERS Or HTTP_QUERY_FLAG_SYSTEMTIME Or HTTP_QUERY_FLAG_NUMBER Or HTTP_QUERY_FLAG_COALESCE
                                                
   HTTP_QUERY_HEADER_MASK = HTTP_QUERY_MODIFIER_FLAGS_MASK
End Enum

' Number of the TCP/IP port on the server to connect to.
Public Enum InternetDefaultPorts
   INTERNET_DEFAULT_FTP_PORT = 21
   INTERNET_DEFAULT_GOPHER_PORT = 70
   INTERNET_DEFAULT_HTTP_PORT = 80
   INTERNET_DEFAULT_HTTPS_PORT = 443
   INTERNET_DEFAULT_SOCKS_PORT = 1080
End Enum

Public Enum InternetReadFileExFlags
   IRF_ASYNC = WININET_API_FLAG_ASYNC
   IRF_SYNC = WININET_API_FLAG_SYNC
   IRF_USE_CONTEXT = WININET_API_FLAG_USE_CONTEXT
   IRF_NO_WAIT = &H8
End Enum

Public Enum ConnectDialogFlags
   CONNDLG_RO_PATH = &H1         'Resource path should be read -only
   CONNDLG_CONN_POINT = &H2      'Netware-style movable connection point enabled
   CONNDLG_USE_MRU = &H4         'Use MRU combobox
   CONNDLG_HIDE_BOX = &H8        'Hide persistent connect checkbox
   ' NOTE:  Set, at most, one of the below flags. If neither flag is
   '        set, then the persistence is set to whatever the user chose
   '        during a previous connection
   CONNDLG_PERSIST = &H10        'Force persistent connection
   CONNDLG_NOT_PERSIST = &H20    'Persistent not allowed
End Enum

Public Enum DisconnectDialogFlags
   DISC_UPDATE_PROFILE = &H1  ' If set, the connection is no longer persistent (automatically restored every time the user logs on).
   DISC_NO_FORCE = &H40       ' If not set, force will be applied when attempting to disconnect (typically because the user has
                              ' open files). The user will be informed if there are open files to the connection and asked if he
                              ' or she still wants to disconnect. If so, the disconnect procedure re-attempts with additional force
End Enum

Public Enum AlignRectsFlags
   CUDR_NORMAL = &H0
   CUDR_NOSNAPTOGRID = &H1
   CUDR_NORESOLVEPOSITIONS = &H2
   CUDR_NOCLOSEGAPS = &H4
   CUDR_NEGATIVECOORDS = &H8
   CUDR_NOPRIMARY = &H10
End Enum

Public Enum JOBOBJECTINFOCLASS
   JobObjectBasicAccountingInformation = 1&
   JobObjectBasicLimitInformation = 2&
   JobObjectBasicProcessIdList = 3&
   JobObjectBasicUIRestrictions = 4&
   JobObjectSecurityLimitInformation = 5&
   JobObjectEndOfJobTimeInformation = 6&
   JobObjectAssociateCompletionPortInformation = 7&
   JobObjectBasicAndIoAccountingInformation = 8&
   JobObjectExtendedLimitInformation = 9&
   MaxJobObjectInfoClass = 10&
End Enum

' GetComputerNameEx computer name types
Public Enum COMPUTER_NAME_FORMAT
   ComputerNameNetBIOS = 0&
   ComputerNameDnsHostname = 1&
   ComputerNameDnsDomain = 2&
   ComputerNameDnsFullyQualified = 3&
   ComputerNamePhysicalNetBIOS = 4&
   ComputerNamePhysicalDnsHostname = 5&
   ComputerNamePhysicalDnsDomain = 6&
   ComputerNamePhysicalDnsFullyQualified = 7&
   ComputerNameMax = 8&
End Enum

Public Enum PlatformId
   DLLVER_PLATFORM_WINDOWS = &H1   ' The DLL was built for all Windows platforms.
   DLLVER_PLATFORM_NT = &H2        ' The DLL was built specifically for Windows NT.
End Enum

' Bit masks for field usriX_flags of USER_INFO_0 / USER_INFO_1
Public Enum User_Info_0_1_Flags
   UF_SCRIPT = &H1
   UF_ACCOUNTDISABLE = &H2
   UF_HOMEDIR_REQUIRED = &H8
   UF_LOCKOUT = &H10
   UF_PASSWD_NOTREQD = &H20
   UF_PASSWD_CANT_CHANGE = &H40
   UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H80

   ' Account type bits as part of usri_flags.
   UF_TEMP_DUPLICATE_ACCOUNT = &H100
   UF_NORMAL_ACCOUNT = &H200
   UF_INTERDOMAIN_TRUST_ACCOUNT = &H800
   UF_WORKSTATION_TRUST_ACCOUNT = &H1000
   UF_SERVER_TRUST_ACCOUNT = &H2000


   UF_MACHINE_ACCOUNT_MASK = (UF_INTERDOMAIN_TRUST_ACCOUNT Or UF_WORKSTATION_TRUST_ACCOUNT Or UF_SERVER_TRUST_ACCOUNT)
   UF_ACCOUNT_TYPE_MASK = (UF_TEMP_DUPLICATE_ACCOUNT Or UF_NORMAL_ACCOUNT Or UF_INTERDOMAIN_TRUST_ACCOUNT Or UF_WORKSTATION_TRUST_ACCOUNT Or UF_SERVER_TRUST_ACCOUNT)

   UF_DONT_EXPIRE_PASSWD = &H10000
   UF_MNS_LOGON_ACCOUNT = &H20000
   UF_SMARTCARD_REQUIRED = &H40000
   UF_TRUSTED_FOR_DELEGATION = &H80000
   UF_NOT_DELEGATED = &H100000
   UF_USE_DES_KEY_ONLY = &H200000
   UF_DONT_REQUIRE_PREAUTH = &H400000
   
   UF_SETTABLE_BITS = (UF_SCRIPT Or UF_ACCOUNTDISABLE Or UF_LOCKOUT Or UF_HOMEDIR_REQUIRED Or UF_PASSWD_NOTREQD Or UF_PASSWD_CANT_CHANGE Or UF_ACCOUNT_TYPE_MASK Or UF_DONT_EXPIRE_PASSWD Or UF_MNS_LOGON_ACCOUNT Or UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED Or UF_SMARTCARD_REQUIRED Or UF_TRUSTED_FOR_DELEGATION Or UF_NOT_DELEGATED Or UF_USE_DES_KEY_ONLY Or UF_DONT_REQUIRE_PREAUTH)
End Enum

Public Enum AuthorizationFlags
   AF_OP_PRINT = &H1
   AF_OP_COMM = &H2
   AF_OP_SERVER = &H4
   AF_OP_ACCOUNTS = &H8
   AF_SETTABLE_BITS = (AF_OP_PRINT Or AF_OP_COMM Or AF_OP_SERVER Or AF_OP_ACCOUNTS)
End Enum

' UAS role manifests under NETLOGON
Public Enum NETLOGON_RoleManifests
   UAS_ROLE_STANDALONE = 0
   UAS_ROLE_MEMBER = 1
   UAS_ROLE_BACKUP = 2
   UAS_ROLE_PRIMARY = 3
End Enum

' Values for ParmError for NetUserSetInfo.
Public Enum NetUserSetInfoParmErrorValues
   USER_NAME_PARMNUM = 1
   USER_PASSWORD_PARMNUM = 3
   USER_PASSWORD_AGE_PARMNUM = 4
   USER_PRIV_PARMNUM = 5
   USER_HOME_DIR_PARMNUM = 6
   USER_COMMENT_PARMNUM = 7
   USER_FLAGS_PARMNUM = 8
   USER_SCRIPT_PATH_PARMNUM = 9
   USER_AUTH_FLAGS_PARMNUM = 10
   USER_FULL_NAME_PARMNUM = 11
   USER_USR_COMMENT_PARMNUM = 12
   USER_PARMS_PARMNUM = 13
   USER_WORKSTATIONS_PARMNUM = 14
   USER_LAST_LOGON_PARMNUM = 15
   USER_LAST_LOGOFF_PARMNUM = 16
   USER_ACCT_EXPIRES_PARMNUM = 17
   USER_MAX_STORAGE_PARMNUM = 18
   USER_UNITS_PER_WEEK_PARMNUM = 19
   USER_LOGON_HOURS_PARMNUM = 20
   USER_PAD_PW_COUNT_PARMNUM = 21
   USER_NUM_LOGONS_PARMNUM = 22
   USER_LOGON_SERVER_PARMNUM = 23
   USER_COUNTRY_CODE_PARMNUM = 24
   USER_CODE_PAGE_PARMNUM = 25
   USER_PRIMARY_GROUP_PARMNUM = 51
   USER_PROFILE = 52
   USER_PROFILE_PARMNUM = 52
   USER_HOME_DIR_DRIVE_PARMNUM = 53
   
   USER_NAME_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_NAME_PARMNUM)
   USER_PASSWORD_AGE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_PASSWORD_AGE_PARMNUM)
   USER_HOME_DIR_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_HOME_DIR_PARMNUM)
   USER_FLAGS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_FLAGS_PARMNUM)
   USER_AUTH_FLAGS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_AUTH_FLAGS_PARMNUM)
   USER_USR_COMMENT_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_USR_COMMENT_PARMNUM)
   USER_WORKSTATIONS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_WORKSTATIONS_PARMNUM)
   USER_LAST_LOGOFF_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_LAST_LOGOFF_PARMNUM)
   USER_MAX_STORAGE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_MAX_STORAGE_PARMNUM)
   USER_LOGON_HOURS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_LOGON_HOURS_PARMNUM)
   USER_NUM_LOGONS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_NUM_LOGONS_PARMNUM)
   USER_COUNTRY_CODE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_COUNTRY_CODE_PARMNUM)
   USER_PRIMARY_GROUP_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_PRIMARY_GROUP_PARMNUM)
   USER_HOME_DIR_DRIVE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + USER_HOME_DIR_DRIVE_PARMNUM)
End Enum

Public Enum AccountTypeFilters
   FILTER_NORMAL_ACCOUNT = &H2               ' Enumerates global user account data on a computer.
   FILTER_PROXY_ACCOUNT = &H4                '
   FILTER_INTERDOMAIN_TRUST_ACCOUNT = &H8    ' Enumerates domain trust account data on a domain controller.
   FILTER_WORKSTATION_TRUST_ACCOUNT = &H10   ' Enumerates workstation or member server account data on a domain controller.
   FILTER_SERVER_TRUST_ACCOUNT = &H20        ' Enumerates domain controller account data on a domain controller.

' ?? FILTER_TEMP_DUPLICATE_ACCOUNT Enumerates local user account data on a domain controller.
End Enum

Public Enum NetUserSetModalsParmErrorValues
   MODALS_MIN_PASSWD_LEN_PARMNUM = 1
   MODALS_MAX_PASSWD_AGE_PARMNUM = 2
   MODALS_MIN_PASSWD_AGE_PARMNUM = 3
   MODALS_FORCE_LOGOFF_PARMNUM = 4
   MODALS_PASSWD_HIST_LEN_PARMNUM = 5
   MODALS_ROLE_PARMNUM = 6
   MODALS_PRIMARY_PARMNUM = 7
   MODALS_DOMAIN_NAME_PARMNUM = 8
   MODALS_DOMAIN_ID_PARMNUM = 9
   MODALS_LOCKOUT_DURATION_PARMNUM = 10
   MODALS_LOCKOUT_OBSERVATION_WINDOW_PARMNUM = 11
   MODALS_LOCKOUT_THRESHOLD_PARMNUM = 12
   MODALS_MIN_PASSWD_LEN_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + MODALS_MIN_PASSWD_LEN_PARMNUM)
   MODALS_MIN_PASSWD_AGE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + MODALS_MIN_PASSWD_AGE_PARMNUM)
   MODALS_PASSWD_HIST_LEN_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + MODALS_PASSWD_HIST_LEN_PARMNUM)
   MODALS_PRIMARY_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + MODALS_PRIMARY_PARMNUM)
   MODALS_DOMAIN_ID_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + MODALS_DOMAIN_ID_PARMNUM)
End Enum

Public Enum NetServerSetInfoParmErrorValues
   SV_COMMENT_PARMNUM = 5
   SV_DISC_PARMNUM = 10
   SV_ALERTS_PARMNUM = 11
   SV_HIDDEN_PARMNUM = 16
   SV_ANNOUNCE_PARMNUM = 17
   SV_ANNDELTA_PARMNUM = 18
   SV_ALERTSCHED_PARMNUM = 37
   SV_ERRORALERT_PARMNUM = 38
   SV_LOGONALERT_PARMNUM = 39
   SV_ACCESSALERT_PARMNUM = 40
   SV_DISKALERT_PARMNUM = 41
   SV_NETIOALERT_PARMNUM = 42
   SV_MAXAUDITSZ_PARMNUM = 43
   SV_PLATFORM_ID_PARMNUM = 101
   SV_NAME_PARMNUM = 102
   SV_VERSION_MAJOR_PARMNUM = 103
   SV_VERSION_MINOR_PARMNUM = 104
   SV_TYPE_PARMNUM = 105
   SV_USERS_PARMNUM = 107
   SV_USERPATH_PARMNUM = 112
   SV_ULIST_MTIME_PARMNUM = 401
   SV_GLIST_MTIME_PARMNUM = 402
   SV_ALIST_MTIME_PARMNUM = 403
   SV_SECURITY_PARMNUM = 405
   SV_NUMADMIN_PARMNUM = 406
   SV_LANMASK_PARMNUM = 407
   SV_GUESTACC_PARMNUM = 408
   SV_CHDEVQ_PARMNUM = 410
   SV_CHDEVJOBS_PARMNUM = 411
   SV_CONNECTIONS_PARMNUM = 412
   SV_SHARES_PARMNUM = 413
   SV_OPENFILES_PARMNUM = 414
   SV_SESSREQS_PARMNUM = 417
   SV_ACTIVELOCKS_PARMNUM = 419
   SV_NUMREQBUF_PARMNUM = 420
   SV_NUMBIGBUF_PARMNUM = 422
   SV_NUMFILETASKS_PARMNUM = 423
   SV_SRVHEURISTICS_PARMNUM = 431
   SV_SESSOPENS_PARMNUM = 501
   SV_SESSVCS_PARMNUM = 502
   SV_OPENSEARCH_PARMNUM = 503
   SV_SIZREQBUF_PARMNUM = 504
   SV_INITWORKITEMS_PARMNUM = 505
   SV_MAXWORKITEMS_PARMNUM = 506
   SV_RAWWORKITEMS_PARMNUM = 507
   SV_IRPSTACKSIZE_PARMNUM = 508
   SV_MAXRAWBUFLEN_PARMNUM = 509
   SV_SESSUSERS_PARMNUM = 510
   SV_SESSCONNS_PARMNUM = 511
   SV_MAXNONPAGEDMEMORYUSAGE_PARMNUM = 512
   SV_MAXPAGEDMEMORYUSAGE_PARMNUM = 513
   SV_ENABLESOFTCOMPAT_PARMNUM = 514
   SV_ENABLEFORCEDLOGOFF_PARMNUM = 515
   SV_TIMESOURCE_PARMNUM = 516
   SV_ACCEPTDOWNLEVELAPIS_PARMNUM = 517
   SV_LMANNOUNCE_PARMNUM = 518
   SV_DOMAIN_PARMNUM = 519
   SV_MAXCOPYREADLEN_PARMNUM = 520
   SV_MAXCOPYWRITELEN_PARMNUM = 521
   SV_MINKEEPSEARCH_PARMNUM = 522
   SV_MAXKEEPSEARCH_PARMNUM = 523
   SV_MINKEEPCOMPLSEARCH_PARMNUM = 524
   SV_MAXKEEPCOMPLSEARCH_PARMNUM = 525
   SV_THREADCOUNTADD_PARMNUM = 526
   SV_NUMBLOCKTHREADS_PARMNUM = 527
   SV_SCAVTIMEOUT_PARMNUM = 528
   SV_MINRCVQUEUE_PARMNUM = 529
   SV_MINFREEWORKITEMS_PARMNUM = 530
   SV_XACTMEMSIZE_PARMNUM = 531
   SV_THREADPRIORITY_PARMNUM = 532
   SV_MAXMPXCT_PARMNUM = 533
   SV_OPLOCKBREAKWAIT_PARMNUM = 534
   SV_OPLOCKBREAKRESPONSEWAIT_PARMNUM = 535
   SV_ENABLEOPLOCKS_PARMNUM = 536
   SV_ENABLEOPLOCKFORCECLOSE_PARMNUM = 537
   SV_ENABLEFCBOPENS_PARMNUM = 538
   SV_ENABLERAW_PARMNUM = 539
   SV_ENABLESHAREDNETDRIVES_PARMNUM = 540
   SV_MINFREECONNECTIONS_PARMNUM = 541
   SV_MAXFREECONNECTIONS_PARMNUM = 542
   SV_INITSESSTABLE_PARMNUM = 543
   SV_INITCONNTABLE_PARMNUM = 544
   SV_INITFILETABLE_PARMNUM = 545
   SV_INITSEARCHTABLE_PARMNUM = 546
   SV_ALERTSCHEDULE_PARMNUM = 547
   SV_ERRORTHRESHOLD_PARMNUM = 548
   SV_NETWORKERRORTHRESHOLD_PARMNUM = 549
   SV_DISKSPACETHRESHOLD_PARMNUM = 550
   SV_MAXLINKDELAY_PARMNUM = 552
   SV_MINLINKTHROUGHPUT_PARMNUM = 553
   SV_LINKINFOVALIDTIME_PARMNUM = 554
   SV_SCAVQOSINFOUPDATETIME_PARMNUM = 555
   SV_MAXWORKITEMIDLETIME_PARMNUM = 556
   SV_MAXRAWWORKITEMS_PARMNUM = 557
   SV_PRODUCTTYPE_PARMNUM = 560
   SV_SERVERSIZE_PARMNUM = 561
   SV_CONNECTIONLESSAUTODISC_PARMNUM = 562
   SV_SHARINGVIOLATIONRETRIES_PARMNUM = 563
   SV_SHARINGVIOLATIONDELAY_PARMNUM = 564
   SV_MAXGLOBALOPENSEARCH_PARMNUM = 565
   SV_REMOVEDUPLICATESEARCHES_PARMNUM = 566
   SV_LOCKVIOLATIONRETRIES_PARMNUM = 567
   SV_LOCKVIOLATIONOFFSET_PARMNUM = 568
   SV_LOCKVIOLATIONDELAY_PARMNUM = 569
   SV_MDLREADSWITCHOVER_PARMNUM = 570
   SV_CACHEDOPENLIMIT_PARMNUM = 571
   SV_CRITICALTHREADS_PARMNUM = 572
   SV_RESTRICTNULLSESSACCESS_PARMNUM = 573
   SV_ENABLEWFW311DIRECTIPX_PARMNUM = 574
   SV_OTHERQUEUEAFFINITY_PARMNUM = 575
   SV_QUEUESAMPLESECS_PARMNUM = 576
   SV_BALANCECOUNT_PARMNUM = 577
   SV_PREFERREDAFFINITY_PARMNUM = 578
   SV_MAXFREERFCBS_PARMNUM = 579
   SV_MAXFREEMFCBS_PARMNUM = 580
   SV_MAXFREELFCBS_PARMNUM = 581
   SV_MAXFREEPAGEDPOOLCHUNKS_PARMNUM = 582
   SV_MINPAGEDPOOLCHUNKSIZE_PARMNUM = 583
   SV_MAXPAGEDPOOLCHUNKSIZE_PARMNUM = 584
   SV_SENDSFROMPREFERREDPROCESSOR_PARMNUM = 585
   SV_MAXTHREADSPERQUEUE_PARMNUM = 586
   SV_CACHEDDIRECTORYLIMIT_PARMNUM = 587
   SV_MAXCOPYLENGTH_PARMNUM = 588
   SV_ENABLEBULKTRANSFER_PARMNUM = 589
   SV_ENABLECOMPRESSION_PARMNUM = 590
   SV_AUTOSHAREWKS_PARMNUM = 591
   SV_AUTOSHARESERVER_PARMNUM = 592
   SV_ENABLESECURITYSIGNATURE_PARMNUM = 593
   SV_REQUIRESECURITYSIGNATURE_PARMNUM = 594
   SV_MINCLIENTBUFFERSIZE_PARMNUM = 595
   SV_CONNECTIONNOSESSIONSTIMEOUT_PARMNUM = 596
   SV_COMMENT_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_COMMENT_PARMNUM)
   SV_DISC_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_DISC_PARMNUM)
   SV_ANNOUNCE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_ANNOUNCE_PARMNUM)
   SV_SESSOPENS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_SESSOPENS_PARMNUM)
   SV_OPENSEARCH_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_OPENSEARCH_PARMNUM)
   SV_MAXRAWBUFLEN_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXRAWBUFLEN_PARMNUM)
   SV_SESSCONNS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_SESSCONNS_PARMNUM)
   SV_MAXPAGEDMEMORYUSAGE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXPAGEDMEMORYUSAGE_PARMNUM)
   SV_ENABLEFORCEDLOGOFF_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_ENABLEFORCEDLOGOFF_PARMNUM)
   SV_LMANNOUNCE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_LMANNOUNCE_PARMNUM)
   SV_MAXCOPYWRITELEN_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXCOPYWRITELEN_PARMNUM)
   SV_MAXKEEPSEARCH_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXKEEPSEARCH_PARMNUM)
   SV_MAXKEEPCOMPLSEARCH_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXKEEPCOMPLSEARCH_PARMNUM)
   SV_MINRCVQUEUE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MINRCVQUEUE_PARMNUM)
   SV_MAXMPXCT_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXMPXCT_PARMNUM)
   SV_OPLOCKBREAKRESPONSEWAIT_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_OPLOCKBREAKRESPONSEWAIT_PARMNUM)
   SV_ENABLEOPLOCKFORCECLOSE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_ENABLEOPLOCKFORCECLOSE_PARMNUM)
   SV_ENABLERAW_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_ENABLERAW_PARMNUM)
   SV_MINFREECONNECTIONS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MINFREECONNECTIONS_PARMNUM)
   SV_INITSESSTABLE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_INITSESSTABLE_PARMNUM)
   SV_INITFILETABLE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_INITFILETABLE_PARMNUM)
   SV_ALERTSCHEDULE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_ALERTSCHEDULE_PARMNUM)
   SV_NETWORKERRORTHRESHOLD_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_NETWORKERRORTHRESHOLD_PARMNUM)
   SV_MAXLINKDELAY_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXLINKDELAY_PARMNUM)
   SV_LINKINFOVALIDTIME_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_LINKINFOVALIDTIME_PARMNUM)
   SV_MAXWORKITEMIDLETIME_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXWORKITEMIDLETIME_PARMNUM)
   SV_PRODUCTTYPE_INFOLOEVEL = (PARMNUM_BASE_INFOLEVEL + SV_PRODUCTTYPE_PARMNUM)
   SV_CONNECTIONLESSAUTODISC_INFOLOEVEL = (PARMNUM_BASE_INFOLEVEL + SV_CONNECTIONLESSAUTODISC_PARMNUM)
   SV_SHARINGVIOLATIONDELAY_INFOLOEVEL = (PARMNUM_BASE_INFOLEVEL + SV_SHARINGVIOLATIONDELAY_PARMNUM)
   SV_REMOVEDUPLICATESEARCHES_INFOLOEVEL = (PARMNUM_BASE_INFOLEVEL + SV_REMOVEDUPLICATESEARCHES_PARMNUM)
   SV_LOCKVIOLATIONOFFSET_INFOLOEVEL = (PARMNUM_BASE_INFOLEVEL + SV_LOCKVIOLATIONOFFSET_PARMNUM)
   SV_MDLREADSWITCHOVER_INFOLOEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MDLREADSWITCHOVER_PARMNUM)
   SV_CRITICALTHREADS_INFOLOEVEL = (PARMNUM_BASE_INFOLEVEL + SV_CRITICALTHREADS_PARMNUM)
   SV_ENABLEWFW311DIRECTIPX_INFOLOEVEL = (PARMNUM_BASE_INFOLEVEL + SV_ENABLEWFW311DIRECTIPX_PARMNUM)
   SV_QUEUESAMPLESECS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_QUEUESAMPLESECS_PARMNUM)
   SV_PREFERREDAFFINITY_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_PREFERREDAFFINITY_PARMNUM)
   SV_MAXFREEMFCBS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXFREEMFCBS_PARMNUM)
   SV_MAXFREEPAGEDPOOLCHUNKS_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXFREEPAGEDPOOLCHUNKS_PARMNUM)
   SV_MAXPAGEDPOOLCHUNKSIZE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXPAGEDPOOLCHUNKSIZE_PARMNUM)
   SV_MAXTHREADSPERQUEUE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXTHREADSPERQUEUE_PARMNUM)
   SV_MAXCOPYLENGTH_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_MAXCOPYLENGTH_PARMNUM)
   SV_ENABLECOMPRESSION_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_ENABLECOMPRESSION_PARMNUM)
   SV_AUTOSHARESERVER_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_AUTOSHARESERVER_PARMNUM)
   SV_REQUIRESECURITYSIGNATURE_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_REQUIRESECURITYSIGNATURE_PARMNUM)
   SV_CONNECTIONNOSESSIONSTIMEOUT_INFOLEVEL = (PARMNUM_BASE_INFOLEVEL + SV_CONNECTIONNOSESSIONSTIMEOUT_PARMNUM)
End Enum

Public Enum ServerTypes
   SV_TYPE_WORKSTATION = &H1
   SV_TYPE_SERVER = &H2
   SV_TYPE_SQLSERVER = &H4
   SV_TYPE_DOMAIN_CTRL = &H8
   SV_TYPE_DOMAIN_BAKCTRL = &H10
   SV_TYPE_TIME_SOURCE = &H20
   SV_TYPE_AFP = &H40
   SV_TYPE_NOVELL = &H80
   SV_TYPE_DOMAIN_MEMBER = &H100
   SV_TYPE_PRINTQ_SERVER = &H200
   SV_TYPE_DIALIN_SERVER = &H400
   SV_TYPE_XENIX_SERVER = &H800
   SV_TYPE_SERVER_UNIX = SV_TYPE_XENIX_SERVER
   SV_TYPE_NT = &H1000
   SV_TYPE_WFW = &H2000
   SV_TYPE_SERVER_MFPN = &H4000
   SV_TYPE_SERVER_NT = &H8000
   SV_TYPE_POTENTIAL_BROWSER = &H10000
   SV_TYPE_BACKUP_BROWSER = &H20000
   SV_TYPE_MASTER_BROWSER = &H40000
   SV_TYPE_DOMAIN_MASTER = &H80000
   SV_TYPE_SERVER_OSF = &H100000
   SV_TYPE_SERVER_VMS = &H200000
   SV_TYPE_WINDOWS = &H400000                ' Windows95 and above
   SV_TYPE_DFS = &H800000                    ' Root of a DFS tree
   SV_TYPE_CLUSTER_NT = &H1000000            ' NT Cluster
   SV_TYPE_DCE = &H10000000                  ' IBM DSS (Directory and Security Services) or equivalent
   SV_TYPE_ALTERNATE_XPORT = &H20000000      ' return list for alternate transport
   SV_TYPE_LOCAL_LIST_ONLY = &H40000000      ' Return local list only
   SV_TYPE_DOMAIN_ENUM = &H80000000
   SV_TYPE_ALL = &HFFFFFFFF                  ' handy for NetServerEnum2
End Enum

' Option bits for options parameter
Public Enum ShFormatOptions
   SHFMT_OPT_FULL = &H1
   SHFMT_OPT_SYSONLY = &H2
End Enum

' Special return values. PLEASE NOTE that these are DWORD values.
Public Enum ShFormatReturnValues
   SHFMT_ERROR = &HFFFFFFFF        ' Error on last format, drive may be formatable
   SHFMT_CANCEL = &HFFFFFFFE       ' Last format was canceled
   SHFMT_NOFORMAT = &HFFFFFFFD     ' Drive is not formatable
End Enum

Public Enum AddPrinterDriverExFlags 'Win2000
   APD_STRICT_UPGRADE = &H1    ' Add the printer driver only if all the files in the printer-driver directory are newer than any corresponding files currently in use.
   APD_STRICT_DOWNGRADE = &H2  ' Add the printer driver only if all the files in the printer-driver directory are older than any corresponding files currently in use.
   APD_COPY_ALL_FILES = &H4    ' Add the printer driver and copy all the files in the printer-driver directory.
   APD_COPY_NEW_FILES = &H8    ' Add the printer driver and copy the files in the printer-driver directory that are newer than any corresponding files that are currently in use.
End Enum

Public Enum MULTIPLE_TRUSTEE_OPERATION
  NO_MULTIPLE_TRUSTEE = 0
  TRUSTEE_IS_IMPERSONATE
End Enum

Public Enum TRUSTEE_FORM
  TRUSTEE_IS_SID = 0&
  TRUSTEE_IS_NAME
  TRUSTEE_BAD_FORM
  TRUSTEE_IS_OBJECTS_AND_SID
  TRUSTEE_IS_OBJECTS_AND_NAME
End Enum

Public Enum TRUSTEE_TYPE
  TRUSTEE_IS_UNKNOWN
  TRUSTEE_IS_USER
  TRUSTEE_IS_GROUP
  TRUSTEE_IS_DOMAIN
  TRUSTEE_IS_ALIAS
  TRUSTEE_IS_WELL_KNOWN_GROUP
  TRUSTEE_IS_DELETED
  TRUSTEE_IS_INVALID
  TRUSTEE_IS_COMPUTER
End Enum

Public Enum WELL_KNOWN_SID_TYPE
   WinNullSid = 0&
   WinWorldSid = 1&
   WinLocalSid = 2&
   WinCreatorOwnerSid = 3&
   WinCreatorGroupSid = 4&
   WinCreatorOwnerServerSid = 5&
   WinCreatorGroupServerSid = 6&
   WinNtAuthoritySid = 7&
   WinDialupSid = 8&
   WinNetworkSid = 9&
   WinBatchSid = 10&
   WinInteractiveSid = 11&
   WinServiceSid = 12&
   WinAnonymousSid = 13&
   WinProxySid = 14&
   WinEnterpriseControllersSid = 15&
   WinSelfSid = 16&
   WinAuthenticatedUserSid = 17&
   WinRestrictedCodeSid = 18&
   WinTerminalServerSid = 19&
   WinRemoteLogonIdSid = 20&
   WinLogonIdsSid = 21&
   WinLocalSystemSid = 22&
   WinLocalServiceSid = 23&
   WinNetworkServiceSid = 24&
   WinBuiltinDomainSid = 25&
   WinBuiltinAdministratorsSid = 26&
   WinBuiltinUsersSid = 27&
   WinBuiltinGuestsSid = 28&
   WinBuiltinPowerUsersSid = 29&
   WinBuiltinAccountOperatorsSid = 30&
   WinBuiltinSystemOperatorsSid = 31&
   WinBuiltinPrintOperatorsSid = 32&
   WinBuiltinBackupOperatorsSid = 33&
   WinBuiltinReplicatorSid = 34&
   WinBuiltinPreWindows2000CompatibleAccessSid = 35&
   WinBuiltinRemoteDesktopUsersSid = 36&
   WinBuiltinNetworkConfigurationOperatorsSid = 37&
   WinAccountAdministratorSid = 38&
   WinAccountGuestSid = 39&
   WinAccountKrbtgtSid = 40&
   WinAccountDomainAdminsSid = 41&
   WinAccountDomainUsersSid = 42&
   WinAccountDomainGuestsSid = 43&
   WinAccountComputersSid = 44&
   WinAccountControllersSid = 45&
   WinAccountCertAdminsSid = 46&
   WinAccountSchemaAdminsSid = 47&
   WinAccountEnterpriseAdminsSid = 48&
   WinAccountPolicyAdminsSid = 49&
   WinAccountRasAndIasServersSid = 50&
End Enum

Public Enum SecurityInformationBitFlags
   OWNER_SECURITY_INFORMATION = &H1
   GROUP_SECURITY_INFORMATION = &H2
   DACL_SECURITY_INFORMATION = &H4
   SACL_SECURITY_INFORMATION = &H8

   PROTECTED_DACL_SECURITY_INFORMATION = &H80000000
   PROTECTED_SACL_SECURITY_INFORMATION = &H40000000
   UNPROTECTED_DACL_SECURITY_INFORMATION = &H20000000
   UNPROTECTED_SACL_SECURITY_INFORMATION = &H10000000
End Enum

Public Enum JobObjectCompletionPortsMessages
   JOB_OBJECT_MSG_END_OF_JOB_TIME = 1
   JOB_OBJECT_MSG_END_OF_PROCESS_TIME = 2
   JOB_OBJECT_MSG_ACTIVE_PROCESS_LIMIT = 3
   JOB_OBJECT_MSG_ACTIVE_PROCESS_ZERO = 4
   JOB_OBJECT_MSG_NEW_PROCESS = 6
   JOB_OBJECT_MSG_EXIT_PROCESS = 7
   JOB_OBJECT_MSG_ABNORMAL_EXIT_PROCESS = 8
   JOB_OBJECT_MSG_PROCESS_MEMORY_LIMIT = 9
   JOB_OBJECT_MSG_JOB_MEMORY_LIMIT = 10
End Enum

Public Enum JobObjectBasicLimits
   JOB_OBJECT_LIMIT_WORKINGSET = &H1
   JOB_OBJECT_LIMIT_PROCESS_TIME = &H2
   JOB_OBJECT_LIMIT_JOB_TIME = &H4
   JOB_OBJECT_LIMIT_ACTIVE_PROCESS = &H8
   JOB_OBJECT_LIMIT_AFFINITY = &H10
   JOB_OBJECT_LIMIT_PRIORITY_CLASS = &H20
   JOB_OBJECT_LIMIT_PRESERVE_JOB_TIME = &H40
   JOB_OBJECT_LIMIT_SCHEDULING_CLASS = &H80
End Enum

Public Enum JobObjectExtendedLimits
   JOB_OBJECT_LIMIT_PROCESS_MEMORY = &H100
   JOB_OBJECT_LIMIT_JOB_MEMORY = &H200
   JOB_OBJECT_LIMIT_DIE_ON_UNHANDLED_EXCEPTION = &H400
   JOB_OBJECT_LIMIT_BREAKAWAY_OK = &H800
   JOB_OBJECT_LIMIT_SILENT_BREAKAWAY_OK = &H1000

   JOB_OBJECT_LIMIT_RESERVED1 = &H2000
   JOB_OBJECT_LIMIT_RESERVED2 = &H4000
   JOB_OBJECT_LIMIT_RESERVED3 = &H8000
   JOB_OBJECT_LIMIT_RESERVED4 = &H10000
   JOB_OBJECT_LIMIT_RESERVED5 = &H20000
   JOB_OBJECT_LIMIT_RESERVED6 = &H40000
End Enum

Public Enum SE_OBJECT_TYPE
   SE_UNKNOWN_OBJECT_TYPE = 0
   SE_FILE_OBJECT
   SE_SERVICE
   SE_PRINTER
   SE_REGISTRY_KEY
   SE_LMSHARE
   SE_KERNEL_OBJECT
   SE_WINDOW_OBJECT
   SE_DS_OBJECT
   SE_DS_OBJECT_ALL
   SE_PROVIDER_DEFINED_OBJECT
   SE_WMIGUID_OBJECT
   SE_REGISTRY_WOW64_32KEY
End Enum

Public Enum COLORTYPE
   COLOR_GRAY = 1
   COLOR_RGB
   COLOR_XYZ
   COLOR_Yxy
   COLOR_Lab
   COLOR_3_CHANNEL
   COLOR_CMYK
   COLOR_5_CHANNEL
   COLOR_6_CHANNEL
   COLOR_7_CHANNEL
   COLOR_8_CHANNEL
   COLOR_NAMED
End Enum

Public Enum SYSGEOCLASS
  GEOCLASS_NATION = 16&
  GEOCLASS_REGION = 14&
End Enum

Public Enum SYSGEOTYPE
  GEO_NATION = &H1&
  GEO_LATITUDE = &H2&
  GEO_LONGITUDE = &H3&
  GEO_ISO2 = &H4&
  GEO_ISO3 = &H5&
  GEO_RFC1766 = &H6&
  GEO_LCID = &H7&
  GEO_FRIENDLYNAME = &H8&
  GEO_OFFICIALNAME = &H9&
  GEO_TIMEZONES = &HA&
  GEO_OFFICIALLANGUAGES = &HB&
End Enum

Public Enum GET_FILEEX_INFO_LEVELS
   GetFileExInfoStandard = 0&
   GetFileExMaxInfoLevel = 1&
End Enum

Public Enum FINDEX_INFO_LEVELS
   FindExInfoStandard = 0&
End Enum

Public Enum FINDEX_SEARCH_OPS
  FindExSearchNameMatch = 0&
  FindExSearchLimitToDirectories = 1&
  FindExSearchLimitToDevices = 2&
End Enum

Public Enum HEAP_INFORMATION_CLASS
   HeapCompatibilityInformation = 0
End Enum

Public Enum ACCESS_MODE
   NOT_USED_ACCESS = 0&
   GRANT_ACCESS
   SET_ACCESS
   DENY_ACCESS
   REVOKE_ACCESS
   SET_AUDIT_SUCCESS
   SET_AUDIT_FAILURE
End Enum

Public Enum SID_NAME_USE
   SidTypeUser = 1&
   SidTypeGroup
   SidTypeDomain
   SidTypeAlias
   SidTypeWellKnownGroup
   SidTypeDeletedAccount
   SidTypeInvalid
   SidTypeUnknown
   SidTypeComputer
End Enum

Public Enum POLICY_INFORMATION_CLASS
   PolicyAuditLogInformation = 1&
   PolicyAuditEventsInformation
   PolicyPrimaryDomainInformation
   PolicyPdAccountInformation
   PolicyAccountDomainInformation
   PolicyLsaServerRoleInformation
   PolicyReplicaSourceInformation
   PolicyDefaultQuotaInformation
   PolicyModificationInformation
   PolicyAuditFullSetInformation
   PolicyAuditFullQueryInformation
   PolicyDnsDomainInformation
End Enum

Public Enum TRUSTED_INFORMATION_CLASS
   TrustedDomainNameInformation = 1&
   TrustedControllersInformation
   TrustedPosixOffsetInformation
   TrustedPasswordInformation
   TrustedDomainInformationBasic
   TrustedDomainInformationEx
   TrustedDomainAuthInformation
   TrustedDomainFullInformation
End Enum

Public Enum ACETypes
   ACCESS_MIN_MS_ACE_TYPE = &H0
   ACCESS_ALLOWED_ACE_TYPE = &H0
   ACCESS_DENIED_ACE_TYPE = &H1
   SYSTEM_AUDIT_ACE_TYPE = &H2
   SYSTEM_ALARM_ACE_TYPE = &H3
   ACCESS_MAX_MS_V2_ACE_TYPE = &H3

   ACCESS_ALLOWED_COMPOUND_ACE_TYPE = &H4
   ACCESS_MAX_MS_V3_ACE_TYPE = &H4

   ACCESS_MIN_MS_OBJECT_ACE_TYPE = &H5
   ACCESS_ALLOWED_OBJECT_ACE_TYPE = &H5
   ACCESS_DENIED_OBJECT_ACE_TYPE = &H6
   SYSTEM_AUDIT_OBJECT_ACE_TYPE = &H7
   SYSTEM_ALARM_OBJECT_ACE_TYPE = &H8
   ACCESS_MAX_MS_OBJECT_ACE_TYPE = &H8

   ACCESS_MAX_MS_V4_ACE_TYPE = &H8
   ACCESS_MAX_MS_ACE_TYPE = &H8
End Enum

Public Enum CRED_MARSHAL_TYPE
   CertCredential& = 1
   UsernameTargetCredential&
End Enum

Public Enum SAFER_OBJECT_INFO_CLASS
    SaferObjectLevelId& = 1                ' get: DWORD
    SaferObjectScopeId&                    ' get: DWORD
    SaferObjectFriendlyName&               ' get/set: LPCWSTR
    SaferObjectDescription&                ' get/set: LPCWSTR
    SaferObjectBuiltin&                    ' get: DWORD boolean

    SaferObjectDisallowed&                 ' get: DWORD boolean
    SaferObjectDisableMaxPrivilege&        ' get: DWORD boolean
    SaferObjectInvertDeletedPrivileges&    ' get: DWORD boolean
    SaferObjectDeletedPrivileges&          ' get: TOKEN_PRIVILEGES
    SaferObjectDefaultOwner&               ' get: TOKEN_OWNER
    SaferObjectSidsToDisable&              ' get: TOKEN_GROUPS
    SaferObjectRestrictedSidsInverted&     ' get: TOKEN_GROUPS
    SaferObjectRestrictedSidsAdded&        ' get: TOKEN_GROUPS

    ' To enumerate all identities, call GetInfo with SaferObjectAllIdentificationGuids.
    
    SaferObjectAllIdentificationGuids&     ' get: SAFER_IDENTIFICATION_GUIDS

    ' To create a new identity, call SetInfo with
    '      SaferObjectSingleIdentification with a new unique GUID that you have generated.
    ' To get details on a single identity, call GetInfo with
    '      SaferObjectSingleIdentification with desired GUID.
    ' To modify details of a single identity, call SetInfo with
    '      SaferObjectSingleIdentification with desired info and GUID.
    ' To delete an identity, call SetInfo with
    '      SaferObjectSingleIdentification with the header.dwIdentificationType set to 0.
    
    SaferObjectSingleIdentification&       ' get/set: SAFER_IDENTIFICATION_*
    SaferObjectExtendedError&              ' get: DWORD dwError
End Enum

Public Enum SAFER_POLICY_INFO_CLASS
  SaferPolicyLevelList& = 1
  SaferPolicyEnableTransparentEnforcement&
  SaferPolicyDefaultLevel&
  SaferPolicyEvaluateUserScope&
  SaferPolicyScopeFlags&
End Enum

 
 