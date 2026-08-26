VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "User_Info_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type User3           ' Level 3
                            ' User_Info_3
    name As String
    password As String
    password_last_set As String
    priv As String
    home_dir As String
    comment As String
    flags As Long
    script_path As String
    auth_flags As Long
    full_name As String
    usr_comment As String
    parms As String
    workstations_allowed As String
    last_logon As String
    last_logoff As String
    acct_expires As String
    max_storage As String
    units_per_week As Long
    logon_hours As Long         'Bad type choice
    bad_pw_count As Long
    num_logons As Long
    logon_server As String
    country_code As Long
    code_page As Long
    user_id As Long
    primary_group_id As Long
    profile As String
    home_dir_drive As String
    password_expired As String
End Type

Private User As User3
    
Public Property Let name(str As String)
  User.name = str
End Property

Public Property Let password(str As String)
  User.password = str
End Property

Public Property Let password_last_set(str As String)
  User.password_last_set = str
End Property

Public Property Let priv(str As String)
  User.priv = str
End Property

Public Property Let home_dir(str As String)
  User.home_dir = str
End Property

Public Property Let comment(str As String)
  User.comment = str
End Property

Public Property Let flags(lng As Long)
  User.flags = lng
End Property

Public Property Let script_path(str As String)
  User.script_path = str
End Property

Public Property Let auth_flags(lng As Long)
  User.auth_flags = lng
End Property
 
Public Property Let full_name(str As String)
  User.full_name = str
End Property
 
Public Property Let usr_comment(str As String)
  User.usr_comment = str
End Property
 
Public Property Let parms(str As String)
  User.parms = str
End Property
 
Public Property Let workstations_allowed(str As String)
  User.workstations_allowed = str
End Property
 
Public Property Let last_logon(str As String)
  User.last_logon = str
End Property
 
Public Property Let last_logoff(str As String)
  User.last_logoff = str
End Property
 
Public Property Let acct_expires(str As String)
  User.acct_expires = str
End Property
 
Public Property Let max_storage(str As String)
  User.max_storage = str
End Property
 
Public Property Let units_per_week(lng As Long)
  User.units_per_week = lng
End Property
 
Public Property Let logon_hours(lng As Long)
  User.logon_hours = lng
End Property
 
Public Property Let bad_pw_count(lng As Long)
  User.bad_pw_count = lng
End Property
 
Public Property Let num_logons(lng As Long)
  User.num_logons = lng
End Property

Public Property Let logon_server(str As String)
  User.logon_server = str
End Property

Public Property Let country_code(lng As Long)
  User.country_code = lng
End Property

Public Property Let code_page(lng As Long)
  User.code_page = lng
End Property

Public Property Let user_id(lng As Long)
  User.user_id = lng
End Property
 
Public Property Let primary_group_id(lng As Long)
  User.primary_group_id = lng
End Property
 
Public Property Let profile(str As String)
  User.profile = str
End Property
 
Public Property Let home_dir_drive(str As String)
  User.home_dir_drive = str
End Property
 
Public Property Let password_expired(str As String)
  User.password_expired = str
End Property

'**************************************************

Public Property Get name() As String
  name = User.name
End Property

Public Property Get password() As String
  password = User.password
End Property

Public Property Get password_last_set() As String
  password_last_set = User.password_last_set
End Property

Public Property Get priv() As String
  priv = User.priv
End Property

Public Property Get home_dir() As String
  home_dir = User.home_dir
End Property

Public Property Get comment() As String
  comment = User.comment
End Property

Public Property Get flags() As Long
  flags = User.flags
End Property

Public Property Get script_path() As String
  script_path = User.script_path
End Property

Public Property Get auth_flags() As Long
  auth_flags = User.auth_flags
End Property
 
Public Property Get full_name() As String
  full_name = User.full_name
End Property
 
Public Property Get usr_comment() As String
  usr_comment = User.usr_comment
End Property
 
Public Property Get parms() As String
  parms = User.parms
End Property
 
Public Property Get workstations_allowed() As String
  workstations_allowed = User.workstations_allowed
End Property
 
Public Property Get last_logon() As String
  last_logon = User.last_logon
End Property
 
Public Property Get last_logoff() As String
  last_logoff = User.last_logoff
End Property
 
Public Property Get acct_expires() As String
  acct_expires = User.acct_expires
End Property
 
Public Property Get max_storage() As String
  max_storage = User.max_storage
End Property
 
Public Property Get units_per_week() As Long
  units_per_week = User.units_per_week
End Property
 
Public Property Get logon_hours() As Long
  logon_hours = User.logon_hours
End Property
 
Public Property Get bad_pw_count() As Long
  bad_pw_count = User.bad_pw_count
End Property
 
Public Property Get num_logons() As Long
  num_logons = User.num_logons
End Property

Public Property Get logon_server() As String
  logon_server = User.logon_server
End Property

Public Property Get country_code() As Long
  country_code = User.country_code
End Property

Public Property Get code_page() As Long
  code_page = User.code_page
End Property

Public Property Get user_id() As Long
  user_id = User.user_id
End Property
 
Public Property Get primary_group_id() As Long
  primary_group_id = User.primary_group_id
End Property
 
Public Property Get profile() As String
  profile = User.profile
End Property
 
Public Property Get home_dir_drive() As String
  home_dir_drive = User.home_dir_drive
End Property
 
Public Property Get password_expired() As String
  password_expired = User.password_expired
End Property
 
 
 
 
