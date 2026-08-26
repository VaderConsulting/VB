VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create DDR's"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create DDR's"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
    Dim DSN As String
    Dim SQL As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim adoRS2 As ADODB.Recordset
    Dim HostName As String, IP As String, Subnet As String, Sitecode As String
    Dim i As Integer
    Dim DDR As New SMSResGen
        
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    Set adoRS2 = CreateObject("ADODB.Recordset")
    
    'DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Domain.mdb;Persist Security Info=False"
    DSN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Domain Control;Data Source=(Local)"
    adoConn.Open DSN
    
    SQL = "SELECT * FROM tblHosts WHERE IP <> '' AND IP <> 'NULL' AND Type <> '' AND Type <> 'NULL' ORDER BY Hostname"
    
    adoRS.Open SQL, adoConn
    Do Until adoRS.EOF
        HostName = adoRS("Hostname")
        IP = adoRS("IP")
        Debug.Print HostName & " " & IP
        
        i = InStrRev(IP, ".")
        Subnet = Left(IP, i - 1) & ".0"
        
        Sitecode = "WSS"
        
        DDR.DDRNew "SYSTEM", "SMS_AD_DISCOVERY", Sitecode
        
        'DDR.DDRAddInteger "Client", Resource.Client, ADDPROP_NONE
        'DDR.DDRAddString "Client Version", Resource.ClientVersion, 15, ADDPROP_NONE
        'DDR.DDRAddStringArray "IP Addresses", CStrings(Resource.IPAddresses), 64, ADDPROP_NONE
        'DDR.DDRAddStringArray "IP Subnets", CStrings(Resource.IPSubnets), 64, ADDPROP_NONE
        DDR.DDRAddString "Netbios Name", HostName, 32, ADDPROP_NAME
        DDR.DDRAddString "Operating System Name and Version", adoRS("Type"), 64, ADDPROP_NONE
        DDR.DDRAddString "Resource Domain OR Workgroup", adoRS("Domainname"), 64, ADDPROP_NONE
        
        'The new property that is being added.
        'DDR.DDRAddString "Organizational Unit", "<organizational unit>", 64, ADDPROP_NONE
        
        DDR.DDRWrite "c:\SMS" & HostName & ".txt"
        DDR.DDRSendToSMS

        
        ' No discovery data
        
        
        
        adoRS.MoveNext
    Loop
    
    adoConn.Close
    
    Set oSMS = Nothing
    
    Set adoRS = Nothing
    Set adoRS2 = Nothing
    Set adoConn = Nothing
End Sub
