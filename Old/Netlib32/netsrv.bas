'Name:		NetSrv.BAS
'Description:	Network Server & Domain Controller API Functions for Windows NT
'Dependencies:	NetDecs.BAS
'		NetErr.BAS
'Included
'Functions:	NetGetPDC
'		
'------------------------------------------------------------------------

Public Function NetGetPDC(ByVal ServerName$, ByVal DomainName$) As String
'THIS FUNCTION RETRIEVES THE PDC'S COMPUTER NAME IF YOU GIVE IT THE DOMAIN NAME AND AT LEAST ONE VALID ENTRUSTED SERVER COMPUTER NAME
'THIS FUNCTION MIGHT ALSO NOT REQUIRE SERVERNAME--I'VE NOT TESTED IT THAT WAY.
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.


Dim arrayDomainName() As Byte, arrayServerName() As Byte, arrayPDC(100) As Byte

    TripError = False
    
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    Call ConvStringToByteArray(DomainName$, arrayDomainName)
  
    Reply& = NetGetDCName(arrayServerName(0), arrayDomainName(0), ptrDCN&)
    Call NetReturnError(Reply&, DomainName$)
  
    Call ConvPointerToArray(ptrDCN&, arrayPDC)
    ConvPointerToNothing (ptrDCN&)
    
    PDCName$ = arrayPDC()
    NetGetPDC = PDCName$
End Function