Here is a VB5 ActiveX DLL that demonstrates many of the WNet and LANMAN  functions and wraps them in OLE - er, ActiveX objects.  Included also is a sample EXE that demonstrates how to use them. 

The code has been created, cut-and-pasted and gathered over a span of time.  The naming conventions, therefore, are not necessarily consistent.  It is intended for educational purposes only.

Bob Hyland
Client Server Solutions
(formerly Not-Quite-Ready Software)

Here is a VB5 sample which demos some of the WNet and LANMAN functions under Win32.  There is a class library - NOblets - implemented in an ActiveX DLL which demos enumeration of domains and servers.  For each server, it reveals shares, users, groups and services.  I had tried also NetEnumFile, but could not get that to work.  Nor could I get NetEnumLocalGroups to work consistently.  If someone does, I would appreciate the update.

Bob Hyland
Comcast Cellular
hylandr@comcell.com
rjhyland@ptdprolog.net
