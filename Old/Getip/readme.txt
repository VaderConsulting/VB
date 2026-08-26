Since I had such a difficult time finding this information, I'm going to
share with the rest of you. Attached is GetIP.CLS. It is a very simple class
module that will calculate the IP address of a host name written in VB6.

The original code was written by David Diedrich and posted at:
http://www.netfokus.dk/vbadmincode/

The original example only returned the host name and IP address of the local
machine. I modified it to return the IP of any named host and wrapped it in
a class. Simply set the HostName property then retrieve the IPAddress
property. Hope it helps you as much as it did me.

