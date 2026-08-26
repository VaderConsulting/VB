WNet Example in VB6
===================

This is an example of using the WNet APIs to reproduce a kind of Network Neighborhood in VB6 for both Windows 9x and Windows NT (2000!)

It is an example so is not complete, but allows you to navigate down through the tree from Network Provider (ie Microsoft or Netware) to share and directory level. It will not navigate back up, but you should get the picture from what is there.

The main function is fnEnum which does the work.

If called passing an empty string, the enumeration starts at the top level. If called with the correct parameters, the enumeration is of the object to which the parameters correspond.

The function returns the object names and the parameters associated with each object in two string arrays.

I place the object names (ie Node or share names) into a ListView control and put the parameters into the ListView as the Key for each object.

When you then double click the ListView on an object, I simply call fnEnum and pass the key for the object.....

Cheers,

Andy

andy.doran@dial.pipex.com
