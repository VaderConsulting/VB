An example of how to read the Windows NT event viewer on local and remote
machine.
This does not return the whole string but only the string parameters (%1,
%2, ecc) which are the only pieces of information useful in automatic jobs
that read event viewer.

I've included this code in a billing application that read windows nt RAS
and RADIUS logs information.

This code is simple to understand and quick to execute.
