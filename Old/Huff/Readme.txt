These programs were built and tested with VB 4/32 and Windows 95.  They have
been tested under VB 5.0 with Windows 95.

These two programs are not, and never were intended to be finished products.
They were written to understand the concepts of the Huffman algorithm. 
You will readily note that there is virtually no error checking.  
The program is not a speed burner.  I leave it to you to take this program 
and improve it to your hearts delight.

COMPRESSION PROGRAM.
You have a choice as to how to enter the input file path.
	1.  Type directly into the Input path text box.
	2.  Use the file menu (Open Input File).
The output path text box will have an option based on the path and filename
of the input path.  To use this option just press the compress command 
button.  You could also type directly into the text box or use the file menu
(Open Compressed File).
The lstCounts list box will contain the raw counts from the input file.  
The Print Counts check box will give you a hard copy of these counts.
The lstWeights list box will display the scaled counts.  
The Print Nodes check box will give a hard copy of the nodes in the binary 
tree and the scaled weights.
Selecting an item in either list box will cause the corresponding item in
the other list box to be selected.  The exception it the EOS (Node 256) in
the weights list box which has no counter part.  
The Print Model check box will give a hard copy of the binary tree, scaled 
weights, and Huffman codes.  This choice is the best to 
get an overall look.The messages are there to give you an idea of what is 
going on with the program.  
The CharacterCounts subroutine and the CompressFile routine take so long, 
that I later added a progress bar to the program.  
The % Compressed will hold a number indicating the compression.  In the 
event that you get a negative number, the program is telling you that
it grew by that percentage.
The ending file lengths are displayed also in their respective frames.

EXPANSION PROGRAM.
This program is simpler to write.  It reads in the counts, builds a tree and
reads in the codes(symbols) and substitutes the ASCII character in the output
file.
You can type into the In Path and name or select from the file menu 
(Open Compressed File).
The output file path and name is selected by typing into the Out Path text
box or from the file menu (Open Expanded File).
There are messages as before and a progress bar when processing 
symbols(Codes).
The file lengths will be displayed in their respective frames.

