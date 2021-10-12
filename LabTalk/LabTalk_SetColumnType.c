/*
The script changes the type of each column from 'startCol' through 'step'. 
Process all layers of specifies Book(s).
You can specify any column type to set.

How to use:
1) Specify Book(s) name(s) in BookNames string array.
2) Specify one-based index of the first column to change.
3) Specify the type number to set (from 1 to 7, see comments).
4) Paste the script to the Command Window and press Enter.
*/

//--------------------------------------------------------
// add the short names of the target books
StringArray BookNames;
BookNames.Add("Book1");

// specify starting column one-based index
int startCol = 1;

// specify step (index increment)
int step = 2;

// specify type to set
int typeToSet = 4;

/* TYPES LIST:
Y     = 1
None  = 2
Y-Err = 3
X     = 4
Lebel = 5
Z     = 6
X-Err = 7
*/

//--------------------------------------------------------
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		page.active = iLayer;
		int ColNumber = wks.nCols;

		for (iCol = startCol; iCol <= ColNumber; iCol=iCol+step)
		{
			wks.col = iCol;
			wks.col.type = typeToSet;
		}
	}
}
