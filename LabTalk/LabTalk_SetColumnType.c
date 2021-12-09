/*
The script changes the type of each column according to pattern. 
Process all layers of specifies Book(s).
You can specify any column type to set.

How to use:
1) Specify Book(s) name(s) in BookNames string array.
2) Specify one-based index of the first column to change.
3) Specify the type number to set (from 1 to 7, see comments).
4) Paste the script to the Command Window and press Enter.



TYPES LIST:
Y     = 1
None  = 2
Y-Err = 3
X     = 4
Lebel = 5
Z     = 6
X-Err = 7
*/


//--------------------------------------------------------
// add the short names of the target books
StringArray BookNames;
BookNames.Add("Book1");

// specify column type pattern
dataset typePattern = {4, 1};

//--------------------------------------------------------
int pattLen = typePattern.GetSize();
// type "pattLen = $(pattLen)";
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	// type "layers count = $(LayersCount)";
	for (iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		page.active = iLayer;
		int ColNumber = wks.nCols;
		// type "page = $(iLayer)";
		// type "ColNumber = $(ColNumber)";


		for (iCol = 1; iCol <= ColNumber; iCol++)
		{
			int iType = Mod(iCol, pattLen);
			// type "idx = $(iCol), type = $(iType)";
			wks.col = iCol;
			wks.col.type = typePattern[iType];
		}
	}
}
