/*
The script replaces the name of all columns on all the layers of specifies Book(s) 
with corresponding string in the ColumnNames array.

How to use:
1) Specify Book(s) name(s) in BookNames string array.
2) Specify new names for all columns (sequentially from first to last). 
   If you don't want to change some column name(s) enter an empty string ("") in that position.
3) Paste the script to the Command Window and press Enter.
*/

// add the short names of the target books
StringArray BookNames;
BookNames.Add("Book1");

// enter the new names for columns (one by one, from first to last)
// enter "" for columns you don't want to rename
StringArray ColumnNames = {"", "PeakX", "", "sqrL", "sqrR"};

//--------------------------------------------------------
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		page.active = iLayer;
		int ColNumber = wks.nCols;

		for (iCol = 1; iCol <= ColNumber && iCol <= ColumnNames.GetSize(); iCol++)
		{
			if (ColumnNames.GetAt(iCol)$ != "")
			{
				wks.col = iCol;
				wks.col.name$ = ColumnNames.GetAt(iCol)$;
			}
		}
	}
}
