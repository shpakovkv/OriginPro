/*
The script inverts all the values of specified column on all the Layers of specified Book(s).

How to use:
1) Specify the Book(s) name(s).
2) Specify the column index (NOTE: first column index is 1)
3) Paste the script to the Command Window and press Enter.
*/

// add the short names of the target book(s)
StringArray BookNames;
BookNames.Add("PeaksSeries1");
BookNames.Add("PeaksSeries2");
BookNames.Add("PeaksSeries3");
BookNames.Add("PeaksSeries4");

// the index of the column to invert values
int ColumnToInvert = 3;

//--------------------------------------------------------
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		page.active = iLayer;
		int ColNumber = wks.nCols;

		wks.col = ColumnToInvert;
		int rowNumber = wks.col.nRows;
		for (iRow = 1; iRow <= rowNumber; iRow++)
		{
			Cell(iRow, ColumnToInvert) = -Cell(iRow, ColumnToInvert);
		}
	}
}
