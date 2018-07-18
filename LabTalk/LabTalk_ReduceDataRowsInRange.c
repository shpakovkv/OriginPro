/*
The script reduces the number of row:
the script will leave first row, then delete ReduceFactor rows, leave next row, again delete ReduceFactor rows, etc.
The script will process only the rows with indexes between Start and Stop.
The script will process all columns of all layers of the specified WorksheetPages (WorkBooks).

How to use:
1) Specify the Book(s) name(s).
2) Specify the ReduceFactor parameter.
3) Specify the Start and the Stop indexes.
5) Paste the script to the Command Window and press Enter.
*/

StringArray BookNames;
BookNames.Add("OutTDS3054BR");
BookNames.Add("OutTDS3054BL");
BookNames.Add("OutLeCroy");
BookNames.Add("OutDPO7054");

int ReduceFactor = 1;
int Start = 185;
int Stop = 309;

//--------------------------------------------------------
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{	
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer <= LayersCount; iLayer++)
	{
		page.active = iLayer;
		// int RowCount = wks.nRows;
		// RowCount = RowCount / ReduceFactor;
		// if (mod(wks.nRows, ReduceFactor) > 0) {
		// 	RowCount++;
		// }
		int ColCount = wks.nCols;
		for (iCol = 1; iCol <= ColCount; iCol++){
			wks.col = iCol;
			int RowCount = wks.col.nRows;
			if (RowCount > Stop) {
				RowCount = Stop;
			}
			for (iRow = RowCount; iRow > Start + ReduceFactor; iRow = iRow - ReduceFactor - 1) {
				range RowsToDel = $(iCol)[$(iRow - ReduceFactor + 1):$(iRow)];
				del RowsToDel;
			}
		}
		// wks.nRows = RowCount;
	}
}
