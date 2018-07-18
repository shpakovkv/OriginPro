/*
The script deletes first N1 rows with X < minX, and last N2 rows with X > maxX. 
The script runs on all layers of the specified WorksheetPages (WorkBooks).

How to use:
1) Specify the Book(s) name(s).
2) Specify the minX and maxX values (double).
3) Choose X column to check MIN values: specify column index (the first column index is 1).
4) Choose X column to check MAX values: specify column index (the first column index is 1).
5) Paste the script to the Command Window and press Enter.
*/
StringArray BookNames;
BookNames.Add("Book1");

double minX = 20.0;
int minXcolumn = 1;
double maxX = 50.0;
int maxXcolumn = 3;


//--------------------------------------------------------
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{	
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer <= LayersCount; iLayer++)
	{
		page.active = iLayer;

		// get number of first rows to delete
		int firstRow = 1;
		int lastRow = wks.col$(minXcolumn).nRows;
		double temp;
		for (iRow = 1; iRow <= lastRow; iRow++) {
			if (minX < Cell(iRow, minXcolumn)) 
				{
					firstRow = iRow;
					break;
				}
		}
		for (iRow = lastRow; iRow > firstRow; iRow--) {
			if (Cell(iRow, maxXcolumn) < maxX) {
				lastRow = iRow;
				break;
			}
		}
		int lastRowsCnt = wks.nRows - lastRow;
		wks.nRows = lastRow;

		// delete last rows
		type "Deleting first $(firstRow - 1) & last $(lastRowsCnt) rows at layer[$(iLayer)] of %(BookNames.GetAt(iBook)$).";

		// delete first rows
		int ColCount = wks.nCols;
		for (iCol = 1; iCol <= ColCount; iCol++){
			if (firstRow > 1) {
				// type DeletingCOLUMN$(iCol)_ROWS_from_1_to_$(firstRow);
				range rowsToDel = $(iCol)[1:$(firstRow - 1)];
				del rowsToDel;
			}
		}
	}
}
