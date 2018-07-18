/*
The script deletes every N row from all worksheets (layers) of specified WorksheetPages (WorkBooks).

How to use:
1) Specify the Book(s) name(s).
2) Specify the delEveryNRow parameter.
5) Paste the script to the Command Window and press Enter.
*/

StringArray BookNames;
BookNames.Add("OutTDS3054BR");
BookNames.Add("OutTDS3054BL");
BookNames.Add("OutTDS2024C");
BookNames.Add("OutLeCroy");
BookNames.Add("OutDPO7054");

int delEveryNRow = 2;


//--------------------------------------------------------
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{	
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer <= LayersCount; iLayer++)
	{
		page.active = iLayer;
		int RowCount = wks.nRows;
		RowCount = RowCount / delEveryNRow;
		if (mod(wks.nRows, delEveryNRow) > 0) {
			RowCount++;
		}
		int ColCount = wks.nCols;
		for (iCol = 1; iCol <= ColCount; iCol++){
			reducerows irng:=col($(iCol)) npts:=$(delEveryNRow) method:=first rd:=col($(iCol));
		}
		wks.nRows = RowCount;
	}
}
