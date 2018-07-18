/*
The script reduces the number of row in all the worksheets (layers) of the specified WorksheetPages (WorkBooks).
The script will average every ReduceFactor numbers in all columns and output the average of each group to thees columns.
The numder of rows will be reduced ReduceFactor times.

The script outputs the final time step of col(1) of layer(1) for each WorkBook in BookNames.

How to use:
1) Specify the Book(s) name(s).
2) Specify the ReduceFactor parameter.
5) Paste the script to the Command Window and press Enter.
*/

StringArray BookNames;
BookNames.Add("OutTDS3054BR");
BookNames.Add("OutTDS3054BL");
BookNames.Add("OutLeCroy");
BookNames.Add("OutDPO7054");

int ReduceFactor = 4;


//--------------------------------------------------------
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{	
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer <= LayersCount; iLayer++)
	{
		page.active = iLayer;
		int RowCount = wks.nRows;
		RowCount = RowCount / ReduceFactor;
		if (mod(wks.nRows, ReduceFactor) > 0) {
			RowCount++;
		}
		int ColCount = wks.nCols;
		for (iCol = 1; iCol <= ColCount; iCol++){
			reducerows irng:=col($(iCol)) npts:=$(ReduceFactor) method:=ave rd:=col($(iCol));
		}
		wks.nRows = RowCount;
	}
}
//---------------------------------------------------------
// Print time step for all worksheetpages
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{	
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	page.active = 1;
	double dt = Cell(2, 1) - Cell(1, 1);
	type "%(BookNames.GetAt(iBook)$), layer[1], X1 column new time step is $(dt)";
}
