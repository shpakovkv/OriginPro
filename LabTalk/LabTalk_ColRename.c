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
StringArray ColumnNames = {"Time1", "Coil_HMO", "Time2", "CamExp", "Time3", "Voltage", "Time4", "Coil_TDS", "Time5", "Open1", "Time6", "Open2"};
StringArray ColumnLongNames = {"Time", "Coil-HMO", "Time", "Camera Exposition", "Time", "Voltage", "Time", "Coil-TDS", "Time", "None", "Time", "None"};
StringArray ColumnUnits = {"ns", "a.u.", "ns", "a.u.", "ns", "MV", "ns", "a.u.", "ns", "a.u.", "ns", "a.u."};


//--------------------------------------------------------
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	type "Processing WorkSheetBook %(BookNames.GetAt(iBook)$)";
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;

	for (iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		page.active = iLayer;
		int ColNumber = wks.nCols;

		int notReady = 0;
		if (ColumnNames.GetSize() > ColNumber)
		{
			type "WARNING! Book %(BookNames.GetAt(iBook)$) Layer $(iLayer) has less layers than ColumnNames values";
		}
		if (ColumnLongNames.GetSize() > ColNumber)
		{
			type "WARNING! Book %(BookNames.GetAt(iBook)$) Layer $(iLayer) has less layers than ColumnLongNames values";
		}
		if (ColumnUnits.GetSize() > ColNumber)
		{
			type "WARNING! Book %(BookNames.GetAt(iBook)$) Layer $(iLayer) has less layers than ColumnUnits values";
		}

		if (ColumnNames.GetSize() < ColNumber) notReady++;
		if (ColumnLongNames.GetSize() < ColNumber) notReady++;
		if (ColumnUnits.GetSize() < ColNumber) notReady++;

		if (notReady > 2)
		{
			type "Error! Not enough input StringArray values!";
			throw 100;
		}

		for (iCol = 1; iCol <= ColNumber && iCol <= ColumnNames.GetSize(); iCol++)
		{
			if (ColumnNames.GetAt(iCol)$ != "")
			{
				wks.col = iCol;

				if (ColumnNames.GetSize() >= ColNumber)
				{
					wks.col.name$ = ColumnNames.GetAt(iCol)$;
				}

				if (ColumnLongNames.GetSize() >= ColNumber)
				{
					wks.col.label$ = ColumnLongNames.GetAt(iCol)$;
				}

				if (ColumnUnits.GetSize() >= ColNumber)
				{
					wks.col.unit$ = ColumnUnits.GetAt(iCol)$;
				}
			}
		}
	}
}
