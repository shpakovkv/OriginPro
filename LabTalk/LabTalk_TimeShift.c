/*
TimeShift script adds specified TimeShift value to all time columns on all layers of specified Book(s) 

How to use:
1) Specify the Book(s) name(s).
2) Specify the TimeShift value.
5) Paste the script to the Command Window and press Enter.

NOTE
Column Types List:
Y     = 1
None  = 2
Y-Err = 3
X     = 4
Label = 5
Z     = 6
X-Err = 7

*/
StringArray BookNames = {"Book1"};

double TimeShift = 50.0;

/*
FOR FUTURE USE

dataset ColValueShift = {};

*/

//--------------------------------------------------------
int XColType = 4;
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	type "Select WorksheetPage[$(iBook)] '%(page.name$)'";
	for (iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		page.active = iLayer;
		int ColNumber = wks.nCols;

		type "Select LAYER $(iLayer) '%(wks.name$)'";
		if (TimeShift != 0.0)
		{
			for (iCol = 1; iCol <= ColNumber; iCol++)
			{
				wks.col = iCol;
				int rowNumber = wks.col.nRows;
				if (wks.col.type == XColType)
				{
					type "Process Column $(iCol) '%(wks.col.name$)' with $(rowNumber) rows";
					for (iRow = 1; iRow <= rowNumber; iRow++)
					{
						Cell(iRow, iCol) = Cell(iRow, iCol) + TimeShift;
					}
				}
			}
		}
	}
}
