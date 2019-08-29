/*
The script will delete the first X rows and the last Y rows on all layers of the specified workbooks.
Where X = deleteFirst value
Y = all last rows starting from deleteLastFrom


How to use:
1) Specify the Book(s) name(s).
2) Specify deleteFirst and deleteLastFrom values.
3) Paste the script to the Command Window and press Enter.
*/

StringArray BookNames = {"Book1", "Book2"};

int deleteFirst = 450000;
int deleteLastFrom = 702100;


//--------------------------------------------------------
int RowsCount = 0;

for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{	
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer <= LayersCount; iLayer++)
	{
		page.active = iLayer;
		RowsCount = wks.nRows;
		wks.deleteRows(deleteLastFrom, RowsCount - deleteLastFrom + 1);
		wks.deleteRows(1, deleteFirst);
	}
}
