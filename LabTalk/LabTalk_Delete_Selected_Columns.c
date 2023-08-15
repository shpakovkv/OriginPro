/*
The script deletes columns with the specified ShortNames on all layers of the specified WorkBook(s).

How to use:
1) Specify Book(s) name(s) in BookNames string array.
2) Specify the ShortName(s) of all columns to be deleted
3) Paste the script to the Command Window and press Enter.
*/

// specify the short names of the target workbook(s)
StringArray BookNames = {"Book1"};

// specify the short names of the columns to be deleted
StringArray ColNames = {"MyCol", "C", "temp"};


//--------------------------------------------------------
/*
exist(name)=
0  Object does not exist.  
1  dataset (column)
2  workbook/worksheet  
3  graph window  
4  numeric variable  
5  matrix book/matrix sheet  
7  tool  
8  macro  
9  notes window  
11  layout window  
12  Excel worksheet  

*/

for (int iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	type "Processing WorkSheetBook %(BookNames.GetAt(iBook)$)";
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;

	for (int iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		page.active = iLayer;
		int ColNumber = wks.nCols;

		for (int iCol = 1; iCol <= ColNames.GetSize(); iCol++)
		{
			// check if selected range exists and represents dataset (column) (exist == 1)
			if (exist("[%(BookNames.GetAt(iBook)$)]$(iLayer)!%(ColNames.GetAt(iCol)$)") == 1)
			{
				// type "DELETE [%(BookNames.GetAt(iBook)$)]$(iLayer)!%(ColNames.GetAt(iCol)$)";
				del col(%(ColNames.GetAt(iCol)$));
			}
			else
			{
				type "Not found [%(BookNames.GetAt(iBook)$)]$(iLayer)!%(ColNames.GetAt(iCol)$)";
			}
		}

	}
}
