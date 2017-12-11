/*
The script swaps two rows on all the Layers of the specified Book(s).

How to use:
1) Specify the Book(s) name(s).
2) Specify the indexes of two rows to swap (NOTE: first row index is 1).
3) Specify the cell type.
4) Paste the script to the Command Window and press Enter.
*/
StringArray BookNames;
BookNames.Add("PeaksSeries1");
BookNames.Add("PeaksSeries2");
BookNames.Add("PeaksSeries3");
BookNames.Add("PeaksSeries4");

int FirstRowToSwap = 4;
int SecondRowToSwap = 5;

// choose from "double", "int", "string" (default = "double")
string cellType = "double"; 

//--------------------------------------------------------
switch (cellType){
	case "int":
		int first;
		int second;
	case "string":
		string first$;
		string second$;
	default:
		double first;
		double second;
}
if
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		page.active = iLayer;
		int ColNumber = wks.nCols;

		for (iCol = 1; iCol <= ColNumber; iCol++)
		{	
			if (cellType != "string"){
				first = Cell(FirstRowToSwap, iCol);
				second = Cell(SecondRowToSwap, iCol);
				Cell(FirstRowToSwap, iCol) = second;
				Cell(SecondRowToSwap, iCol) = first;
			}
			else{
				first$ = Cell(FirstRowToSwap, iCol)$;
				second$ = Cell(SecondRowToSwap, iCol)$;
				Cell(FirstRowToSwap, iCol)$ = second$;
				Cell(SecondRowToSwap, iCol)$ = first$;
			}
		}
	}
}