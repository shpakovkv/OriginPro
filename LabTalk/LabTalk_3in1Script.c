/*
The 3 in 1 script:
	- replaces the name of all columns on all the layers of specifies Book(s) 
	  with corresponding string in the ColumnNames array.
	- inverts all the values of specified column on all the Layers of specified Book(s).
	- swaps two rows on all the Layers of the specified Book(s).
You can enable or disable any of this steps.

How to use:
1) Specify the Book(s) name(s).
2) Enable the FillColumnNames option and specify new names for all columns (sequentially from first to last). 
   If you don't want to change some column name(s) enter an empty string ("") in that position.
3) Enable the SwapRows option and Specify the indexes of two rows to swap (NOTE: first row index is 1).
4) Enable the InvertColumnValues option and specify the column index (NOTE: first column index is 1)
5) Paste the script to the Command Window and press Enter.
*/
StringArray BookNames;
BookNames.Add("PeaksSeries1");
BookNames.Add("PeaksSeries2");
BookNames.Add("PeaksSeries3");
BookNames.Add("PeaksSeries4");
BookNames.Add("PeaksSeries5");
BookNames.Add("PeaksSeries6");
BookNames.Add("PeaksSeries7");

// replace column names
int FillColumnNames = 1;
StringArray ColumnNames = {"", "PeakX", "PeakY", "sqrL", "sqrR", "sqrSum"};
 
int InvertColumnValues = 1;
int ColumnToInvert = 3;

int SwapRows = 0;
int FirstRowToSwap = 4;
int SecondRowToSwap = 5;

//--------------------------------------------------------
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		page.active = iLayer;
		int ColNumber = wks.nCols;

		if (FillColumnNames > 0)
		{
			for (iCol = 1; iCol <= ColNumber; iCol++)
			{
				if (ColumnNames.GetAt(iCol)$ != "")
				{
					wks.col = iCol;
					wks.col.name$ = ColumnNames.GetAt(iCol)$;
				}
			}
		}
			
		if (InvertColumnValues > 0)
		{
			wks.col = ColumnToInvert;
			int rowNumber = wks.col.nRows;
			for (iRow = 1; iRow <= rowNumber; iRow++)
			{
				Cell(iRow, ColumnToInvert) = -Cell(iRow, ColumnToInvert);
			}
		}

		if (SwapRows > 0)
		{
			for (iCol = 1; iCol <= ColNumber; iCol++)
			{
				double first = Cell(FirstRowToSwap, iCol);
				double second = Cell(SecondRowToSwap, iCol);
				Cell(FirstRowToSwap, iCol) = second;
				Cell(SecondRowToSwap, iCol) = first;
			}
		}
	}
}