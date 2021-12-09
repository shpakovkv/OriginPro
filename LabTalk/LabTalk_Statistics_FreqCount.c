/*
The script calculates frequency count for the specified column on all the Layers of specified Book(s).
Creates new WorkBook for output data.

How to use:
1) Specify the input data Book(s) name(s).
2) Specify the output  Book(s) name(s).
3) Specify the column index (NOTE: first column index is 1)
4) Specify the start value, stop value and increment.
5) Paste the script to the Command Window and press Enter.
*/

// add the short names of the target books
StringArray BookNames = {"Book2"};

// add the short names of the output books to create
StringArray outBookName = {"Hist2"};

// output books postfix
string outBookPostfix$ = "_freq1";

// the one-based index of the column to calculate
int colNumber = 4;

double startValue = 0.1;
double stopValue = 9.1;
double incrementValue = 0.1;

//-------------------------------------------------------------
string layerName$;
for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	newbook name:=outBookName.GetAt(iBook)$ sheet:=0 option:=lsname;	// create new WorkBook with 0 Worksheets and with Short Name = ss$
	string newName$ = page.name$;

	window -a %(BookNames.GetAt(iBook)$);
	LayersCount = page.nLayers;
	for (iLayer = 1; iLayer<=LayersCount; iLayer++)
	{
		window -a %(BookNames.GetAt(iBook)$);
		page.active = iLayer;
		layerName$ = layer.name$;
		// type "[%(outBookName.GetAt(iBook)$)]$(iLayer)!col($(colNumber))";

		// int ColNumber = wks.nCols;
		// freqcounts irng:=col($(colNumber)) min:=startValue max:=stopValue stepby:=increment inc:=incrementValue bin:=center outleft:=0 outright:=0 cumulcount:=0 freq:=1 rd:=[newName$]<new>;
		
		freqcounts irng:=[%(BookNames.GetAt(iBook)$)]$(iLayer)!col($(colNumber)) min:=startValue max:=stopValue stepby:=increment inc:=incrementValue bin:=center outleft:=0 outright:=0 cumulcount:=0 freq:=1 rd:=[newName$]<new>;
		window -a %(outBookName.GetAt(iBook)$);
		page.active = iLayer;
		layer.name$ = "%(layerName$)%(outBookPostfix$)";
	}
}