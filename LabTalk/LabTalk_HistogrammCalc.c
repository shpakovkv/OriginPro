/*
The script calcs histogramm and write hist data to new workbook.
The script will processed all worksheet of the specified data workbook.

How to use:
1) Specify the Book(s) name(s).
2) Specify all parameters
3) Paste the script to the Command Window and press Enter.
*/

// add the short names of the target books
StringArray BookNameList = {"Film03"};

// save results to new workbook:
StringArray ResultsBookList = {"Hist03"};

// index of 1-based column to calc frequency counts
int ColNumber = 2;  

// bining parameters
double start_bin_from = 0;
double end_bin_to = 30;
double bin_step = 0.2;

// =========================================================================

String DataWksName$;

// check Rusult Book Name is not used
type "Check result books names are not used";
for (iBook = 1; iBook<=ResultsBookList.GetSize(); iBook++)
{
	document -cw %(ResultsBookList.GetAt(iBook)$)
	if (count > 0)  
	{
		type -b "WorksheetPage containing name '%(ResultsBookList.GetAt(iBook)$)' already exist!";
		throw 100;
	}
}

// page count match check
type "page count match check";
if (BookNameList.GetSize() != ResultsBookList.GetSize())
{
	type "The number of WorksheetPages in BookNameList and ResultsBookList must be the same!";
	throw 100;
}

// int ExtraLayersNum;
// create result workbooks
type "create result workbooks";
for (iBook = 1; iBook<=ResultsBookList.GetSize(); iBook++)
{
	newbook %(ResultsBookList.GetAt(iBook)$) option:=lsname;
	window -a %(ResultsBookList.GetAt(iBook)$);

	// ExtraLayersNum = page.nLayers;
}

type "main loop";
for (iBook = 1; iBook<=BookNameList.GetSize(); iBook++)
{
	window -a %(BookNameList.GetAt(iBook)$);
	count = page.nLayers;
	for (iLayer = 1; iLayer<=count; iLayer++)
	{
		window -a %(BookNameList.GetAt(iBook)$);
		page.active = iLayer;
		DataWksName$ = wks.name$;

		window -a %(ResultsBookList.GetAt(iBook)$);					// activates acceptor Book
		newsheet name:=DataWksName$;

		freqcounts irng:=[%(BookNameList.GetAt(iBook)$)]$(iLayer)!Col($(ColNumber)) bin:=end min:=start_bin_from max:=end_bin_to stepby:=increment inc:=bin_step freq:=1 rd:=[%(ResultsBookList.GetAt(iBook)$)]$(iLayer + 1)!;
		// freqcounts irng:=[Resultswater3]$(1)!Col(2) bin:=end min:=0.1 max:=30 stepby:=increment inc:=0.1 freq:=1 rd:=[Book2]$(2)!;

		// [%(copy_from.GetAt(book_idx)$)]$(shoot_idx)!
		// "%(unames.GetAt(jj)$))[%(rlayer.name$)]"
	}

	// delete first extra layer
	window -a %(ResultsBookList.GetAt(iBook)$);	
	layer -d 1;

}