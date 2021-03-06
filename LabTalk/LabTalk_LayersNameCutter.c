/*
The script extracts N characters from a Layer name, starting from StartIdx index, 
then replaces the old layer name with the new string.
The script repeats the operation for all Layers of all specified Book(s).

How to use:
1) Specify the Book(s) name(s).
2) Specify the index of the first character from which to start extracting substring.
3) Specify the number of characters to extract.
4) Paste the script to the Command Window and press Enter.
*/
//=================================================
// SETUP

// add the short names of the target books
StringArray bookNames = {"Book1", "Book2"};  // the list of the WorksheetPages to be processed

// the index of the first letter to be left (previous letters will be deleted)
// NOTE: the index of the first letter is 1
int StartIdx = 5;

// the number of characters to be left (all the others will be deleted)
int CharCount = 11;

//---------------------------------------------------
for (iBook = 1; iBook<=bookNames.GetSize(); iBook++)
{
	window -a %(bookNames.GetAt(iBook)$);
	int count = page.nLayers;	
	for (iLayer = 1; iLayer<=count; iLayer++)
	{
		page.active = iLayer;
		string NewName$ = layer.name$;
		NewName$ = NewName.Mid(StartIdx, CharCount)$;
		layer.name$ = NewName$;
	}
}