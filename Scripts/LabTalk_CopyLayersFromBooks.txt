// Discription:
// Script copies layers from books with names in given string array to the book with specified name ('copy_to$')
//
// NOTE: The maximum number of WorksheetPage's (book's) layers is 255.

string copy_to$ = "Out2";	// book to copy layers to
// window -a %(copy_to$);
// layer -dk;				// use this to delete active layer of active book (this won't delete book)

StringArray copy_from;       // books to copy layers from
copy_from.Add("Dec26Out2");
copy_from.Add("Dec27Out2");

dataset start={1, 1};	// start copy from specified layer
dataset step={1, 1};		// step for layers loop (number of layers belonging to one shot)

StringArray postfix;
postfix.Add("");
postfix.Add("");


//==================================================================
window -a %(copy_to$);				// activates acceptor Book
int current_layer = page.nLayers;

for (book_idx=1; book_idx<=copy_from.GetSize(); book_idx++)  
{
    window -a %(copy_from.GetAt(book_idx)$);	// activates source Book
    int layers_count = page.nLayers;		
    for (shoot_idx=start[book_idx]; shoot_idx<=layers_count; shoot_idx+=step[book_idx])
    {
    	current_layer = current_layer + 1;
    	window -a %(copy_from.GetAt(book_idx)$);	// activates source Book
		page.active = shoot_idx;					// activates source Data worksheet

		string old_layer_name$;
		old_layer_name$ = wks.name$;
		string new_layer_name$;
		// new_layer_name$ = old_layer_name$ + "b" + Format(book_idx + 1, "#2")$;	// modify new layer name
		new_layer_name$ = old_layer_name$ + postfix.GetAt(book_idx)$;
		// new_layer_name$ = old_layer_name$;	// copy layer name

		StringArray ColNames;
		for (cc=1; cc<=wks.nCols; cc++)
		{
			wks.col = cc;						// activates source data column
			ColNames.Add(wks.col.name$);
		}
		window -a %(copy_to$);					// activates acceptor Book
		newsheet name:=new_layer_name$;

		// LabTalk adress formula: 
		// [WorkBookName]SheetNameOrIndex!ColumnNameOrIndex[RowIndex] 
		// Note: all LabTalk indexes are one-based
		type "Copying sheet %(copy_from.GetAt(book_idx)$)[$(shoot_idx)] to %(copy_to$)[$(current_layer)]";
		wrcopy iw:=[%(copy_from.GetAt(book_idx)$)]$(shoot_idx)! ow:=[%(copy_to$)]$(current_layer)! label:=LUC;
		// LUC <==> copy LongName, Units, Comments
		
		page.active = current_layer;					// activates acceptor worksheet
		for (cc=1; cc<=wks.nCols; cc++)
		{
			wks.col = cc;
			wks.col.name$ = ColNames.GetAt(cc)$;
		}
	}
}
type "Done.";