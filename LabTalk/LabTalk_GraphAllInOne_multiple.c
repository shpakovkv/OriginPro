string graph_template$ = "TmpAllInOne";
int start_sheet = 1;  // one-based index of the first Workbook's sheet to be processed
int count = 105;  // amount of Workbook's sheets to be processed
dataset skip_list ={}; // list of one-based indexes of layers to be skipped
StringArray col_names={"U", "I", "rotate"};
StringArray book_names={"OutXRay", "OutXRay", "rotate"};

StringArray new_graph_names = {"AllInPMTD2","AllInPMTD4","AllInPMTD3","AllInPMTD1","AllInPMTD6","AllInPMTD8","AllInPMTD7","AllInPMTD5","AllInPMTD11","AllInPMTD12","AllInPMTD10","AllInPMTD9","AllInPMTDS1","AllInPMTDS2"};
string new_graph_name_postfix$ = "_v2";
StringArray rotate_cols = {"D2Pb3","D4Pb10","D3Pb30","D1Fe40","D6Pb3","D8Pb10","D7Pb30","D5Fe40","D11Pb50","D12Pb50","D10","D9","DS1","DS2Pb50"};
StringArray rotate_books = {"OutAxial","OutAxial","OutAxial","OutAxial","OutLatheral","OutLatheral","OutLatheral","OutLatheral","OutHama","OutHama","OutHama","OutHama","OutXRay","OutXRay"};
int rotate_on_graph_layer = 3;
for (idx=1; idx<=rotate_cols.GetSize(); idx++)
{
	col_names.SetAt(rotate_on_graph_layer, %(rotate_cols.GetAt(idx)$));
	book_names.SetAt(rotate_on_graph_layer, %(rotate_books.GetAt(idx)$));
	string name$ = new_graph_names.GetAt(idx)$;
	string postfix$ = new_graph_name_postfix$;
	GraphAllInOne(graph_template$, name$, postfix$, start_sheet, count, skip_list, col_names, book_names);
}
