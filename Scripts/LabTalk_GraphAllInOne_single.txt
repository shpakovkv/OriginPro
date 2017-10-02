string graph_template$ = "TmpAllInOne";
string new_graph_name$ = "TestGraph";
string postfix$ = "_v1";
int start_sheet = 1;  // one-based index of the first Workbook's sheet to be processed
int count = 105;  // amount of Workbook's sheets to be processed
dataset skip_list ={}; // list of one-based indexes of layers to be skipped
StringArray col_names={"U", "I", "D1Fe40"};
StringArray book_names={"OutXRay", "OutXRay", "OutAxial"};

GraphAllInOne(graph_template$, new_graph_name$, postfix$, start_sheet, count, skip_list, col_names, book_names);