/*
The script prints the names of all columns located on the active layer of the active workbook.

How to use:
1) Select the target layer of the target WorkBook.
2) Paste the script to the Command Window and press Enter.
*/

type "";
string output$ = "{";
for (int iCol = 1; iCol <= wks.nCols; iCol++)
{

	wks.col = iCol;
	type %(wks.col.name$);
	output$ += ""%(wks.col.name$)", ";
}
output$ += "}";

type "";
type output$;
