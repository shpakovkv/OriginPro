/*
Discription:
Save all layers of selected WorksheetPage(s) as .csv files
to the specified folder

How to use:
1) Choose WorksheetPage with layers to be exported
2) Get ShortName of that WorksheetPage
3) Assign that name to the variable "book_name"
4) Create a folder for exported files
5) Specify the path to that folder via variable save_path$
7) Copy below script and paste it into the Origin's Command Window
8) Press Enter
*/

StringArray book_name;
book_name.Add("OutDPO7054");
book_name.Add("OutLeCroy");
book_name.Add("OutTDS3054BL");
book_name.Add("OutTDS3054BR");
book_name.Add("OutTDS2024C");

string save_path$ = "C:\temp\temp";

//===========================================================
path_len = save_path.GetLength();
if (save_path.GetAt(path_len) != 92)
{
	save_path.Insert(path_len + 1, '\');
}

for (bb = 1; bb<=book_name.GetSize(); bb++)
{
	window -a %(book_name.GetAt(bb)$);
	int layers_count = page.nLayers;
	for (ii=1; ii<=layers_count; ii++)
	{
		page.active = ii;
		string save_as$ = save_path$ + book_name.GetAt(bb)$ + "_" + wks.name$ + ".csv";
		//save_as$=;
		expASC type:=csv path:=save_as$ csvsep:=Semicolon longname:=0 units:=0 comment:=0;
	}
}