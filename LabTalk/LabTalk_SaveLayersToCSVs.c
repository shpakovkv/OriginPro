/*
Discription:
Save all layers of selected WorksheetPage(s) as .csv files
to the specified folder

How to use:
1) Choose WorksheetPage with layers to be exported
2) Get ShortName of that WorksheetPage
3) Assign that name to the variable "bookNames"
4) Create a folder for exported files
5) Specify the path to that folder via variable saveToPath$
7) Copy below script and paste it into the Origin's Command Window
8) Press Enter
*/

//==========================================================
// SETUP
StringArray bookNames;
bookNames.Add("OutDPO7054");
bookNames.Add("OutLeCroy");
bookNames.Add("OutTDS3054BL");
bookNames.Add("OutTDS3054BR");
bookNames.Add("OutTDS2024C");

string saveToPath$ = "C:\temp\temp";

//===========================================================
// PROCESS 
path_len = saveToPath.GetLength();
if (saveToPath.GetAt(path_len) != 92)
{
	saveToPath.Insert(path_len + 1, '\');
}

for (bb = 1; bb<=bookNames.GetSize(); bb++)
{
	window -a %(bookNames.GetAt(bb)$);
	int layers_count = page.nLayers;
	for (ii=1; ii<=layers_count; ii++)
	{
		page.active = ii;
		string save_as$ = saveToPath$ + bookNames.GetAt(bb)$ + "_" + wks.name$ + ".csv";
		//save_as$=;
		expASC type:=csv path:=save_as$ csvsep:=Semicolon longname:=0 units:=0 comment:=0;
	}
}