//=============================================
// ---  SETUP  --------------------------------

string graph_template$ = "TmpAllInOne";
StringArray col_names={"U", "I", "rotate"};
StringArray book_names={"OutXRay", "OutXRay", "rotate"};

StringArray new_graph_names = {"AllInPMTD2","AllInPMTD4","AllInPMTD3","AllInPMTD1","AllInPMTD6","AllInPMTD8","AllInPMTD7","AllInPMTD5","AllInPMTD11","AllInPMTD12","AllInPMTD10","AllInPMTD9","AllInPMTDS1","AllInPMTDS2"};
string new_graph_name_postfix$ = "_v2";
StringArray rotate_cols = {"D2Pb3","D4Pb10","D3Pb30","D1Fe40","D6Pb3","D8Pb10","D7Pb30","D5Fe40","D11Pb50","D12Pb50","D10","D9","DS1","DS2Pb50"};
StringArray rotate_books = {"OutAxial","OutAxial","OutAxial","OutAxial","OutLatheral","OutLatheral","OutLatheral","OutLatheral","OutHama","OutHama","OutHama","OutHama","OutXRay","OutXRay"};

int rotate_on_graph_layer = 3;

// specify intervals to include 
// (pairs of numbers where number is 1-baset index of Book sheet)
// for example {1,3,10,12} means include numbers 1,2,3,10,11,12   
dataset add_intervals = {1, 10};
// you may leave intervals empty and specify single numbers in add_numbers dataset

// specify intervals to exclude from graph 
//(pairs of numbers where number is 1-baset index of Book sheet)
dataset skip_intervals = {5, 8};

// secify single numbers to include
dataset add_numbers = {6};

// secify single numbers to exclude
dataset skip_numbers = {2, 3};

/*=================================================================
-------------   INTERVALS EXAMPLE   -------------------------------

dataset add_intervals = {1, 10};
Means we have sheets to process: 1,2,3,4,5,6,7,8,9,10  

dataset skip_intervals = {5, 8};
Now we cut previous interval, and get: 1,2,3,4,9,10

dataset add_numbers = {6};
Now we add single numbers, and get: 1,2,3,4,6,9,10 

dataset skip_numbers = {2, 3};
Now we remove single numbers, and get final list: 1,4,6,9,10

So there will be only curves form sheets number 1,4,6,9,10 on the graph. 
*/

//================================================
// ---  PROCESS  ---------------------------------
dataset skip_list = {0};  // need initial value

// int here is actually uint32_t 
int start_sheet = 4294967295; 
int last_sheet = 0;

//checkout paired values
// empty dataset's size is 1 with NANUM as first element
if ((skip_intervals[1] != NANUM) && (int(skip_intervals.GetSize()/2)*2 != skip_intervals.GetSize()))
{
	type "Error! Numbers in skip_intervals are not paired!";
	throw 100;
}
if ((add_intervals[1] != NANUM) && (int(add_intervals.GetSize()/2)*2 != add_intervals.GetSize()))
{
	type "Error! Numbers in add_intervals are not paired!";
	throw 100;
}
if ((add_intervals[1] == NANUM) && (add_numbers[1] == NANUM))
{
	type "Error! No numbers to include!";
	type "Add one or more intervals to the add_intervals dataset OR/AND single values to the add_numbers dataset.";
	throw 100;
}

// find first index to include
if (add_intervals.GetSize() > 1){
	type "Setting start/stop according intervals...";
	start_sheet = add_intervals[1];
	last_sheet = add_intervals[add_intervals.GetSize()];
}
else type "No intervals to include.";

// add_numbers.GetSize() is always >= 1
// because empty dataset's size is 1 with NANUM as first element 
if ((add_numbers.GetSize() > 1) || (add_numbers[1] != NANUM))
{
	if ((add_numbers[1] > 0) && (start_sheet > add_numbers[1])){
		start_sheet = add_numbers[1];
		type "Changing start according add_numbers[1] to $(start_sheet)";
	}
	if (last_sheet < add_numbers[add_numbers.GetSize()]){
		last_sheet = add_numbers[add_numbers.GetSize()];
		type "Changing last number according add_numbers[last] to $(last_sheet)";
	}
}
else type "No single numbers to add.";

// add skip numbers from between add_intervals
for (int interval = 2; interval <= int(add_intervals.GetSize() / 2); interval++)
{
	// from previous interval second number + 1
	// to current interval first number - 1
	for (int jj=add_intervals[(interval - 1) * 2] + 1; jj < add_intervals[interval * 2 - 1]; jj++)
	{
		// if current number is not in add_list, then add in to skip_list
		if (List(jj, add_numbers) == 0) 
		{
			skip_list[skip_list.GetSize() + 1] = jj;
		}
	}
}

// when there is no intervals added
// means we have skip interval from first to last number
if (add_intervals.GetSize() == 1)
{
	skip_intervals[1] = start_sheet;
	skip_intervals[2] = last_sheet;
}

// add skip numbers from skip_intervals
for (interval = 1; interval <= int(skip_intervals.GetSize() / 2); interval++)
{
	for (int jj=skip_intervals[interval * 2 - 1]; jj <= skip_intervals[interval * 2]; jj++)
	{
		// if not in add_numbers list AND not included before
		if ((List(jj, add_numbers) == 0) && (List(jj, skip_list) == 0))
		{
			skip_list[skip_list.GetSize() + 1] = jj;
		}
	}
}

// add skip numbers from skip_numbers
for (int jj = 1; jj <= skip_numbers.GetSize(); jj++)
{
	// if not included before
	if (List(skip_numbers[jj], skip_list) == 0)
	{
		skip_list[skip_list.GetSize() + 1] = skip_numbers[jj];
	}
}

//=========================================================
// Numbers and sheets names control

// Get book names without repeats
type "Getting unique names list...";
StringArray unames = {};
for (int idx = 1; idx <= book_names.GetSize(); idx++)
{
	type "Loop started.";
	if (unames.Find(book_names.GetAt(idx)$) == 0)
	{
		type "Adding new to unames...";
		unames.Append(book_names.GetAt(idx)$);
	}
}
type "Got unique names list.";

// // print start and stop sheet numbers and all numbers to skip
// type "=================================================";
// type "Start = $(start_sheet)   Stop = $(last_sheet)";
// for (int ii=1; ii<=skip_list.GetSize(); ii++)
// {
// 	type "Skip $(skip_list[ii])";
// }

type "=================================================";
type "Start = $(start_sheet)   Stop = $(last_sheet)";
type "Numbers included:";
string line$ = "";
for (int ii=start_sheet; ii<=last_sheet; ii++)
{
	if (List(ii, skip_list) == 0)
	{
		line$ = "Include # $(ii) == ";
		for (int jj = 1; jj <= unames.GetSize(); jj++)
		{
			range rlayer = [%(unames.GetAt(jj)$)]$(ii)!;
			line$ += "(%(unames.GetAt(jj)$))[%(rlayer.name$)] ";
			if ("%(rlayer.name$)" == "")
			{
				type "ERROR! Not enough layer in WorksheetPage '%(unames.GetAt(jj)$)'";
				throw 100;
			}
		}
		type "%(line$)";
	}
}

type "-------------------------------------------------";
type "The end of the checkout. Graph construction has begun.";


for (idx=1; idx<=rotate_cols.GetSize(); idx++)
{
	col_names.SetAt(rotate_on_graph_layer, %(rotate_cols.GetAt(idx)$));
	book_names.SetAt(rotate_on_graph_layer, %(rotate_books.GetAt(idx)$));
	string name$ = new_graph_names.GetAt(idx)$;
	string postfix$ = new_graph_name_postfix$;
	GraphAllInOne(graph_template$, name$, postfix$, start_sheet, (last_sheet - start_sheet + 1), skip_list, col_names, book_names);
}
