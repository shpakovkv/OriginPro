
//=============================================
// ---  SETUP  --------------------------------
del -op;
del -as;

string graph_template$ = "Template1";
string new_graph_name$ = "TestGraph";
string postfix$ = "_v1";

StringArray col_names = {"Voltage", "ShuntA", "ShuntB", "PMTS2"};
StringArray book_names_for_col_1 = {"OSC1ch1gr1", "OSC1ch1gr2", "OSC1ch1gr3"};
StringArray book_names_for_col_2 = {"OSC1ch2gr1", "OSC1ch2gr2", "OSC1ch2gr3"};
StringArray book_names_for_col_3 = {"OSC1ch3gr1", "OSC1ch3gr2", "OSC1ch3gr3"};
StringArray book_names_for_col_4 = {"OSC1ch4gr1", "OSC1ch4gr2", "OSC1ch4gr3"};
StringArray book_names_for_col_5 = {};
StringArray book_names_for_col_6 = {};
StringArray book_names_for_col_7 = {};
StringArray book_names_for_col_8 = {};
StringArray book_names_for_col_9 = {};
StringArray book_names_for_col_10 = {};

// specify intervals to include 
// (pairs of numbers where number is 1-baset index of Book sheet)
// for example {1,3,10,12} means include numbers 1,2,3,10,11,12
dataset add_intervals = {2, 8};
// you may leave intervals empty and specify single numbers in add_numbers dataset

// specify intervals to exclude from graph 
//(pairs of numbers where number is 1-baset index of Book sheet)
dataset skip_intervals = {4, 5};

// secify single numbers to include
dataset add_numbers = {1, 4, 5, 9};

// secify single numbers to exclude
dataset skip_numbers = {7};

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

//===========================================================================
//-------   FUNCTIONS   -----------------------------------------------------

Function int appendUniqueBookNameScript(ref StringArray setOfNames, string name, int alreadyAdded)
{	int addedNames = 0;
	if (setOfNames.Find(name$) == 0)
		{
			setOfNames.Append(name$);
			addedNames++;
		}
		else 
			if (setOfNames.Find(name$) <= alreadyAdded)
			{
				type "Error! Book %(name$) was added on one of the previous turns!";
				throw 100;
			}
	return addedNames;
}

function int printStringArrayScript(StringArray arr, string name$)
{
	for (int idx=1; idx <= arr.GetSize(); idx++)
	{
		type "%(name$)[$(idx)] == %(arr.GetAt(idx)$)";
	}
	type "---";
	return 0;
}

//===========================================================================
// ---  PROCESS  ------------------------------------------------------------

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

StringArray unames = {};
int addedOnPrevTurns = 0;
int GraphLayersCnt = col_names.GetSize();
type "GraphLayersCnt = $(GraphLayersCnt)";

for (int idx = 1; idx <= book_names_for_col_1.GetSize(); idx++)
{
	int addedOnThisTurn = 0;
	type "Group of books #$(idx)...";

	if (GraphLayersCnt > 0) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_1.GetAt(idx)$, addedOnPrevTurns);
	if (GraphLayersCnt > 1) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_2.GetAt(idx)$, addedOnPrevTurns);
	if (GraphLayersCnt > 2) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_3.GetAt(idx)$, addedOnPrevTurns);
	if (GraphLayersCnt > 3) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_4.GetAt(idx)$, addedOnPrevTurns);
	if (GraphLayersCnt > 4) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_5.GetAt(idx)$, addedOnPrevTurns);
	if (GraphLayersCnt > 5) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_6.GetAt(idx)$, addedOnPrevTurns);
	if (GraphLayersCnt > 6) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_7.GetAt(idx)$, addedOnPrevTurns);
	if (GraphLayersCnt > 7) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_8.GetAt(idx)$, addedOnPrevTurns);
	if (GraphLayersCnt > 8) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_9.GetAt(idx)$, addedOnPrevTurns);
	if (GraphLayersCnt > 9) addedOnThisTurn = addedOnThisTurn + appendUniqueBookNameScript(unames, book_names_for_col_10.GetAt(idx)$, addedOnPrevTurns);

	addedOnPrevTurns = addedOnPrevTurns + addedOnThisTurn;
}
type " ";
type "-------------------------------------------------";
type "-------  GraphAllInOne_MultiBooks args:  --------";
type " "
type "graph_template = %(graph_template$)";
type "new_graph_name = %(new_graph_name$)";
type "postfix = %(postfix$)";
type "start_sheet = $(start_sheet)";
type "shots_count = $(last_sheet - start_sheet + 1)";
type "---";
printStringArrayScript(skip_list, "skip_list");
printStringArrayScript(col_names, "col_names");
printStringArrayScript(book_names_for_col_1, "book_names_for_col_1");
printStringArrayScript(book_names_for_col_2, "book_names_for_col_2");
printStringArrayScript(book_names_for_col_3, "book_names_for_col_3");
printStringArrayScript(book_names_for_col_4, "book_names_for_col_4");
printStringArrayScript(book_names_for_col_5, "book_names_for_col_5");
printStringArrayScript(book_names_for_col_6, "book_names_for_col_6");
printStringArrayScript(book_names_for_col_7, "book_names_for_col_7");
printStringArrayScript(book_names_for_col_8, "book_names_for_col_8");
printStringArrayScript(book_names_for_col_9, "book_names_for_col_9");
printStringArrayScript(book_names_for_col_10, "book_names_for_col_10");
type "-------------------------------------------------";
type "The end of the checkout. GraphAllInOne_MultiBooks() has been called.";
type " ";

int numberOfSheets = (last_sheet - start_sheet + 1);

GraphAllInOne_MultiBooks(graph_template$, new_graph_name$, postfix$, start_sheet, numberOfSheets, skip_list, col_names, book_names_for_col_1, book_names_for_col_2, book_names_for_col_3, book_names_for_col_4, book_names_for_col_5, book_names_for_col_6, book_names_for_col_7, book_names_for_col_8, book_names_for_col_9, book_names_for_col_10);
