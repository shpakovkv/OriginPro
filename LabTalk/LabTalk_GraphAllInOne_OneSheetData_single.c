/*
Script for AllInOne graph

Input data structure:
- One WorksheetBook for each curve
- All WorksheetBooks has one layer
- Layer's columns filled with the same data type pattern (for example: [X], [Y], [X], [Y], etc.)
- Data type pattern is the same for all WorksheetBooks
*/

//=============================================
// ---  SETUP  --------------------------------

string graph_template$ = "Tmp01";
string new_graph_name$ = "TestGraph";
string postfix$ = "";

StringArray book_names={"Book1", "Book2"};

// Number of different data-column types for each curve
// (for example: for [X], [Y], [X], [Y], etc. pattern step = 2)
// (for example: for [X], [Y], [Z], [Err], [X], [Y], [Z], [Err], etc. pattern step = 4)
step = 2;

// specify intervals to include 
// (pairs of numbers where number is 1-baset index of Book sheet)
// for example {1,3,10,12} means include numbers 1,2,3,10,11,12
dataset add_intervals = {1, 10, 20, 25};
// you may leave intervals empty and specify single numbers in add_numbers dataset

// specify intervals to exclude from graph 
//(pairs of numbers where number is 1-baset index of Book sheet)
dataset skip_intervals = {4, 5, 11, 12, 22, 22};

// secify single numbers to include
dataset add_numbers = {15, 17};

// secify single numbers to exclude
dataset skip_numbers = {24};

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
int start_idx = 4294967295; 
int last_idx = 0;

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
	start_idx = add_intervals[1];
	last_idx = add_intervals[add_intervals.GetSize()];
}
else type "No intervals to include.";

// add_numbers.GetSize() is always >= 1
// because empty dataset's size is 1 with NANUM as first element 
if ((add_numbers.GetSize() > 1) || (add_numbers[1] != NANUM))
{
	if ((add_numbers[1] > 0) && (start_idx > add_numbers[1])){
		start_idx = add_numbers[1];
		type "Changing start according add_numbers[1] to $(start_idx)";
	}
	if (last_idx < add_numbers[add_numbers.GetSize()]){
		last_idx = add_numbers[add_numbers.GetSize()];
		type "Changing last number according add_numbers[last] to $(last_idx)";
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
	skip_intervals[1] = start_idx;
	skip_intervals[2] = last_idx;
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
// type "Start = $(start_idx)   Stop = $(last_idx)";
// for (int ii=1; ii<=skip_list.GetSize(); ii++)
// {
// 	type "Skip $(skip_list[ii])";
// }

type "=================================================";
type "Start = $(start_idx)   Stop = $(last_idx)";
type "Numbers included:";
string line$ = "";
for (int jj = 1; jj <= unames.GetSize(); jj++)
{
	
	line$ = "Include # $(ii) == ";
	for (int ii=start_idx; ii<=last_idx; ii++)
	{
		if (List(ii, skip_list) == 0)
		{
			range r_col = [%(unames.GetAt(jj)$)]0!$(ii);
			line$ += "(%(unames.GetAt(jj)$))(0)[%(r_col.name$)] ";
			if ("%(r_col.name$)" == "")
			{
				last_idx = ii - 1;
				type "Boundaries updated:";
				type "Start = $(start_idx)   Stop = $(last_idx)";
			}
		}
	}
	type "%(line$)";
}

type "-------------------------------------------------";
type "The end of the checkout. Graph construction has begun.";

GraphAllInOne_OneSheet(graph_template$, new_graph_name$, postfix$, start_idx, (last_idx - start_idx + 1), skip_list, book_names, step);