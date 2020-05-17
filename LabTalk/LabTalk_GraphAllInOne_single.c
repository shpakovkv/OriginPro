
//=============================================
// ---  SETUP  --------------------------------

string graph_template$ = "TmpAllInOne";
string new_graph_name$ = "TestGraph";
string postfix$ = "_v1";

StringArray col_names={"B", "B"};
StringArray book_names={"Book1", "Book1"};

// specify intervals to include 
// (pairs of numbers where number is 1-baset index of Book sheet)
// for example {1,3,10,12} means include numbers 1,2,3,10,11,12   
dataset add_intervals = {1, 5, 10, 15, 20, 22};  

// specify intervals to exclude from graph 
//(pairs of numbers where number is 1-baset index of Book sheet)
dataset skip_intervals = {4, 5, 11, 12, 22, 22};

// secify single numbers to include
dataset add_numbers = {};

// secify single numbers to exclude
dataset skip_numbers = {3};

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
	int zz = 1/0;
}
if ((add_intervals[1] != NANUM) && (int(add_intervals.GetSize()/2)*2 != add_intervals.GetSize()))
{
	type "Error! Numbers in add_intervals are not paired!";
	int zz = 1/0;
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
// dataset unames = {};
// for (int ii = 1; ii <= book_names.GetSize(), ii++)
// {
// 	if (List(book_names[ii], unames) == 0)
// }

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
for (int ii=start_sheet; ii<=last_sheet; ii++)
{
	if (List(ii, skip_list) == 0)
	{
		type "Include $(ii)";	
	}
}

GraphAllInOne(graph_template$, new_graph_name$, postfix$, start_sheet, (last_sheet - start_sheet + 1), skip_list, col_names, book_names);
