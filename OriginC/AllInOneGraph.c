/*------------------------------------------------------------------------------*
 * File Name:				 													*
 * Creation: 																	*
 * Purpose: OriginC Source C file												*
 * Copyright (c) ABCD Corp.	2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010		*
 * All Rights Reserved															*
 * 																				*
 * Modification Log:															*
 *------------------------------------------------------------------------------*/
 
////////////////////////////////////////////////////////////////////////////////////
// Including the system header file Origin.h should be sufficient for most Origin
// applications and is recommended. Origin.h includes many of the most common system
// header files and is automatically pre-compiled when Origin runs the first time.
// Programs including Origin.h subsequently compile much more quickly as long as
// the size and number of other included header files is minimized. All NAG header
// files are now included in Origin.h and no longer need be separately included.
//
// Right-click on the line below and select 'Open "Origin.h"' to open the Origin.h
// system header file.
#include <Origin.h>
////////////////////////////////////////////////////////////////////////////////////

//#pragma labtalk(0) // to disable OC functions for LT calling.

////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.

#define DEBUG 0

// FUNCTIONS:
/*
void FillAllInOneGraphLayer(GraphLayer Graph, string DataBookName, string ColName, int CurvesCount, int StartLayer, vector<int> skip)
void GraphAllInOne(string TmpGrName, string NewGrName, string postfix, int StartLayer, int CurvesCount, vector<int> skip, StringArray ColNames, StringArray BookNames)
*/
///////////////////////////////////////////////////////

void AssertEqualBookLayersCount(int GraphLayersCount,
                                StringArray BookNamesForLayer1, 
                                StringArray BookNamesForLayer2, 
                                StringArray BookNamesForLayer3, 
                                StringArray BookNamesForLayer4, 
                                StringArray BookNamesForLayer5, 
                                StringArray BookNamesForLayer6, 
                                StringArray BookNamesForLayer7, 
                                StringArray BookNamesForLayer8, 
                                StringArray BookNamesForLayer9, 
                                StringArray BookNamesForLayer10)
{
#if (DEBUG == 1)
	out_str("DEBUG: inside AssertEqualBookLayersCount; GraphLayersCount = " + GraphLayersCount + ".");
#endif
	int booksNumber = BookNamesForLayer1.GetSize();
	int shotsNumber = GetLayersSum(BookNamesForLayer1);
#if (DEBUG == 1)
	out_str("DEBUG: GetLayersSum(BookNamesForLayer1) = " + shotsNumber + ".");
#endif
	if (GraphLayersCount > 1) Assert(shotsNumber == GetLayersSum(BookNamesForLayer2), "Error! The number of layers in books from books list #2 is different!");
	if (GraphLayersCount > 2) Assert(shotsNumber == GetLayersSum(BookNamesForLayer3), "Error! The number of layers in books from books list #3 is different!");
	if (GraphLayersCount > 3) Assert(shotsNumber == GetLayersSum(BookNamesForLayer4), "Error! The number of layers in books from books list #4 is different!");
	if (GraphLayersCount > 4) Assert(shotsNumber == GetLayersSum(BookNamesForLayer5), "Error! The number of layers in books from books list #5 is different!");
	if (GraphLayersCount > 5) Assert(shotsNumber == GetLayersSum(BookNamesForLayer6), "Error! The number of layers in books from books list #6 is different!");
	if (GraphLayersCount > 6) Assert(shotsNumber == GetLayersSum(BookNamesForLayer7), "Error! The number of layers in books from books list #7 is different!");
	if (GraphLayersCount > 7) Assert(shotsNumber == GetLayersSum(BookNamesForLayer8), "Error! The number of layers in books from books list #8 is different!");
	if (GraphLayersCount > 8) Assert(shotsNumber == GetLayersSum(BookNamesForLayer9), "Error! The number of layers in books from books list #9 is different!");
	if (GraphLayersCount > 9) Assert(shotsNumber == GetLayersSum(BookNamesForLayer10), "Error! The number of layers in books from books list #10 is different!");
	

}

int GetLayersSum(StringArray BookNames)
{
#if (DEBUG == 1)
	string msg = "DEBUG: inside GetLayersSum; BookNames: [";
	for (int idx=0; idx < BookNames.GetSize(); idx++)
	{
		msg = msg + BookNames[idx] + ", "
	}
	msg = msg + "]";
	out_str(msg);
#endif
	int booksNumber = BookNames.GetSize();
	int layersNumber = 0;
	for (int i=0; i < booksNumber; i++)
	{
		WorksheetPage BookNamesWP(BookNames[i]);
		if (!BookNamesWP)
		{
			out_str("Error!\n Can't find WorksheetPage '" + BookNames[i] + "'!\n");
			throw 100;
		}
		layersNumber += BookNamesWP.Layers.Count();
		BookNamesWP.Detach();
		BookNamesWP.Destroy();
	}
	return layersNumber;
}

void Assert(bool statment, string message)
{
#if (DEBUG == 1)
	string res = "False";
	if (statment) res = "True";
	out_str("DEBUG: inside Assert;  statment == " + res + "; message: '" + message + "'.");
#endif
	if (!statment)
	{
		out_str(message);
		throw 100;
	}
}

void printIntVectorArg(vector<int> arr, string name)
{
	string msg = "vector<int> " + name + ": [";
	for(int n = 0 ; n < arr.GetSize(); n++)
	{
		msg = msg + arr[n] + ", ";
	}
	msg = msg + "]";
	out_str(msg);
}

void printStringArrayArg(StringArray arr, string name)
{
	string msg = "<StringArray> " + name + ": [";
	for(int n = 0 ; n < arr.GetSize(); n++)
	{
		msg = msg + arr[n] + ", ";
	}
	msg = msg + "]";
	out_str(msg);
}


void FillAllInOneGraphLayer(GraphLayer Graph, string DataBookName, string ColName, int CurvesCount, int StartLayer, vector<int> skip)
{//filling single GraphLayer with Labels, Curves & Legend.
//GraphNumber >= 1
	
	// SKIP - one-based indexes of DATA-Book layer with data to be skiped in AllInOne graph process
	vector<uint> temp_vec;
	
	//for all ColumnNames
	for (int i=StartLayer; i<CurvesCount+StartLayer; i++)
	{//plotting curves
		if (skip.Find(temp_vec, i + 1) == 0)	// skip contains one-based indexes, 'i' - zero-based index
		{
			Worksheet TempWks(Project.WorksheetPages(DataBookName).Layers(i));
			
			//data access check. Exit plotting if fail.
			if (!TempWks.Columns(ColName)) 
			{
				out_str("Error!\nData access error!");
			}
			else 
			{
				//TempDS.Attach(TempWks.Columns(VariableNames[i]));
				Curve cr;
				cr.Attach(TempWks.Columns(ColName));
				Graph.AddPlot(cr);
			}
			TempWks.Detach();
		}
	}
	//Grouping plots
	Graph.GroupPlots(0, CurvesCount-1);
}


void GraphAllInOne(string TmpGrName, string NewGrName, string postfix, int StartLayer, int CurvesCount, vector<int> skip, StringArray ColNames, StringArray BookNames)
{
	// GraphAllInOne(tmp_graph$, name$, postfix$, start_layer, count, skip_list, col_names, book_names);
	
	int GraphLayersCount = ColNames.GetSize();

	GraphPage TemplateGraphPage;
	TemplateGraphPage = Project.GraphPages(TmpGrName);
	Tree trTemplateFormat;
	trTemplateFormat = TemplateGraphPage.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	StartLayer = StartLayer - 1;	// convert one-based index to zero-based
	if (StartLayer < 0) { StartLayer = 0; }

	GraphPage NewGraphPage; //create new GraphPage window
	NewGraphPage.Create("origin");
	
	NewGraphPage.SetName(NewGrName + postfix);
	NewGraphPage.Show = 0;  //hiding new graph window
	
	//...and add Postfix to it for LongName
	NewGraphPage.SetLongName(NewGrName + postfix);
	
	//loading Layers from TemplategraphPage
	if (GraphLayersCount > 0)
	{//add enough Layers with Template format to NewGraphPage	
		//NewGraphPage.Layers(0).Destroy();
		
		NewGraphPage.AddLayers(TemplateGraphPage);
		for (int j=1; j<=GraphLayersCount; j++) 
		{//for all GraphLayers in TemplateGraphPage
			//NewGraphPage.AddLayer(TemplateGraphPage.Layers(j-1));
			
			if (NewGraphPage.Layers(j).DataPlots.Count() > 0)
			{//deleting excess Curves from new Layer if needed
				for (int p = NewGraphPage.Layers(j).DataPlots.Count(); p > 0; p--)
				{
					NewGraphPage.Layers(j).RemovePlot(0);
				}
			}
		}
	}
	//deleting excess (first one) layer
	NewGraphPage.Layers(0).Delete();
	
	//filling Layers of NewGraphPage with Curves
	for (int j=0; j<GraphLayersCount; j++)
	{//for all GraphLayers in NewGraphPage
		FillAllInOneGraphLayer(NewGraphPage.Layers(j), BookNames[j], ColNames[j], CurvesCount, StartLayer, skip);
	}
	
	bool FormatUpdErr = NewGraphPage.ApplyFormat( trTemplateFormat, true, true, true ); //changing format
	if (!FormatUpdErr) {out_str("Warning!\n GraphPage ('"+NewGraphPage.GetName()+"') format update failed!\n");}
	
	NewGraphPage.Show = 0; //hiding new graph window
	NewGraphPage.Detach();
}

//=============================================================================================================================
//-----     MULT-BOOK VERSION     ---------------------------------------------------------------------------------------------
//=============================================================================================================================

void FillAllInOneGraphLayer_MultiBooks(GraphLayer Graph, StringArray BookNames, string ColName, int ShotsCount, int StartShoot, vector<int> skip)
{//filling single GraphLayer with Labels, Curves & Legend.
//GraphNumber >= 1
	/*
	TODO:
	1. BookNames ==> array
	2. end-to-end numbering
	3. delete StartShoot
	*/
	// SKIP - one-based indexes of DATA-Book layer with data to be skiped in AllInOne graph process
	vector<uint> temp_vec;

	int bookIdx = 0;
	int wksIdx = StartShoot;
	int prevBooksWksCount = 0;

	// when StartShoot is not in the first book
	for (int i=0; i < BookNames.GetSize(); i++)
	{
		int currentBookWksCnt = Project.WorksheetPages(BookNames[bookIdx]).Layers.Count();
		if (wksIdx - prevBooksWksCount > currentBookWksCnt)
		{
			wksIdx = wksIdx - currentBookWksCnt;
			prevBooksWksCount = prevBooksWksCount + currentBookWksCnt;
			bookIdx++;
		}
		else break;
	}
	
	for (int shotIdx = StartShoot; shotIdx < ShotsCount + StartShoot; shotIdx++)
	{
		// when current shot is in the next book
		if (shotIdx - prevBooksWksCount >= Project.WorksheetPages(BookNames[bookIdx]).Layers.Count())
		{
			prevBooksWksCount = prevBooksWksCount + Project.WorksheetPages(BookNames[bookIdx]).Layers.Count();
			bookIdx++;
			wksIdx = 0;
		}

		// if current shot is not in skip list
		if (skip.Find(temp_vec, shotIdx + 1) == 0)	// skip contains one-based indexes, 'shotIdx' - zero-based index
		{
			//out_str("Book '" + BookNames[bookIdx] + "';  wksIdx = " + wksIdx + ";  shotIdx = " + shotIdx + ";  prevBooksWksCount = " + prevBooksWksCount);
			Worksheet TempWks(Project.WorksheetPages(BookNames[bookIdx]).Layers(wksIdx));
			//out_str("Worksheet '" + TempWks.GetName() + "' columnt '" + ColName + "'");
			
			//data access check. Exit plotting if fail. 
			if (!TempWks.Columns(ColName)) 
			{
				out_str("Error!\nCan't find column with data! WorksheetPage '" + BookNames[bookIdx] + "', Worksheet #" + wksIdx + ", Column '" + ColName + "'");
			}
			else 
			// add new single curve to group
			{
				Curve cr;
				cr.Attach(TempWks.Columns(ColName));
				Graph.AddPlot(cr);
			}
			TempWks.Detach();
		}

		wksIdx++;
	}
	//Grouping plots
	Graph.GroupPlots(0, ShotsCount-1);
}

//================================================================================================================================

// void GraphAllInOne(string TmpGrName, string NewGrName, string postfix, int StartLayer, int CurvesCount, vector<int> skip, StringArray ColNames, StringArray BookNames)
// void GraphAllInOne_MultiBooks(graph_template$, new_graph_name$, postfix$, start_sheet, (last_sheet - start_sheet + 1), skip_list, col_names, books_for_column_1, books_for_column_2, books_for_column_3, books_for_column_4, books_for_column_5, books_for_column_6, books_for_column_7, books_for_column_8, books_for_column_9, books_for_column_10)

//void GraphAllInOne_MultiBooks(string TmpGrName, string NewGrName, string postfix, int StartLayer, int AllCurvesCount, vector<int> skip, StringArray ColNames, StringArray BNames1)

//void GraphMult(string TmpGrName, string NewGrName, string postfix, int StartLayer, int AllCurvesCount, vector<int> skip, StringArray ColNames, StringArray BNames1, StringArray BNames2)

void GraphAllInOne_MultiBooks(string TmpGrName, string NewGrName, string postfix, int StartLayer, int AllCurvesCount, vector<int> skip, StringArray ColNames, StringArray BNames1, StringArray BNames2, StringArray BNames3, StringArray BNames4, StringArray BNames5, StringArray BNames6, StringArray BNames7, StringArray BNames8, StringArray BNames9, StringArray BNames10)
{
#if (DEBUG == 1)
	out_str("DEBUG: inside GraphAllInOne_MultiBooks; ");
	out_str("TmpGrName = " + TmpGrName);
	out_str("NewGrName = " + NewGrName);
	out_str("postfix = " + postfix);
	out_str("StartLayer = " + StartLayer);
	out_str("AllCurvesCount = " + AllCurvesCount);
	printIntVectorArg(skip, "skip");
	printStringArrayArg(ColNames, "ColNames");
	printStringArrayArg(BNames1, "BNames1");
	printStringArrayArg(BNames2, "BNames2");
	printStringArrayArg(BNames3, "BNames3");
	printStringArrayArg(BNames4, "BNames4");
	printStringArrayArg(BNames5, "BNames5");
	printStringArrayArg(BNames6, "BNames6");
	printStringArrayArg(BNames7, "BNames7");
	printStringArrayArg(BNames8, "BNames8");
	printStringArrayArg(BNames9, "BNames9");
	printStringArrayArg(BNames10, "BNames10");
#endif
	
	int MAX_LAYERS = 10;
	// GraphAllInOne(tmp_graph$, name$, postfix$, start_layer, count, skip_list, col_names, book_names);
	
	int GraphLayersCount = ColNames.GetSize();

	if(GraphLayersCount > MAX_LAYERS)
	{
		out_str("Error!\n Column names count  (" + GraphLayersCount + ") is more than " + MAX_LAYERS + "!\n");
		throw 100; // Force error
	}

	AssertEqualBookLayersCount(GraphLayersCount, BNames1, BNames2, BNames3, BNames4, BNames5, BNames6, BNames7, BNames8, BNames9, BNames10);

	// get template format
	GraphPage TemplateGraphPage;
	TemplateGraphPage = Project.GraphPages(TmpGrName);
	if(TemplateGraphPage.Layers.Count() > MAX_LAYERS)
	{
		out_str("Error!\n The number of layers (" + TemplateGraphPage.Layers.Count() + ") in template graph page is more than " + MAX_LAYERS + "!\n");
		throw 100; // Force error	
	}
	Tree trTemplateFormat;
	trTemplateFormat = TemplateGraphPage.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	StartLayer = StartLayer - 1;	// convert one-based index to zero-based
	if (StartLayer < 0) { StartLayer = 0; }

	GraphPage NewGraphPage; //create new GraphPage window
	NewGraphPage.Create("origin");
	
	NewGraphPage.SetName(NewGrName + postfix);
	NewGraphPage.Show = 0;  //hiding new graph window
	
	//...and add Postfix to it for LongName
	NewGraphPage.SetLongName(NewGrName + postfix);
	
	//loading Layers from TemplategraphPage
	if (GraphLayersCount > 0)
	{
		//copy layers
		NewGraphPage.AddLayers(TemplateGraphPage);

		for (int j=1; j<=GraphLayersCount; j++) 
		{//for all GraphLayers in NewGraphPage
		 //deleting excess Curves from new Layer if needed
			
			if (NewGraphPage.Layers(j).DataPlots.Count() > 0)
			{
				for (int p = NewGraphPage.Layers(j).DataPlots.Count(); p > 0; p--)
				{
					NewGraphPage.Layers(j).RemovePlot(0);
				}
			}
		}
	}
	//deleting excess (first one) layer
	NewGraphPage.Layers(0).Delete();

	// Sort skip-list
	skip.Sort(SORT_ASCENDING);

	//filling Layers of NewGraphPage with Curves
	for (int j=0; j<GraphLayersCount; j++)
	{//for all GraphLayers in NewGraphPage
		switch(j)
		{
			case 0:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames1, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;

			case 1:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames2, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;

			case 2:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames3, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;

			case 3:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames4, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;

			case 4:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames5, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;

			case 5:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames6, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;

			case 6:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames7, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;

			case 7:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames8, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;

			case 8:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames9, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;

			case 9:
				FillAllInOneGraphLayer_MultiBooks(NewGraphPage.Layers(j), BNames10, ColNames[j], AllCurvesCount, StartLayer, skip);
				break;
		}
	}
	
	bool FormatUpdErr = NewGraphPage.ApplyFormat( trTemplateFormat, true, true, true ); //changing format
	if (!FormatUpdErr) {out_str("Warning!\n GraphPage ('"+NewGraphPage.GetName()+"') format update failed!\n");}
	
	NewGraphPage.Show = 0; //hiding new graph window
	NewGraphPage.Detach();
}

/*

//------------------      LAB_TALK SCRIPT EXAMPLE         -----------------------

//==================================================
//-----     ONE GRAPH     --------------------------
//==================================================
string tmp_graph$ = "TmpAllInOne";
string new_graph$ = "TestGraph"; // check if exists
string postfix$ = "_v1";
int start_layer = 1;
int count = 105;
dataset skip_list ={}; // one-based index of layers to be skipped
StringArray col_names={"U", "I", "D1Fe40"};
StringArray book_names={"OutXRay", "OutXRay", "OutAxial"};

GraphAllInOne(tmp_graph$, new_graph$, postfix$, start_layer, count, skip_list, col_names, book_names);


//==================================================
//-----     MULTIPLE GRAPHS     --------------------
//==================================================
string tmp_graph$ = "TmpAllInOne";

int start_layer = 1;
int count = 105;
dataset skip_list ={}; // one-based index of layers to be skipped
StringArray col_names={"U", "I", "rotate"};
StringArray book_names={"OutXRay", "OutXRay", "rotate"};

StringArray new_graph_names = {"AllInPMTD2","AllInPMTD4","AllInPMTD3","AllInPMTD1","AllInPMTD6","AllInPMTD8","AllInPMTD7","AllInPMTD5","AllInPMTD11","AllInPMTD12","AllInPMTD10","AllInPMTD9","AllInPMTDS1","AllInPMTDS2"};
string new_graph_name_postfix$ = "_v2";
StringArray rotate_cols = {"D2Pb3","D4Pb10","D3Pb30","D1Fe40","D6Pb3","D8Pb10","D7Pb30","D5Fe40","D11Pb50","D12Pb50","D10","D9","DS1","DS2Pb50"};
StringArray rotate_books = {"OutAxial","OutAxial","OutAxial","OutAxial","OutLatheral","OutLatheral","OutLatheral","OutLatheral","OutHama","OutHama","OutHama","OutHama","OutXRay","OutXRay"};
int rotate_graph_layer = 3;
for (idx=1; idx<=rotate_cols.GetSize(); idx++)
{
	col_names.SetAt(rotate_graph_layer, %(rotate_cols.GetAt(idx)$));
	book_names.SetAt(rotate_graph_layer, %(rotate_books.GetAt(idx)$));
	string name$ = new_graph_names.GetAt(idx)$;
	string postfix$ = new_graph_name_postfix$;
	GraphAllInOne(tmp_graph$, name$, postfix$, start_layer, count, skip_list, col_names, book_names);
}

*/


//=================================================================================================
//-----     SINGLE-SHEET BOOKs     ----------------------------------------------------------------
//=================================================================================================


void FillAllInOneGraphLayer_FromOneSheet(GraphLayer Graph, string DataBookName, int CurvesCount, int StartIdx, vector<int> skip, int step)
{//filling single GraphLayer with Labels, Curves & Legend.
//GraphNumber >= 1
	
	// SKIP - one-based indexes of DATA-Book layer with data to be skiped in AllInOne graph process
	vector<uint> temp_vec;
	
	//use only first layer
	Worksheet TempWks(Project.WorksheetPages(DataBookName).Layers(0));
	for (int i=StartIdx; i<CurvesCount+StartIdx; i = i + step)
	{//StartIdx is zero-based index
		if (skip.Find(temp_vec, i + 1) == 0)	// skip contains one-based indexes, 'i' - zero-based index
		{
			//data access check. Exit plotting if fail.
			out_str("Number of cols == " + TempWks.GetNumCols());
			while (i < TempWks.GetNumCols() && TempWks.Columns(i).GetType() != 0)
			{
				i++;
			}
			if (TempWks.Columns(i).GetType() == 0) 
			{
				//TempDS.Attach(TempWks.Columns(VariableNames[i]));
				Curve cr;
				cr.Attach(TempWks.Columns(i));
				Graph.AddPlot(cr);
			}
		}
	}
	TempWks.Detach();
	//Grouping plots
	Graph.GroupPlots(0, CurvesCount-1);
}






void GraphAllInOne_OneSheet(string TmpGrName, string NewGrName, string postfix, int StartIdx, int CurvesCount, vector<int> skip, StringArray BookNames, int step)
{
	// GraphAllInOne(tmp_graph$, name$, postfix$, start_layer, count, skip_list, col_names, book_names);
	
	int GraphLayersCount = BookNames.GetSize();

	GraphPage TemplateGraphPage;
	TemplateGraphPage = Project.GraphPages(TmpGrName);
	Tree trTemplateFormat;
	trTemplateFormat = TemplateGraphPage.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	StartIdx = StartIdx - 1;	// convert one-based index to zero-based
	if (StartIdx < 0) { StartIdx = 0; } // convert one-based index to zero-based

	GraphPage NewGraphPage; //create new GraphPage window
	NewGraphPage.Create("origin");
	
	NewGraphPage.SetName(NewGrName + postfix);
	NewGraphPage.Show = 0;  //hiding new graph window
	
	//...and add Postfix to it for LongName
	NewGraphPage.SetLongName(NewGrName + postfix);
	
	//loading Layers from TemplategraphPage
	if (GraphLayersCount > 0)
	{//add enough Layers with Template format to NewGraphPage	
		//NewGraphPage.Layers(0).Destroy();
		
		NewGraphPage.AddLayers(TemplateGraphPage);
		for (int j=1; j<=GraphLayersCount; j++) 
		{//for all GraphLayers in TemplateGraphPage
			//NewGraphPage.AddLayer(TemplateGraphPage.Layers(j-1));
			
			if (NewGraphPage.Layers(j).DataPlots.Count() > 0)
			{//deleting excess Curves from new Layer if needed
				for (int p = NewGraphPage.Layers(j).DataPlots.Count(); p > 0; p--)
				{
					NewGraphPage.Layers(j).RemovePlot(0);
				}
			}
		}
	}
	//deleting excess (first one) layer
	NewGraphPage.Layers(0).Delete();
	
	//filling Layers of NewGraphPage with Curves
	for (int j=0; j<GraphLayersCount; j++)
	{//for all GraphLayers in NewGraphPage
		FillAllInOneGraphLayer_FromOneSheet(NewGraphPage.Layers(j), BookNames[j], CurvesCount, StartIdx, skip, step);
	}
	
	bool FormatUpdErr = NewGraphPage.ApplyFormat( trTemplateFormat, true, true, true ); //changing format
	if (!FormatUpdErr) {out_str("Warning!\n GraphPage ('"+NewGraphPage.GetName()+"') format update failed!\n");}
	
	NewGraphPage.Show = 0; //hiding new graph window
	NewGraphPage.Detach();
}