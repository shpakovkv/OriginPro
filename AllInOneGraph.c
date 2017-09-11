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
// FUNCTIONS:
/*
void FillAllInOneGraphLayer(GraphLayer Graph, string DataBookName, string ColName, int CurvesCount, int StartLayer, vector<int> skip)
void GraphAllInOne(string TemplateGraphPageName, string NewGraphName, string postfix, int StartLayer, int CurvesCount, vector<int> skip, StringArray ColNames, StringArray BookNames)
*/
///////////////////////////////////////////////////////


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


void GraphAllInOne(string TemplateGraphPageName, string NewGraphName, string postfix, int StartLayer, int CurvesCount, vector<int> skip, StringArray ColNames, StringArray BookNames)
{
	// GraphAllInOne(tmp_graph$, name$, postfix$, start_layer, count, skip_list, col_names, book_names);
	
	int GraphLayersCount = ColNames.GetSize();

	GraphPage TemplateGraphPage;
	TemplateGraphPage = Project.GraphPages(TemplateGraphPageName);
	Tree trTemplateFormat;
	trTemplateFormat = TemplateGraphPage.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	StartLayer = StartLayer - 1;	// convert one-based index to zero-based
	if (StartLayer < 0) { StartLayer = 0; }

	GraphPage NewGraphPage; //create new GraphPage window
	NewGraphPage.Create("origin");
	
	NewGraphPage.SetName(NewGraphName + postfix);
	NewGraphPage.Show = 0;  //hiding new graph window
	
	//...and add Postfix to it for LongName
	NewGraphPage.SetLongName(NewGraphName + postfix);
	
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