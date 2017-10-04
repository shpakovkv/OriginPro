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


////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.

#include <graph.h>
////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.

void ShowMessage(string m)
{//message box and CommandWindow log
//CommandWindow log appears only when it active
	out_str(m);
	LT_execute("type -b " + m);
}

string StringPostfixIncrease(string s)
{	
	int n, l;
	string OutStr;
	
	l = 0;
	
	while ((NANUM != atof(s.Mid(s.GetLength()-l-1))) && (l < 15))
	{
		l++;
	}
	
	if (l==0) 	{	l=15;	n=2;	}
		else 	
		{	
			n = atoi(s.Mid(s.GetLength()-l));
			n++;
			l = s.GetLength()-l;
		}
	
	OutStr = s.Mid(0,l) + n;
	
	if ((OutStr.GetLength() > 17) && (l<15)) {	OutStr = s.Mid(0,l-1) + n;	}
	
	return OutStr;
}



string StringPrefixIncrease(string s)
{
	int NumLen;
	NumLen = 1;
	
	if (NANUM == atof(s.Mid(0, 1)))
	{
		return s;
	}
	else
	{
		while (NANUM != atof(s.Mid(0, NumLen)))	{	NumLen++;	}
		NumLen--;
		ShowMessage("NumLen = " + NumLen);
		int Num = atoi(s.Mid(0, NumLen));
		Num++;
		string out;
		out = Num;
		ShowMessage("StringLen = " + out.GetLength());
		while (NumLen > out.GetLength())
		{	
			out = "0" + out;	
			ShowMessage("CurrentString = " + out);
		}
		
		out = out + s.Right(s.GetLength()-NumLen);
		ShowMessage("String = " + out);
		return out;
	}
	
}

//---------------------------------------------------------------------------------------------------------------------------------
void FillGraphLayer(GraphLayer Graph, Worksheet OptionsWks, WorksheetPage DataBook, vector<int> OscCount, int GraphNumber, int GraphLayerNumber, StringArray DefaultOptionsLabel, StringArray DataBookNames, bool y_rescale)
{//filling single GraphLayer with Labels, Curves & Legend.
//GraphNumber >= 1
	
	Dataset TempDS, OscNumber;
	StringArray VariableNames;
	vector<int> DataBookNumbers(0);
	
	//getting data columns names list
	OptionsWks.Columns(DefaultOptionsLabel[1]).GetStringArray(VariableNames); 
	
	//getting access to 'OscNumber'
	OscNumber.Attach(OptionsWks.Columns(DefaultOptionsLabel[0]));
	
	TempDS.Attach(OptionsWks.Columns(DefaultOptionsLabel[3]));
	for (int i=0; i<=TempDS.GetUpperBound(); i++)
	{
		DataBookNumbers.Add(TempDS[i]);
	}
	
	double y_max = 0;
	double y_min = 0;
	double local_min, local_max;
	uint nIndexMin, nIndexMax, nCountNonMissingValues;
	//for all ColumnNames
	for (i=0; i<VariableNames.GetSize(); i++)
	{//plotting curves
		Worksheet TempWks(Project.WorksheetPages(DataBookNames[DataBookNumbers[i]-1]).Layers((GraphNumber-1)*OscCount[DataBookNumbers[i]-1]+OscNumber[i]-1));
		
		//data access check. Exit plotting if fail.
		if (!TempWks.Columns(VariableNames[i])) 
		{
			out_str("Error!\nData access error in Graph#"+GraphNumber+", Layer#" + GraphLayerNumber + ", Curve#" + (i+1) + "\nColumn '"+VariableNames[i]+"' at layer #"+((GraphNumber-1)*OscCount[DataBookNumbers[i]-1]+OscNumber[i])+" of '"+DataBookNames[DataBookNumbers[i]-1]+"' WorksheetPage is missing.\nCan't plot another curves in that layer.");
		}
		else 
		{
			Dataset y_ds;
			y_ds.Attach(TempWks.Columns(VariableNames[i]));
			vector y_vec(y_ds);
			nCountNonMissingValues = y_vec.GetMinMax(local_min, local_max, &nIndexMin, &nIndexMax);
			if (y_max < local_max) {y_max = local_max;}
			if (y_min > local_min) {y_min = local_min;}
			y_ds.Detach();

			//TempDS.Attach(TempWks.Columns(VariableNames[i]));
			Curve cr;
			cr.Attach(TempWks.Columns(VariableNames[i]));
			Graph.AddPlot(cr);
		}
		TempWks.Detach();
	}
	
	//Grouping plots
	TempDS.Attach(OptionsWks.Columns(DefaultOptionsLabel[2]));
	
	i=0;
	//for all plots exept last
	while (i<VariableNames.GetSize()-1)
	{
		//if present Group¹ == next Group¹ 
		if ((TempDS[i] != 0) && (TempDS[i] == TempDS[i+1]))
		{
			int StartIndex = i;
			//serching for last group index
			while ((i<VariableNames.GetSize()-1) && (TempDS[i] == TempDS[i+1])) {	i++;	}
			Graph.GroupPlots(StartIndex, i); //grouping
		}
		i++;
	}
	TempDS.Detach();
	
	if (y_rescale)
	{
		Graph.Rescale(OKAXISTYPE_Y);
	}

	OscNumber.Detach();
	VariableNames.RemoveAll();
	DataBookNumbers.RemoveAll();
	
	
}

void YRescale(GraphLayer Graph)
{//RESCALE Y axis
	Graph.Rescale(OKAXISTYPE_Y);
}


void FillGraphLayerRepeat(GraphLayer Graph, Worksheet OptionsWks, WorksheetPage DataBook, vector<int> OscCount, int GraphNumber, int GraphLayerNumber, StringArray DefaultOptionsLabel, StringArray DataBookNames, int Repeat, bool IsMainSignal)
{//filling single GraphLayer with Labels, Curves & Legend.
//GraphNumber >= 1
	
	Dataset TempDS, OscNumber;
	StringArray VariableNames;
	vector<int> DataBookNumbers(0);
	
	//getting data columns names list
	OptionsWks.Columns(DefaultOptionsLabel[1]).GetStringArray(VariableNames); 
	
	if ((Repeat > 0) && (!IsMainSignal))
	{
		// NEEDS to exept main signal from cycle 
		for (int i=0;i<Repeat; i++)
		{
			for (int v=0; v<VariableNames.GetSize(); v++)
			{
				VariableNames[v] = StringPostfixIncrease(VariableNames[v]);
			}
		}
	}
	
	//getting access to 'OscNumber'
	OscNumber.Attach(OptionsWks.Columns(DefaultOptionsLabel[0]));
	
	TempDS.Attach(OptionsWks.Columns(DefaultOptionsLabel[3]));
	for (int i=0; i<=TempDS.GetUpperBound(); i++)
	{
		DataBookNumbers.Add(TempDS[i]);
	}
	
	//for all ColumnNames
	for (i=0; i<VariableNames.GetSize(); i++)
	{//plotting curves
		Worksheet TempWks(Project.WorksheetPages(DataBookNames[DataBookNumbers[i]-1]).Layers((GraphNumber-1)*OscCount[DataBookNumbers[i]-1]+OscNumber[i]-1));
		
		//data access check. Exit plotting if fail.
		if (!TempWks.Columns(VariableNames[i])) 
		{
			out_str("Error!\nData access error in Graph#"+GraphNumber+", Layer#" + GraphLayerNumber + ", Curve#" + (i+1) + "\nColumn '"+VariableNames[i]+"' at layer #"+((GraphNumber-1)*OscCount[DataBookNumbers[i]-1]+OscNumber[i])+" of '"+DataBookNames[DataBookNumbers[i]-1]+"' WorksheetPage is missing.\nCan't plot another curves in that layer.");
		}
		else 
		{
			//TempDS.Attach(TempWks.Columns(VariableNames[i]));
			Curve cr;
			cr.Attach(TempWks.Columns(VariableNames[i]));
			Graph.AddPlot(cr);
		}
		TempWks.Detach();
	}
	
	//Grouping plots
	TempDS.Attach(OptionsWks.Columns(DefaultOptionsLabel[2]));
	
	i=0;
	//for all plots exept last
	while (i<VariableNames.GetSize()-1)
	{
		//if present Group¹ == next Group¹ 
		if ((TempDS[i] != 0) && (TempDS[i] == TempDS[i+1]))
		{
			int StartIndex = i;
			//serching for last group index
			while ((i<VariableNames.GetSize()-1) && (TempDS[i] == TempDS[i+1])) {	i++;	}
			Graph.GroupPlots(StartIndex, i); //grouping
		}
		i++;
	}
	TempDS.Detach();
	
	OscNumber.Detach();
	VariableNames.RemoveAll();
	DataBookNumbers.RemoveAll();
	
	
}

void AutoTextToGraphLayer(GraphPage gp, int GraphLayerIndex, string text)
{
	int FontSize = 18;
	
	
	GraphLayer gl = gp.Layers(GraphLayerIndex);
	
	//creatting Text object
	GraphObject TextObj = gl.CreateGraphObject(GROT_TEXT); 
	TextObj.Text = text;
	TextObj.Attach = 0;  //attach to Layer
	
	
	Tree trFormat;
	trFormat.Root.Dimension.Units.nVal = UNITS_PAGE; // units in % of page
	trFormat.Root.Font.Size.nVal = FontSize;
	//trFormat.Root.Font.Color.nVal = SYSCOLOR_GRAY;
	if (0 == TextObj.UpdateThemeIDs(trFormat.Root)) { TextObj.ApplyFormat(trFormat, true, true);	}
    
	// right text
	trFormat.Root.Dimension.Left.nVal = 75; // units in % of page
	if (0 == TextObj.UpdateThemeIDs(trFormat.Root)) { TextObj.ApplyFormat(trFormat, true, true);	}
	int TextLeft = TextObj.Left; //in page (i.e. paper, logical) coordinate units
		
	int TextWidth = TextObj.Width;
		
	TextLeft = TextLeft - TextObj.Width;
	TextObj.Left = TextLeft;

		

	trFormat = TextObj.GetFormat(FPB_ALL, FOB_ALL, true, true);
	trFormat.Root.Dimension.Top.nVal = 95;	
	if (0 == TextObj.UpdateThemeIDs(trFormat.Root)) { TextObj.ApplyFormat(trFormat, true, true);	}

		
	gl.Detach();
	trFormat.Remove();
		
	//out_str("TextWidth = " + TextObj.Width);
	//out_str("TextLeft = " + TextObj.Left);
	//out_str("Height = " + TextObj.Height);
}

//----------------------------------------------------------------------------------------------------------------
//=================================================================================================================================
void ProcessGraph(bool y_rescale=false)
{
	//variables
	Dataset GraphLayersCountDS, TempDS;
	string DataBookName, GraphNamePostfix, TemplateGraphPageName, TextPrefix;
	int GraphLayersCount, ShootsCount;
	vector<int> OscCounts(0);
	WorksheetPage OptionsBook, DataBook;
	Worksheet OptionsWks;
	GraphPage TemplateGraphPage;
	vector<string> DataBookNames(1);
	bool NeedSaveToFile = false;
	//bool y_rescale = true;
	//constants
	string OptionsBookName = "GOptions";
	StringArray DefaultGlobalOptionsLabel(7), DefaultOptionsLayersPrefix(2), DefaultOptionsLabel(4);
	DefaultGlobalOptionsLabel[0] = "DataBookName";
	DefaultGlobalOptionsLabel[1] = "OscCount";
	DefaultGlobalOptionsLabel[2] = "TemplateGraph";
	DefaultGlobalOptionsLabel[3] = "LayersCount";
	DefaultGlobalOptionsLabel[4] = "NamePostfix";
	DefaultGlobalOptionsLabel[5] = "TextPrefix";
	DefaultGlobalOptionsLabel[6] = "SaveToFile";
	DefaultOptionsLayersPrefix[0] = "GlobalOptions";
	DefaultOptionsLayersPrefix[1] = "GraphLayer";
	DefaultOptionsLabel[0] = "OscNumber";
	DefaultOptionsLabel[1] = "ColumnName";
	DefaultOptionsLabel[2] = "Group";
	DefaultOptionsLabel[3] = "DataBook";
	
	
	OptionsBook = Project.WorksheetPages(OptionsBookName);
	
	// trying connect to 'Options' WorksheetPage
	if (OptionsBook)
	{
		bool accessError = false;
		
		//GlobalOptions Worksheet existence check
		if (OptionsBook.Layers(DefaultOptionsLayersPrefix[0]))
		{	//get direct access to GlobalOptions Worksheet of Options Worksheetpage
			OptionsWks = OptionsBook.Layers(DefaultOptionsLayersPrefix[0]);		
		}
		else 
		{	ShowMessage("Error!\n'"+DefaultOptionsLayersPrefix[0]+"' layer of '"+OptionsBookName+"' worksheetpage is missing.\n");
			accessError = true;		}
		
	//'OSC-COUNT' column check
		if (!accessError)
		{	
			//existence check
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[1]))
			{
				//get access to 'OscCount' column
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[1]));
				TempDS.SetLowerBound(0);
				for (int i=0; i<=TempDS.GetUpperBound(); i++)
				{
					OscCounts.Add(TempDS[i]); //save oscilligraphs count
				}
				TempDS.Detach();
				
				//OscCount value compatibility check (for Dataset NANUM value OscCount will be 0)
				for (i=0; i<OscCounts.GetSize(); i++)
				{
					if (OscCounts[i] < 1)
					{	
						ShowMessage("Error!\n '"+DefaultGlobalOptionsLabel[1]+"' in '"+OptionsBookName+"' WorksheetPage has incompatible value.\n");
						accessError = true;		
					}
				}
			}
			else 
			{	ShowMessage("Error!\n'"+DefaultGlobalOptionsLabel[1]+"' column of "+OptionsBookName+" worksheet Page is missing!\n");
				accessError = true;		}
		}
		
			
	//'TEMPLATE Graph' column existence check
		if (!accessError) 
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[2]))
			{
				//getting TemplateGraph name
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[2]));
				TempDS.SetLowerBound(0);
    			TempDS.GetText(0, TemplateGraphPageName);
	    		TempDS.Detach();
	    		if (Project.GraphPages(TemplateGraphPageName))
	    		{
	    			TemplateGraphPage = Project.GraphPages(TemplateGraphPageName);
	    		}
	    		else
	    		{
	    			ShowMessage("Error!\n'"+TemplateGraphPageName+"' template GraphPage is missing!\n");
					accessError = true;
	    		}
			}
			else {ShowMessage("Error!\n'"+DefaultGlobalOptionsLabel[2]+"' column of '"+OptionsBookName+"' WotksheetPage is missing!\n");}
		}
		
		

	//column 'LAYERS-COUNT' check
		//existence check
		if (!accessError)
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[3]))
			{
				//getting access to column 'LayersCount'
				GraphLayersCountDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[3]));
				//save count of the graph layers needed to create (for Dataset NANUM value LayersCount will be 0)
				GraphLayersCount = GraphLayersCountDS[0];	
				GraphLayersCountDS.Detach();
			
				//LayersCount value compatibility check (OptionsBook.Layers sufficiency check)
				if (GraphLayersCount < 1) 
				{
					ShowMessage("Error!\n'"+DefaultGlobalOptionsLabel[3]+"' has incompatible value.\n");
					accessError = true;
				}
				else if (GraphLayersCount >= OptionsBook.Layers.Count())
				{
					ShowMessage("Error!\n Not enough Sheets of settings in '"+OptionsBookName+"' WorksheetPage\nfor that '"+DefaultGlobalOptionsLabel[3]+"' number!\n");
					accessError = true;
				}
				else if (GraphLayersCount > TemplateGraphPage.Layers.Count())
				{
					ShowMessage("Error!\n Not enough graph layers("+TemplateGraphPage.Layers.Count()+") in '"+TemplateGraphPage.GetName()+"' GraphPage\n("+DefaultGlobalOptionsLabel[3]+" = "+GraphLayersCount+")\n");
					accessError = true;
				}
			}
			else 
			{	ShowMessage("Error!\n'"+DefaultGlobalOptionsLabel[3]+"' column of "+OptionsBookName+" worksheet Page is missing!\n");
				accessError = true;		}
		}
			
	//column 'DATA_BOOK_NAME' existence check
		if (!accessError)
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[0]))
			{
				//getting DataWorksheetpage name
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[0]));
				TempDS.SetLowerBound(0);
				TempDS.GetText(0, DataBookNames[0]);
				
				// More DataBooks are presented 
				if (TempDS.GetUpperBound() > 0)
				{
					for (int i=1; i<=TempDS.GetUpperBound(); i++)
					{
						string tempstr;
						TempDS.GetText(i, tempstr);
						DataBookNames.Add(tempstr);
					}
				}
				TempDS.Detach();
				if (DataBookNames.GetSize() != OscCounts.GetSize())
				{
					ShowMessage("Error!\nThe number of rows in\n"+DefaultGlobalOptionsLabel[0]+" and "+DefaultGlobalOptionsLabel[1]+"columns are different!");
					accessError = true;
				}
			}
			else 
			{	ShowMessage("Error!\n'DataBookName' column of "+OptionsBookName+" worksheet Page is missing!\n");
				accessError = true;		}
		}
						
	//'NamePostfix' column existence check
		if (!accessError) 
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[4]))
			{
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[4]));
				TempDS.SetLowerBound(0);
				TempDS.GetText(0, GraphNamePostfix);
				TempDS.Detach();
			}
			else {	out_str("Attention.\n'"+DefaultGlobalOptionsLabel[4]+"' column of '"+DefaultOptionsLayersPrefix[0]+"' layer of '"+OptionsBookName+"' WorksheetPage is missing.\n");	}
		}	
		
	//'TextPrefix' column existence check
		if (!accessError) 
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[5]))
			{
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[5]));
				TempDS.SetLowerBound(0);
				TempDS.GetText(0, TextPrefix);
				TempDS.Detach();
			}
			else {	out_str("Attention.\n'"+DefaultGlobalOptionsLabel[5]+"' column of '"+DefaultOptionsLayersPrefix[0]+"' layer of '"+OptionsBookName+"' WorksheetPage is missing.\n");	}
		}	
		
		//'SaveToFile' column existence check
		if (!accessError) 
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[6]))
			{
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[6]));
				TempDS.SetLowerBound(0);
				TempDS.SetUpperBound(0);
				if (TempDS[0] == 1) {NeedSaveToFile = true;}
					else {NeedSaveToFile = false;}
				TempDS.Detach();
			}
			else {	out_str("Attention.\n'"+DefaultGlobalOptionsLabel[6]+"' column of '"+DefaultOptionsLayersPrefix[0]+"' layer of '"+OptionsBookName+"' WorksheetPage is missing.\n");	}
		}	
		
		OptionsWks.Detach();
			
	
	//getting access to DataWorksheetPage
		if (!accessError) 
		{			
		//DataBook access check
			if (Project.WorksheetPages(DataBookNames[0]))
			{//***************************************************************************************************
	    	//Options data check.
	    	
	    		DataBook = Project.WorksheetPages(DataBookNames[0]);
	    		ShootsCount = DataBook.Layers.Count()/OscCounts[0];
	    		
	    		if (DataBookNames.GetSize() > 1)
	    		{//Other DataBooks access check
	    			for (int i=1; i<DataBookNames.GetSize(); i++)
	    			{
	    				if(Project.WorksheetPages(DataBookNames[i])) 
	    				{
	    					WorksheetPage tempwp = Project.WorksheetPages(DataBookNames[i]);
	    					int ShootsCount2 = tempwp.Layers.Count()/OscCounts[i];
	    					if (ShootsCount != ShootsCount2)
	    					{
	    						string lessormore = "less";
	    						if (ShootsCount2 > ShootsCount) {lessormore = "more";}
	    						ShowMessage("Error!\n'"+ DataBookNames[i] + "' WorksheetPage has "+lessormore+" shoots than '"+DataBookNames[0]+"'!\n(Shoots count = LayersCount / OscCount)\n");
	    						accessError = true;
	    						tempwp.Detach();
	    						break;
	    					}
	    					tempwp.Detach();

	    						
	    				}
	    				else 
	    				{
	    					ShowMessage("Error!\n'"+ DataBookNames[i] + "' WorksheetPage with data is missing!\n");
	    					accessError = true;
	    					break;
	    				}
	    			}
	    		}
	    		
	    		//'GraphLayer#' layers data check
				for (int c=1; c<=GraphLayersCount; c++)
	    		{
	    			//'Oscillograph_#' layer existion check
	    			if (!OptionsBook.Layers(DefaultOptionsLayersPrefix[1]+c)) 
	    			{
	    				ShowMessage("Error!\n'"+DefaultOptionsLayersPrefix[1]+c + "' layer of '"+OptionsBookName+"' WorksheetPage is missing!\n");
	    				accessError = true;
	    				break;
	    			}
	    			
	    			//'GraphLayer#' layer existence check
	    			Worksheet wks = OptionsBook.Layers(DefaultOptionsLayersPrefix[1]+c);
	    			
	    		//'ColumnName' column check
    				//existion check
	    			if (!wks.Columns(DefaultOptionsLabel[1])) 
	    			{
	    				ShowMessage("Error!\n '"+DefaultOptionsLabel[1]+"' column access error in "+OptionsBookName+".Layer["+c+"].\n");
		    			accessError = true;
		    		}
	    			else if (0 > wks.Columns(DefaultOptionsLabel[1]).GetUpperBound())
	    			{//is empty check
	    				ShowMessage("Error!\n '"+DefaultOptionsLabel[1]+"' column of "+OptionsBookName+"."+DefaultOptionsLayersPrefix[1]+c+" is empty.\n");
	    				accessError = true;	    				
		    		}
		    		else 
	    			{
	    				StringArray sr;
	    				//getting array of column names [0..column.size]
	    				wks.Columns(DefaultOptionsLabel[1]).GetStringArray(sr, 0, wks.Columns(DefaultOptionsLabel[1]).GetUpperBound());
		    			for (int j=0; j<sr.GetSize(); j++)
		    			{
	    					if (sr[j] == "") 
	    					{
	    						ShowMessage("Error!\n "+OptionsBookName+".["+DefaultOptionsLayersPrefix[1]+c+"]."+DefaultOptionsLabel[1]+"["+(j+1)+"] is empty.\n");
	    						accessError = true;
	    					}
		    			}
		    		}//------------------------------
		    			
		    	//'OscNumber' check
		    		//OscCount compatibility check
		    		for (int i=0; i<OscCounts.GetSize(); i++)
		    		{
		    			if (OscCounts[i] > Project.WorksheetPages(DataBookNames[i]).Layers.Count())
	    				{
	    					ShowMessage("Error!\nWrong value in '"+DefaultGlobalOptionsLabel[1]+"["+i+"]'.\n(Not enough layers in '"+DataBookNames[i]+"')\n");
	    					accessError = true;
	    				}
		    		}
	    			
	    			//'OscNumber' existence check
	    			if (!wks.Columns(DefaultOptionsLabel[0])) 
	    			{//if missing
	    				ShowMessage("Error!\n '"+DefaultOptionsLabel[0]+"' column access error in "+OptionsBookName+".Layer[" + c + "].\n");
	    				accessError = true;
	    			}
	    			else 
	    			{//if existing
	    				//getting access to 'OscNumber' column
	    				TempDS.Attach(wks.Columns(DefaultOptionsLabel[0]));
	    				Dataset TempBookNumber;
	    				TempBookNumber.Attach(wks.Columns(DefaultOptionsLabel[3]));
	    				
	    				//changeing lower and upper index to prevent access errors
	    				TempDS.SetLowerBound(0);
	    				TempBookNumber.SetLowerBound(0);
	    				TempDS.SetUpperBound(wks.Columns(DefaultOptionsLabel[1]).GetUpperBound());
	    				TempBookNumber.SetUpperBound(wks.Columns(DefaultOptionsLabel[1]).GetUpperBound());

	    				//for all cells in column
	    				for (int k=0; k<=wks.Columns(DefaultOptionsLabel[1]).GetUpperBound(); k++)
	    				{
	    					//value format check
	    					if (TempDS[k] == NANUM)
	    					{
	    						ShowMessage("Error!\nValue "+DefaultOptionsLabel[0]+"["+(k+1)+"] in '"+DefaultOptionsLayersPrefix[1]+c+"' of '"+OptionsBookName+"' WorkshhetPage has incompatible format.\n");
	    						accessError = true;
	    					}
	    					//value compatibility check
	    					else if (TempDS[k] > OscCounts[TempBookNumber[k]-1])
	    					{
	    						ShowMessage("Error!\nWrong value in "+DefaultOptionsLabel[0]+"["+(k+1)+"] in '"+DefaultOptionsLayersPrefix[1]+c+"' of '"+OptionsBookName+"' WorkshhetPage.\n("+DefaultOptionsLabel[0]+" can't be more than "+DefaultGlobalOptionsLabel[1]+")\n");
	    						accessError = true;
	    					} 
	    				}
	    				TempDS.Detach();
	    				TempBookNumber.Detach();
	    				
	    			}//-----------------------------------
	    						
	    			//breaking cyle if error
	    			if (accessError) {	break;	}
	    			
	    		//'Group' column check
	    			//existion check
	    			if (!wks.Columns(DefaultOptionsLabel[2])) 
	    			{
	    				out_str("Error!\n '"+DefaultOptionsLabel[2]+"' column access error in "+OptionsBookName+".Layer[" + c + "].\n");
	    				accessError = true;
	    			}
	    			else 
	    			{
	    				//getting access to 'Group' column
	    				TempDS.Attach(wks.Columns(DefaultOptionsLabel[2]));
	    				
	    				//changeing lower and upper index to prevent access errors
	    				TempDS.SetLowerBound(0);
	    				TempDS.SetUpperBound(wks.Columns(DefaultOptionsLabel[1]).GetUpperBound());

	    				//for all cells in column
	    				for (int k=0; k<=wks.Columns(DefaultOptionsLabel[1]).GetUpperBound(); k++)
	    				{
	    					//replaceing nonumeric values
	    					if (TempDS[k] == NANUM)
	    					{
	    						TempDS[k] = 0;
	    					}
	    				}
	    				TempDS.Detach();
	    			}
	    			wks.Detach();
	    				
	    		}
	    		//end of Options data check.
	    		//****************************************************************************************************************
	    		
	    		if (!accessError)
	    	   	{	
					if (Project.GraphPages(TemplateGraphPageName))
					//Graph template access check
					{
						TemplateGraphPage = Project.GraphPages(TemplateGraphPageName);
						//TemplateGraph layers sufficiency check
						if (TemplateGraphPage.Layers.Count() >= GraphLayersCount)
						{	
							//DataBook.Layers.Count multiplicity check
							if ((DataBook.Layers.Count()%OscCounts[0] > 0) || (DataBook.Layers.Count() < OscCounts[0])) 
								{ShowMessage("Warning! \n DataBook layers count is \n not a multiple of OscCount. \nSome data won't be plotted.\n");}
							
							//getting template format as Tree
							Tree trTemplateFormat;
							trTemplateFormat = TemplateGraphPage.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
							
							
							for (int i=0; i<DataBook.Layers.Count()/OscCounts[0]; i++)
							{//for all Layers/OscCount
								GraphPage NewGraphPage; //create new GraphPage window
						    	NewGraphPage.Create("origin");
						    	
						    	//set first 4 letters of Layer.Name as GraphPage Name...
						    	string name;
								name = DataBook.Layers(i*OscCounts[0]).GetName().Mid(0,4) + GraphNamePostfix;
								while (Project.WorksheetPages(name))
								{
									name = StringPrefixIncrease(name);
								}
				    		    //NewGraphPage.SetName(DataBook.Layers(i*OscCounts[0]).GetName().Mid(0,4) + GraphNamePostfix);
				    		    NewGraphPage.SetName(DataBook.Layers(i*OscCounts[0]).GetName());
				    		    NewGraphPage.Show = 0;  //hiding new graph window
				    		    
				    		    //...and add Postfix to it for LongName
					   		    NewGraphPage.SetLongName(NewGraphPage.GetName());
					   		    
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
				    		    			for (int p=NewGraphPage.Layers(j).DataPlots.Count(); p>0; p--)
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
				    		    	FillGraphLayer(NewGraphPage.Layers(j), OptionsBook.Layers(DefaultOptionsLayersPrefix[1] + (j+1)), DataBook, OscCounts, i+1, j+1, DefaultOptionsLabel, DataBookNames, y_rescale);
				    		    }
				    		    
				    		    
				    		    
				    		    bool FormatUpdErr = NewGraphPage.ApplyFormat( trTemplateFormat, true, true, true ); //changing format
				    		    if (!FormatUpdErr) {out_str("Warning!\n GraphPage[index=" + i + "] ('"+NewGraphPage.GetName()+"') format update failed!\n");}
				    		    	
				    		    	//trTemplateFormat.Save("c:\Tektronix\TemlateFormat.xml" );
				    		    	
				    		    // string text = TextPrefix + DataBook.Layers(i*OscCounts[0]).GetName().Mid(0,4);
				    		    // string text = TextPrefix + DataBook.Layers(i*OscCounts[0]).GetName().Mid(4, 11);
				    		    string text = TextPrefix + DataBook.Layers(i*OscCounts[0]).GetName();
				    		    AutoTextToGraphLayer(NewGraphPage, GraphLayersCount-1, text);
				    		    
				    		    if (y_rescale)
				    		    {
				    		    	for (int j=0; j<GraphLayersCount; j++)
									{//for all GraphLayers in NewGraphPage
										YRescale(NewGraphPage.Layers(j));
									}
				    		    }
				    		    
				    		    
				    		    if (NeedSaveToFile)
									{
										string FullPath = Project.GetPath() + NewGraphPage.GetName() + ".ogg";
										NewGraphPage.SaveToFile(FullPath);
									}
				    		    
				    		    
				    		    //NewGraphPage.Show = 0; //hiding new graph window
				    		    NewGraphPage.Detach();
				    		    //break;
				    		    
							} // For all Layers in DataBook
							trTemplateFormat.Remove();
							
						}// TemplateGraph Layers count check
						else {ShowMessage("Error!\n Not enough Layers in TemplateGraphPage!\n");}
						TemplateGraphPage.Detach();
						
					}//'TemplateGraph' check
					else {ShowMessage("Error!\n Can't connect to TemplateGraphPage!\n");}
					
				}//accessError
    		    else {ShowMessage("Error!\n Data access error!\n");}
				
			}// input book check
			else {ShowMessage("Error!\n Can't connect to WorksheetPage with curve data! \n(Wrong DataBookName)\n"); }
			
		}//accessError
		
	}//'Options' book check
	else {ShowMessage("Error!\n Can't connect to table 'Options'!\n");}
		
	OptionsBook.Detach();
	DataBook.Detach();

	OscCounts.RemoveAll();
	DataBookNames.RemoveAll();
	DefaultGlobalOptionsLabel.RemoveAll();
	DefaultOptionsLayersPrefix.RemoveAll();
	DefaultOptionsLabel.RemoveAll();
}

// ProcessGraph_END
//==================================================================================================================================
void ProcessGraphRepeat(int Repeat_Count)
{
	//variables
	Dataset GraphLayersCountDS, TempDS;
	string DataBookName, GraphNamePostfix, TemplateGraphPageName, TextPrefix;
	int GraphLayersCount, ShootsCount;
	vector<int> OscCounts(0);
	WorksheetPage OptionsBook, DataBook;
	Worksheet OptionsWks;
	GraphPage TemplateGraphPage;
	vector<string> DataBookNames(1);
	bool NeedSaveToFile = true;
	
	
	//constants
	string OptionsBookName = "GOptions";
	StringArray DefaultGlobalOptionsLabel(7), DefaultOptionsLayersPrefix(2), DefaultOptionsLabel(4);
	DefaultGlobalOptionsLabel[0] = "DataBookName";
	DefaultGlobalOptionsLabel[1] = "OscCount";
	DefaultGlobalOptionsLabel[2] = "TemplateGraph";
	DefaultGlobalOptionsLabel[3] = "LayersCount";
	DefaultGlobalOptionsLabel[4] = "NamePostfix";
	DefaultGlobalOptionsLabel[5] = "TextPrefix";
	DefaultGlobalOptionsLabel[6] = "SaveToFile";
	DefaultOptionsLayersPrefix[0] = "GlobalOptions";
	DefaultOptionsLayersPrefix[1] = "GraphLayer";
	DefaultOptionsLabel[0] = "OscNumber";
	DefaultOptionsLabel[1] = "ColumnName";
	DefaultOptionsLabel[2] = "Group";
	DefaultOptionsLabel[3] = "DataBook";
	
	
	OptionsBook = Project.WorksheetPages(OptionsBookName);
	
	// trying connect to 'Options' WorksheetPage
	if (OptionsBook)
	{
		bool accessError = false;
		
		//GlobalOptions Worksheet existence check
		if (OptionsBook.Layers(DefaultOptionsLayersPrefix[0]))
		{	//get direct access to GlobalOptions Worksheet of Options Worksheetpage
			OptionsWks = OptionsBook.Layers(DefaultOptionsLayersPrefix[0]);		
		}
		else 
		{	ShowMessage("Error!\n'"+DefaultOptionsLayersPrefix[0]+"' layer of '"+OptionsBookName+"' worksheetpage is missing.\n");
			accessError = true;		}
		
	//'OSC-COUNT' column check
		if (!accessError)
		{	
			//existence check
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[1]))
			{
				//get access to 'OscCount' column
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[1]));
				TempDS.SetLowerBound(0);
				for (int i=0; i<=TempDS.GetUpperBound(); i++)
				{
					OscCounts.Add(TempDS[i]); //save oscilligraphs count
				}
				TempDS.Detach();
				
				//OscCount value compatibility check (for Dataset NANUM value OscCount will be 0)
				for (i=0; i<OscCounts.GetSize(); i++)
				{
					if (OscCounts[i] < 1)
					{	
						ShowMessage("Error!\n '"+DefaultGlobalOptionsLabel[1]+"' in '"+OptionsBookName+"' WorksheetPage has incompatible value.\n");
						accessError = true;		
					}
				}
			}
			else 
			{	ShowMessage("Error!\n'"+DefaultGlobalOptionsLabel[1]+"' column of "+OptionsBookName+" worksheet Page is missing!\n");
				accessError = true;		}
		}
		
			
	//'TEMPLATE Graph' column existence check
		if (!accessError) 
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[2]))
			{
				//getting TemplateGraph name
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[2]));
				TempDS.SetLowerBound(0);
    			TempDS.GetText(0, TemplateGraphPageName);
	    		TempDS.Detach();
	    		if (Project.GraphPages(TemplateGraphPageName))
	    		{
	    			TemplateGraphPage = Project.GraphPages(TemplateGraphPageName);
	    		}
	    		else
	    		{
	    			ShowMessage("Error!\n'"+TemplateGraphPageName+"' template GraphPage is missing!\n");
					accessError = true;
	    		}
			}
			else {ShowMessage("Error!\n'"+DefaultGlobalOptionsLabel[2]+"' column of '"+OptionsBookName+"' WotksheetPage is missing!\n");}
		}
		
		

	//column 'LAYERS-COUNT' check
		//existence check
		if (!accessError)
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[3]))
			{
				//getting access to column 'LayersCount'
				GraphLayersCountDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[3]));
				//save count of the graph layers needed to create (for Dataset NANUM value LayersCount will be 0)
				GraphLayersCount = GraphLayersCountDS[0];	
				GraphLayersCountDS.Detach();
			
				//LayersCount value compatibility check (OptionsBook.Layers sufficiency check)
				if (GraphLayersCount < 1) 
				{
					ShowMessage("Error!\n'"+DefaultGlobalOptionsLabel[3]+"' has incompatible value.\n");
					accessError = true;
				}
				else if (GraphLayersCount >= OptionsBook.Layers.Count())
				{
					ShowMessage("Error!\n Not enough Sheets of settings in '"+OptionsBookName+"' WorksheetPage\nfor that '"+DefaultGlobalOptionsLabel[3]+"' number!\n");
					accessError = true;
				}
				else if (GraphLayersCount > TemplateGraphPage.Layers.Count())
				{
					ShowMessage("Error!\n Not enough graph layers("+TemplateGraphPage.Layers.Count()+") in '"+TemplateGraphPage.GetName()+"' GraphPage\n("+DefaultGlobalOptionsLabel[3]+" = "+GraphLayersCount+")\n");
					accessError = true;
				}
			}
			else 
			{	ShowMessage("Error!\n'"+DefaultGlobalOptionsLabel[3]+"' column of "+OptionsBookName+" worksheet Page is missing!\n");
				accessError = true;		}
		}
			
	//column 'DATA_BOOK_NAME' existence check
		if (!accessError)
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[0]))
			{
				//getting DataWorksheetpage name
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[0]));
				TempDS.SetLowerBound(0);
				TempDS.GetText(0, DataBookNames[0]);
				
				// More DataBooks are presented 
				if (TempDS.GetUpperBound() > 0)
				{
					for (int i=1; i<=TempDS.GetUpperBound(); i++)
					{
						string tempstr;
						TempDS.GetText(i, tempstr);
						DataBookNames.Add(tempstr);
					}
				}
				TempDS.Detach();
				if (DataBookNames.GetSize() != OscCounts.GetSize())
				{
					ShowMessage("Error!\nThe number of rows in\n"+DefaultGlobalOptionsLabel[0]+" and "+DefaultGlobalOptionsLabel[1]+"columns are different!");
					accessError = true;
				}
			}
			else 
			{	ShowMessage("Error!\n'DataBookName' column of "+OptionsBookName+" worksheet Page is missing!\n");
				accessError = true;		}
		}
						
	//'NamePostfix' column existence check
		if (!accessError) 
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[4]))
			{
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[4]));
				TempDS.SetLowerBound(0);
				TempDS.GetText(0, GraphNamePostfix);
				TempDS.Detach();
			}
			else {	out_str("Attention.\n'"+DefaultGlobalOptionsLabel[4]+"' column of '"+DefaultOptionsLayersPrefix[0]+"' layer of '"+OptionsBookName+"' WorksheetPage is missing.\n");	}
		}	
		
	//'TextPrefix' column existence check
		if (!accessError) 
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[5]))
			{
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[5]));
				TempDS.SetLowerBound(0);
				TempDS.GetText(0, TextPrefix);
				TempDS.Detach();
			}
			else {	out_str("Attention.\n'"+DefaultGlobalOptionsLabel[5]+"' column of '"+DefaultOptionsLayersPrefix[0]+"' layer of '"+OptionsBookName+"' WorksheetPage is missing.\n");	}
		}	
		
	//'SaveToFile' column existence check
		if (!accessError) 
		{
			if (OptionsWks.Columns(DefaultGlobalOptionsLabel[6]))
			{
				TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[6]));
				TempDS.SetLowerBound(0);
				TempDS.SetUpperBound(0);
				if (TempDS[0] == 1) {NeedSaveToFile = true;}
					else {NeedSaveToFile = false;}
				TempDS.Detach();
			}
			else {	out_str("Attention.\n'"+DefaultGlobalOptionsLabel[6]+"' column of '"+DefaultOptionsLayersPrefix[0]+"' layer of '"+OptionsBookName+"' WorksheetPage is missing.\n");	}
		}	
		
		OptionsWks.Detach();
			
	
	//getting access to DataWorksheetPage
		if (!accessError) 
		{			
		//DataBook access check
			if (Project.WorksheetPages(DataBookNames[0]))
			{//***************************************************************************************************
	    	//Options data check.
	    	
	    		DataBook = Project.WorksheetPages(DataBookNames[0]);
	    		ShootsCount = DataBook.Layers.Count()/OscCounts[0];
	    		
	    		if (DataBookNames.GetSize() > 1)
	    		{//Other DataBooks access check
	    			for (int i=1; i<DataBookNames.GetSize(); i++)
	    			{
	    				if(Project.WorksheetPages(DataBookNames[i])) 
	    				{
	    					WorksheetPage tempwp = Project.WorksheetPages(DataBookNames[i]);
	    					int ShootsCount2 = tempwp.Layers.Count()/OscCounts[i];
	    					if (ShootsCount != ShootsCount2)
	    					{
	    						string lessormore = "less";
	    						if (ShootsCount2 > ShootsCount) {lessormore = "more";}
	    						ShowMessage("Error!\n'"+ DataBookNames[i] + "' WorksheetPage has "+lessormore+" shoots than '"+DataBookNames[0]+"'!\n(Shoots count = LayersCount / OscCount)\n");
	    						accessError = true;
	    						tempwp.Detach();
	    						break;
	    					}
	    					tempwp.Detach();

	    						
	    				}
	    				else 
	    				{
	    					ShowMessage("Error!\n'"+ DataBookNames[i] + "' WorksheetPage with data is missing!\n");
	    					accessError = true;
	    					break;
	    				}
	    			}
	    		}
	    		
	    		//'GraphLayer#' layers data check
				for (int c=1; c<=GraphLayersCount; c++)
	    		{
	    			//'Oscillograph_#' layer existion check
	    			if (!OptionsBook.Layers(DefaultOptionsLayersPrefix[1]+c)) 
	    			{
	    				ShowMessage("Error!\n'"+DefaultOptionsLayersPrefix[1]+c + "' layer of '"+OptionsBookName+"' WorksheetPage is missing!\n");
	    				accessError = true;
	    				break;
	    			}
	    			
	    			//'GraphLayer#' layer existence check
	    			Worksheet wks = OptionsBook.Layers(DefaultOptionsLayersPrefix[1]+c);
	    			
	    		//'ColumnName' column check
    				//existion check
	    			if (!wks.Columns(DefaultOptionsLabel[1])) 
	    			{
	    				ShowMessage("Error!\n '"+DefaultOptionsLabel[1]+"' column access error in "+OptionsBookName+".Layer["+c+"].\n");
		    			accessError = true;
		    		}
	    			else if (0 > wks.Columns(DefaultOptionsLabel[1]).GetUpperBound())
	    			{//is empty check
	    				ShowMessage("Error!\n '"+DefaultOptionsLabel[1]+"' column of "+OptionsBookName+"."+DefaultOptionsLayersPrefix[1]+c+" is empty.\n");
	    				accessError = true;	    				
		    		}
		    		else 
	    			{
	    				StringArray sr;
	    				//getting array of column names [0..column.size]
	    				wks.Columns(DefaultOptionsLabel[1]).GetStringArray(sr, 0, wks.Columns(DefaultOptionsLabel[1]).GetUpperBound());
		    			for (int j=0; j<sr.GetSize(); j++)
		    			{
	    					if (sr[j] == "") 
	    					{
	    						ShowMessage("Error!\n "+OptionsBookName+".["+DefaultOptionsLayersPrefix[1]+c+"]."+DefaultOptionsLabel[1]+"["+(j+1)+"] is empty.\n");
	    						accessError = true;
	    					}
		    			}
		    		}//------------------------------
		    			
		    	//'OscNumber' check
		    		//OscCount compatibility check
		    		for (int i=0; i<OscCounts.GetSize(); i++)
		    		{
		    			if (OscCounts[i] > Project.WorksheetPages(DataBookNames[i]).Layers.Count())
	    				{
	    					ShowMessage("Error!\nWrong value in '"+DefaultGlobalOptionsLabel[1]+"["+i+"]'.\n(Not enough layers in '"+DataBookNames[i]+"')\n");
	    					accessError = true;
	    				}
		    		}
	    			
	    			//'OscNumber' existence check
	    			if (!wks.Columns(DefaultOptionsLabel[0])) 
	    			{//if missing
	    				ShowMessage("Error!\n '"+DefaultOptionsLabel[0]+"' column access error in "+OptionsBookName+".Layer[" + c + "].\n");
	    				accessError = true;
	    			}
	    			else 
	    			{//if existing
	    				//getting access to 'OscNumber' column
	    				TempDS.Attach(wks.Columns(DefaultOptionsLabel[0]));
	    				Dataset TempBookNumber;
	    				TempBookNumber.Attach(wks.Columns(DefaultOptionsLabel[3]));
	    				
	    				//changeing lower and upper index to prevent access errors
	    				TempDS.SetLowerBound(0);
	    				TempBookNumber.SetLowerBound(0);
	    				TempDS.SetUpperBound(wks.Columns(DefaultOptionsLabel[1]).GetUpperBound());
	    				TempBookNumber.SetUpperBound(wks.Columns(DefaultOptionsLabel[1]).GetUpperBound());

	    				//for all cells in column
	    				for (int k=0; k<=wks.Columns(DefaultOptionsLabel[1]).GetUpperBound(); k++)
	    				{
	    					//value format check
	    					if (TempDS[k] == NANUM)
	    					{
	    						ShowMessage("Error!\nValue "+DefaultOptionsLabel[0]+"["+(k+1)+"] in '"+DefaultOptionsLayersPrefix[1]+c+"' of '"+OptionsBookName+"' WorkshhetPage has incompatible format.\n");
	    						accessError = true;
	    					}
	    					//value compatibility check
	    					else if (TempDS[k] > OscCounts[TempBookNumber[k]-1])
	    					{
	    						ShowMessage("Error!\nWrong value in "+DefaultOptionsLabel[0]+"["+(k+1)+"] in '"+DefaultOptionsLayersPrefix[1]+c+"' of '"+OptionsBookName+"' WorkshhetPage.\n("+DefaultOptionsLabel[0]+" can't be more than "+DefaultGlobalOptionsLabel[1]+")\n");
	    						accessError = true;
	    					} 
	    				}
	    				TempDS.Detach();
	    				TempBookNumber.Detach();
	    				
	    			}//-----------------------------------
	    						
	    			//breaking cyle if error
	    			if (accessError) {	break;	}
	    			
	    		//'Group' column check
	    			//existion check
	    			if (!wks.Columns(DefaultOptionsLabel[2])) 
	    			{
	    				out_str("Error!\n '"+DefaultOptionsLabel[2]+"' column access error in "+OptionsBookName+".Layer[" + c + "].\n");
	    				accessError = true;
	    			}
	    			else 
	    			{
	    				//getting access to 'Group' column
	    				TempDS.Attach(wks.Columns(DefaultOptionsLabel[2]));
	    				
	    				//changeing lower and upper index to prevent access errors
	    				TempDS.SetLowerBound(0);
	    				TempDS.SetUpperBound(wks.Columns(DefaultOptionsLabel[1]).GetUpperBound());

	    				//for all cells in column
	    				for (int k=0; k<=wks.Columns(DefaultOptionsLabel[1]).GetUpperBound(); k++)
	    				{
	    					//replaceing nonumeric values
	    					if (TempDS[k] == NANUM)
	    					{
	    						TempDS[k] = 0;
	    					}
	    				}
	    				TempDS.Detach();
	    			}
	    			wks.Detach();
	    				
	    		}
	    		//end of Options data check.
	    		//****************************************************************************************************************
	    		
	    		if (!accessError)
	    	   	{	
					if (Project.GraphPages(TemplateGraphPageName))
					//Graph template access check
					{
						TemplateGraphPage = Project.GraphPages(TemplateGraphPageName);
						//TemplateGraph layers sufficiency check
						if (TemplateGraphPage.Layers.Count() >= GraphLayersCount)
						{	
							//DataBook.Layers.Count multiplicity check
							if ((DataBook.Layers.Count()%OscCounts[0] > 0) || (DataBook.Layers.Count() < OscCounts[0])) 
								{ShowMessage("Warning! \n DataBook layers count is \n not a multiple of OscCount. \nSome data won't be plotted.\n");}
							
							//getting template format as Tree
							Tree trTemplateFormat;
							trTemplateFormat = TemplateGraphPage.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
							
							
							for (int i=0; i<DataBook.Layers.Count()/OscCounts[0]; i++)
							{//for all Layers/OscCount
								for (int r=0; r<Repeat_Count; r++)
								{
									GraphPage NewGraphPage; //create new GraphPage window
									NewGraphPage.Create("origin");
									
									//set first 3 letters of Layer.Name as GraphPage Name...
									string AddPostfix = "Sample00";
									AddPostfix =AddPostfix.Mid(0, 8 - (r+1)/10) + (r+1);
									
									NewGraphPage.SetName(DataBook.Layers(i*OscCounts[0]).GetName().Mid(0,4) + GraphNamePostfix + AddPostfix);
									NewGraphPage.Show = 0;  //hiding new graph window
									
									//...and add Postfix to it for LongName
									NewGraphPage.SetLongName(NewGraphPage.GetName());
									
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
												for (int p=NewGraphPage.Layers(j).DataPlots.Count(); p>0; p--)
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
										if (j==0) 
										{
											FillGraphLayerRepeat(NewGraphPage.Layers(j), OptionsBook.Layers(DefaultOptionsLayersPrefix[1] + (j+1)), DataBook, OscCounts, i+1, j+1, DefaultOptionsLabel, DataBookNames, r, true);
										}
										else
										{
											FillGraphLayerRepeat(NewGraphPage.Layers(j), OptionsBook.Layers(DefaultOptionsLayersPrefix[1] + (j+1)), DataBook, OscCounts, i+1, j+1, DefaultOptionsLabel, DataBookNames, r, false);
										}
									}
									
									
									
									bool FormatUpdErr = NewGraphPage.ApplyFormat( trTemplateFormat, true, true, true ); //changing format
									if (!FormatUpdErr) {out_str("Warning!\n GraphPage[index=" + i + "] ('"+NewGraphPage.GetName()+"') format update failed!\n");}
										
										//trTemplateFormat.Save("c:\Tektronix\TemlateFormat.xml" );
										
									string text = TextPrefix + DataBook.Layers(i*OscCounts[0]).GetName().Mid(0,4);		
									AutoTextToGraphLayer(NewGraphPage, GraphLayersCount-1, text);
										
									if (NeedSaveToFile)
									{
										string FullPath = Project.GetPath() + NewGraphPage.GetName() + ".ogg";
										NewGraphPage.SaveToFile(FullPath);
									}
									
									//NewGraphPage.Show = 0; //hiding new graph window
									NewGraphPage.Detach();
								}
				    		    
							} // For all Layers in DataBook
							trTemplateFormat.Remove();
							
						}// TemplateGraph Layers count check
						else {ShowMessage("Error!\n Not enough Layers in TemplateGraphPage!\n");}
						TemplateGraphPage.Detach();
						
					}//'TemplateGraph' check
					else {ShowMessage("Error!\n Can't connect to TemplateGraphPage!\n");}
					
				}//accessError
    		    else {ShowMessage("Error!\n Data access error!\n");}
				
			}// input book check
			else {ShowMessage("Error!\n Can't connect to WorksheetPage with curve data! \n(Wrong DataBookName)\n"); }
			
		}//accessError
		
	}//'Options' book check
	else {ShowMessage("Error!\n Can't connect to table 'Options'!\n");}
		
	OptionsBook.Detach();
	DataBook.Detach();

	OscCounts.RemoveAll();
	DataBookNames.RemoveAll();
	DefaultGlobalOptionsLabel.RemoveAll();
	DefaultOptionsLayersPrefix.RemoveAll();
	DefaultOptionsLabel.RemoveAll();
}
