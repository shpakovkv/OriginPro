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
#include <../OriginLab/fft.h>
#include <fft_utils.h>
////////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.


////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.

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

void ShowMessage(string m)
{// message box and CommandWindow log
// CommandWindow log appears only when it active
	out_str(m);
	LT_execute("type -b " + m);
}

void Test(int n)
{
	n -= 1;
	out_str("n = " + n);
}

int GetColumnIndex (Worksheet wks, string ColName)
{
	int index=-1;
	for (int i=0; i<wks.Columns.Count(); i++)
	{
		if (wks.Columns(i).GetName() == ColName) {index = i;}
	}
	return index;
	// return "-1" if worksheet don't contain column with this name
}



  //=======================================================================================================================
 //		GET X column INDEX 
//======================================

int GetDataXColumnIndex(Worksheet wks, string ColName)
{// finding column with 'X' data for column 'ColName' containing 'Y' data in order to plot graph or something...
	int index;
	index = GetColumnIndex(wks, ColName);
	
	if ((index == -1) || (wks.Columns(index).GetType() == OKDATAOBJ_DESIGNATION_X))
	{
		return -1;
	}
	else
	{
		for (int i=index; i>=0; i--)
		{
			if (wks.Columns(i).GetType() == OKDATAOBJ_DESIGNATION_X) 
			{return i;}
		}
		return index;
	}
}

void SpectrumFillGraphLayer(GraphLayer Graph, Worksheet wks, int GraphNumber, int GraphLayerNumber, StringArray VariableNames, vector<int> Groups)
{//filling single GraphLayer with Labels, Curves & Legend.
//GraphNumber >= 1
	
	//getting data columns names list
	//OptionsWks.Columns(DefaultOptionsLabel[1]).GetStringArray(VariableNames); 
	Dataset TempDS;
	
	
	//for all ColumnNames
	for (int i=0; i<VariableNames.GetSize(); i++)
	{//plotting curves
		
		//data access check. Exit plotting if fail.
		if (!wks.Columns(VariableNames[i])) 
		{
			out_str("Error!\nData access error in Graph#"+GraphNumber+", Layer#" + GraphLayerNumber + ", Curve#" + (i+1) + "\nColumn '"+VariableNames[i]+"' at worksheet #"+wks.GetName()+" is missing.\nCan't plot another curves in that layer.");
		}
		else 
		{
			//TempDS.Attach(TempWks.Columns(VariableNames[i]));
			Curve cr;
			cr.Attach(wks.Columns(VariableNames[i]));
			Graph.AddPlot(cr);
		}
	}
	
	i=0;
	//for all plots exept last
	while (i<VariableNames.GetSize()-1)
	{
		//if present Group¹ == next Group¹ 
		if ((Groups[i] != 0) && (Groups[i] == Groups[i+1]))
		{
			int StartIndex = i;
			//serching for last group index
			while ((i<VariableNames.GetSize()-1) && (Groups[i] == Groups[i+1])) {	i++;	}
			Graph.GroupPlots(StartIndex, i); //grouping
		}
		i++;
	}
}

void SpectrumAutoTextToGraphLayer(GraphPage gp, int GraphLayerIndex, string text)
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
	trFormat.Root.Dimension.Left.nVal = 95; // units in % of page
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

void SpectrumGraphProcess(string TemplateGraphPageName, Worksheet Spectrum_wks, int Sampling_Count)
{
	
	GraphPage TemplateGraphPage;
	WorksheetPage SpectrumBook;
	vector<int> OscCount;
	Tree trTemplateFormat;
	//int Sampling_Count;
	
	vector<int> Groups;
	
	StringArray VariableNames1[1], VariableNames2[3];
	
	
	int GraphLayersCount = 3;
	
	StringArray VariableNames(4);
	VariableNames[0] = "Freq";
	VariableNames[1] = "Amp";
	VariableNames[2] = "X";
	VariableNames[3] = "Line1x";
	VariableNames[4] = "X";
	VariableNames[5] = "Line2x";
	
	TemplateGraphPage = Project.GraphPages(TemplateGraphPageName);
	if(TemplateGraphPage)
	{
		//TemplateGraph layers sufficiency check
		if (TemplateGraphPage.Layers.Count() >= GraphLayersCount)
		{	
			//getting template format as Tree
			trTemplateFormat = TemplateGraphPage.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
			
				
			for (int i=0; i<Sampling_Count; i++)
			{//for all Samples
				GraphPage NewGraphPage; //create new GraphPage window
				NewGraphPage.Create("origin");
				
				string GraphNamePostfix = "Sample00";
				int j = (i+1)/10;
				GraphNamePostfix = GraphNamePostfix.Mid(0, 8 - j) + (i+1);
				
				
				//set first 4 letters of Layer.Name as GraphPage Name...
				NewGraphPage.SetName(Spectrum_wks.GetName().Mid(0,4) + GraphNamePostfix);
				NewGraphPage.Show = 0;  //hiding new graph window
					
			    //...and add Postfix to it for LongName
				NewGraphPage.SetLongName(Spectrum_wks.GetName() + " - " + GraphNamePostfix);
				
				//add enough Layers with Template format to NewGraphPage	
		    	//NewGraphPage.Layers(0).Destroy();
				
				NewGraphPage.AddLayers(TemplateGraphPage);
				for (j=1; j<=GraphLayersCount; j++) 
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
				//deleting excess (first one) layer
				NewGraphPage.Layers(0).Delete();
					
				//filling Layers of NewGraphPage with Curves
				for (j=0; j<GraphLayersCount; j++)
				{//for all GraphLayers in NewGraphPage
					if (j == 0) 			{SpectrumFillGraphLayer(NewGraphPage.Layers(j), Spectrum_wks, i+1, j+1, VariableNames[1], 0);}
						else if (j == 1) 	{SpectrumFillGraphLayer(NewGraphPage.Layers(j), Spectrum_wks, i+1, j+1, VariableNames[1], 0);}
	//				SpectrumFillGraphLayer(NewGraphPage.Layers(j), OptionsBook.Layers(DefaultOptionsLayersPrefix[1] + (j+1)), data_wp, OscCounts, i+1, j+1, DefaultOptionsLabel, data_wpNames);
					SpectrumFillGraphLayer(NewGraphPage.Layers(j), Spectrum_wks, i+1, j+1, VariableNames, Groups);
				}  
				bool FormatUpdErr = NewGraphPage.ApplyFormat( trTemplateFormat, true, true, true ); //changing format
				if (!FormatUpdErr) {out_str("Warning!\n GraphPage[" + i + "] ('"+NewGraphPage.GetName()+"') format update failed!\n");}
				
				//trTemplateFormat.Save("c:\Tektronix\TemlateFormat.xml" );
				
				string text = Spectrum_wks.GetName() + " - " + GraphNamePostfix;		
	//			SpectrumAutoTextToGraphLayer(NewGraphPage, GraphLayersCount-1, text);
				
				NewGraphPage.Detach();
				
			} // For all Layers in data_wp
			trTemplateFormat.Remove();
				
		}// TemplateGraph Layers count check
		else {ShowMessage("Error!\n Not enough Layers in TemplateGraphPage!\n");}
		TemplateGraphPage.Detach();
	}
	else {ShowMessage("Error!\nCan't find template graphpage '" + TemplateGraphPageName + "'.\nProcess stopped");}
}

int SpectrumSamplingSingle(Column col_data, Worksheet wks_data, Worksheet wks_out, int sampling_length = -1, int start_index = 0, int stop_index = -1)
// SINGLE column process
{
/*
Split data (one column) to samples with equal length; makes FFT of each sample; and writes output data to output worksheet

Parameters:
col_data		- [column object] column with input data
wks_data		- [worksheet object] worksheet with input data (needed for "X" column search)
wks_out 		- [worksheet object] output worksheet
sampling_length 	- [int] length of sample (the number of array elements) 			(Default = the whole input data length)
start_index 	- [int] data with index less than start_index will be ignored		(Default = 0)
stop_index 		- [int] data with index greater than start_index will be ignored	(Default = end of input data)

*/
	Dataset ds_data = col_data; // attachs the dataset to the input column
	int data_length, samples_count;
	int out_col_count = 6;

	// output column prefix list
	StringArray out_col_name_list(out_col_count);
	out_col_name_list[0] = "Freq";
	out_col_name_list[1] = "Amp";
	out_col_name_list[2] = "X";
	out_col_name_list[3] = "Line1y";
	out_col_name_list[4] = "X";
	out_col_name_list[5] = "Line2y";
	
	if (start_index < 0)	{start_index = 0;}
	
	
	if (start_index < ds_data.GetSize()-10)
	{
		// gets length of the data array
		data_length = ds_data.GetSize()-start_index;
		// if sampling_length <= 0 then process the whole data as single piece
		if ((sampling_length <= 0)||(sampling_length > data_length))	{sampling_length = data_length;}
		samples_count = data_length/sampling_length;

		// if there are too many samples asks permission to continue
		string strMsg = "Frames count is " + samples_count + "\nContinue process?";
		int nMB = MB_OKCANCEL|MB_ICONEXCLAMATION;
		if( IDCANCEL == MessageBox(GetWindow(), strMsg, "Warning!", nMB) )
		{
		    out_str("warning!\nThe process have been interrupted by user.");
		    throw 100; // breaks the whole process by throwing exception
		}


		out_str("Spectrum samples count = " + samples_count);
		for (int s=0; s<samples_count; s++)
		// for all Samples
		{
			for (int i=0; i<out_col_count; i++)
			// adding columns to output worksheet
			{
				if (wks_out.Columns.Count() < i+s*out_col_count+1)	{	wks_out.AddCol();	}
					
				wks_out.Columns(i+s*out_col_count).SetLowerBound(0);
				wks_out.Columns(i+s*out_col_count).SetUpperBound(1);
				if (i % 2 == 0)	
				{	
					//wks_out.Columns(i+s*out_col_count).SetName("X" + (i/2));
					wks_out.Columns(i+s*out_col_count).SetType(OKDATAOBJ_DESIGNATION_X);
				}
				else 	{wks_out.Columns(i+s*out_col_count).SetType(OKDATAOBJ_DESIGNATION_Y);}
					
				string name;
				name = out_col_name_list[i];
				
				// checks if the current name is used
				while (wks_out.Columns(name))
				{	
					name = StringPostfixIncrease(name);
				}
				
				wks_out.Columns(i+s*out_col_count).SetName(name);
			}
			
			Dataset xDS;
			int xDSi = GetDataXColumnIndex(wks_data, col_data.GetName());
			vector<double> vData_X, vData_Y;
			xDS.Attach(wks_data.Columns(xDSi));
			vData_X.SetSize(sampling_length);
			vData_Y.SetSize(sampling_length);
			for (i=0; i<sampling_length; i++)
			{
				vData_X[i] = xDS[start_index + i + s*sampling_length];
				vData_Y[i] = ds_data[start_index + i + s*sampling_length];
			}
			vector<double> vFreq, vAmp;
			fft_amp(vData_X, vData_Y, vFreq, vAmp, sampling_length, true, true);
			
			Dataset TempDS;
			// FREQUENCY
			TempDS.Attach(wks_out.Columns(s*out_col_count));
			TempDS.SetUpperBound(vFreq.GetSize());
			TempDS = vFreq;
			TempDS.Detach();
			
			// AMPLITUDE
			TempDS.Attach(wks_out.Columns(s*out_col_count+1));
			TempDS.SetUpperBound(vAmp.GetSize());
			TempDS = vAmp;
			TempDS.Detach();
			
			// Line_1_X
			TempDS.Attach(wks_out.Columns(s*out_col_count+2));
			TempDS[0] = xDS[start_index + s*sampling_length];
			TempDS[1] = TempDS[0];
			TempDS.Detach();
			
			double MaxY, MinY, Data_Amp;
			ds_data.GetMinMax(MinY, MaxY);
			Data_Amp = MaxY - MinY;
			//Line_1_Y
			TempDS.Attach(wks_out.Columns(s*out_col_count+3));
			TempDS[0] = MinY - Data_Amp*0,05;
			TempDS[1] = MaxY + Data_Amp*0,05;
			TempDS.Detach();
			// Line_2_X
			TempDS.Attach(wks_out.Columns(s*out_col_count+4));
			TempDS[0] = xDS[start_index + (s+1)*sampling_length - 1];
			TempDS[1] = TempDS[0];
			TempDS.Detach();
			//Line_2_Y
			TempDS.Attach(wks_out.Columns(s*out_col_count+5));
			TempDS[0] = MinY - Data_Amp*0,05;
			TempDS[1] = MaxY + Data_Amp*0,05;
			TempDS.Detach();
			
		}
		// all's Ok
		return 0;
	}
	else 
	// start_index value check
	{
		ShowMessage("Error!\n Zero-based index start_index (" + start_index + ") > or ~= Signal width (" + data_length + ").\nOr start_index (" + start_index + ") < 0.\nProcess stopped.");
		return 3;
	}
	
}


void SpectrumProcessRun8(string data_wp_name, string col_name, WorksheetPage out_wp, int sampling_length = -1, int osc_count = 1, int osc_number = 1, int start_index = 0, int stop_index = -1)
// MAIN PROCESSS 
/* 
8 input parameters:
data_wp_name	- [string] short_name  of the data WorksheetPage
col_name 		- [string] short_name of the column with data
out_wp 			- [object] output Worksheet_Page
osc_count		- [int] the number of sheets which belong to the same "shoot" and contain data from different oscilloscopes
osc_number 		- [int] the number of the sheet within one "shoot" where the target column is (index of target column = (Shoot_Number - 1) * osc_count + osc_number)
where Shoot_Number is 1-based index

start_index 	- [int] data with an index less than this one will be cut
End_Index		- [int] data with an index greater than this one will be cut
sampling_length 	- [int] data fron start_index to End_Index will be split into samples of equal length
Repeat_Count	- [int] process only first Repeat_Count "shoots" from target Worksheet_Page

*/
{
	WorksheetPage data_wp;
	Worksheet data_wks;
	int samples_count, data_length, shoots_count;
	
	int out_col_count = 6;
	StringArray out_col_name_list(out_col_count);
	out_col_name_list[0] = "Freq";
	out_col_name_list[1] = "Amp";
	out_col_name_list[2] = "X";
	out_col_name_list[3] = "Line1x";
	out_col_name_list[4] = "X";
	out_col_name_list[5] = "Line2x";
	
	// Start INDEX must be 1-based. But we need zero-based index:
	if(start_index != 0) {start_index -= 1;}
	
	data_wp = Project.WorksheetPages(data_wp_name);
	if (data_wp)
	{
		out_str("Connecting WorksheetPage '" + data_wp_name + "'.....OK");
		bool Error = false;

		// Layers and Osc count check
		if (data_wp.Layers.Count() < osc_count) 
		{
			ShowMessage("Error!\nWorksheetPage '"+data_wp_name+"' has not enouth layers ("+data_wp.Layers.Count()+") for osc_count = "+osc_count);
			Error = true;
		}

		// Layers and Osc count check
		if (data_wp.Layers.Count() % osc_count != 0) 
		{
			ShowMessage("Error!\nWorksheetPage '"+data_wp.GetName()+"' layers count("+data_wp.Layers.Count()+") is not a multiple of osc_count("+osc_count+").");
			Error = true;
		}

		// sampling_length
		if ((sampling_length < 10) && (sampling_length > 0))	// if sampling_length < 1 means 
		{
			ShowMessage("Error!\nsampling_length ["+sampling_length+"] should be at least 10.");
			Error = true;
		}

		// osc_count 
		if (osc_count <= 0) 
		{
			ShowMessage("Error!\nosc_count [" + osc_count + "] < or = 0.");	
			Error = true;
		}

		// osc_count 
		if (osc_number < 0) { osc_number = -osc_number; }
		if (osc_number > osc_count) 
		{
			ShowMessage("Error!\nosc_number [" + osc_number + "] is greater than osc_count [" + osc_count + "]. ");	
			Error = true;
		}


		
		if (!Error)
		{
			shoots_count = data_wp.Layers.Count() / osc_count;
			// Number of repeats (shoots)
			if ((Repeat_Count > shoots_count) || (Repeat_Count <= 0)) 
			{
				Repeat_Count = shoots_count;
			}

			out_str("Number of shoots = " + Repeat_Count);
			for(int s=0; s<Repeat_Count; s++)
			// for all shoot in Repeat cycle 
			{
				if(out_wp.Layers.Count() < s+1) {out_wp.AddLayer("origin");}
				Worksheet data_wks = data_wp.Layers(s*osc_count+osc_number-1);
				out_str("Layer	" + data_wks.GetName());
				out_wp.Layers(s).SetName(data_wks.GetName());
				Worksheet Out_wks = out_wp.Layers(s);
				
				if(data_wks.Columns(col_name))
				{
					// SPECTRUM
					SpectrumSamplingSingle(data_wks.Columns(col_name), data_wks, Out_wks, sampling_length, start_index, stop_index);
					out_str("Spectrum process done.");
					
					// GRAPH
					
				}
				else {out_str("Warning!\nCan't find column '" + col_name + "' at layer '" + data_wks.GetName() + "' of worksheetpage '" + data_wp.GetName() + "'.");}
			}
		}
		else {ShowMessage("Process stopped.");}
		
	}
	// Data WorksheetPage check
	else	{ShowMessage("Error!\nCan't find WorksheetPage '" + data_wp_name + "'.\nProcess stopped.");}
}

void SpectrumProcessRun(string data_wp_name, int osc_count, int osc_number, string col_name, int sampling_length, int Repeat_Count, int start_index, string out_wp_Name)
// Override
{
	if (!Project.WorksheetPages(out_wp_Name))
	{
		WorksheetPage out_wp;
		out_wp.Create("origin");
		out_wp.SetName(out_wp_Name);
		SpectrumProcessRun8(data_wp_name, osc_count, osc_number, col_name, sampling_length, Repeat_Count, start_index, out_wp);
	}
	else {ShowMessage("Error!\nWorksheetPage '" + data_wp_name + "' already exists.\nProcess stopped.");}
}

void SpectrumProcessRun7(string data_wp_name, int osc_count, int osc_number, string col_name, int sampling_length, int Repeat_Count, int start_index)
// Override
{
	WorksheetPage out_wp;
	out_wp.Create("origin");
	SpectrumProcessRun8(data_wp_name, osc_count, osc_number, col_name, sampling_length, Repeat_Count, start_index, out_wp);
}

void SpectrumProcessRun6(string data_wp_name, int osc_count, int osc_number, string col_name, int sampling_length, int Repeat_Count)
// Override
{
	WorksheetPage out_wp;
	out_wp.Create("origin");
	SpectrumProcessRun8(data_wp_name, osc_count, osc_number, col_name, sampling_length, Repeat_Count, 0, out_wp);
}

void SpectrumProcessRun5(string data_wp_name, int osc_count, int osc_number, string col_name, int sampling_length)
// Override
{
	WorksheetPage out_wp;
	out_wp.Create("origin");
	SpectrumProcessRun8(data_wp_name, osc_count, osc_number, col_name, sampling_length, -1, 0, out_wp);
}

void SpectrumProcessRun4(string data_wp_name, int osc_count, int osc_number, string col_name)
// Override
{
	WorksheetPage out_wp;
	out_wp.Create("origin");
	SpectrumProcessRun8(data_wp_name, osc_count, osc_number, col_name, -1, -1, 0, out_wp);
}
