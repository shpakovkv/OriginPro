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
#include <fft_utils.h>
#include <../OriginLab/fft.h>
////////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.


////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.

string ShortNamePostfixIncrease(string input_str)
{	
	// encreases the numerical postfix of the string by 1 or adds 2-digits postfix if no postfix found
	// cuts the string at the end if the length of string + index exceeded 17

	int current_index, index_digits, str_length;
	string out_str;
	
	index_digits = 0;	// initial value
	
	// gets index (postfix) length
	while ((NANUM != atof(input_str.Mid(input_str.GetLength()-index_digits-1))) && (index_digits < 15))		
	{
		index_digits++;	// searchs from the end of the string and increase length value if current character is numeric
	}
	
	if (index_digits==0) 						// if no postfix found
	{	
		str_length = input_str.GetLength();					// gets the input string length
		if (str_length > 15)	{	str_length = 15;	}	// if the input string length exceeds 15 cuts the string
		current_index=2;									// if no postfix found adds 2-digits postfix
	}					
	else 	
	{	
		current_index = atoi(input_str.Mid(input_str.GetLength() - index_digits));	// gets existing postfix value
		current_index++;											// increase index by 1
		str_length = input_str.GetLength() - index_digits;					// gets string length (without index)
	}
	
	out_str = input_str.Mid(0,str_length) + current_index;					// concatenate string and index
	if ((out_str.GetLength() > 17) && (str_length < 15)) 			// checks if the length of the index increased
	{
		out_str = input_str.Mid(0,str_length - 1) + current_index;	// corrects string length if it exceeds 17
	}
	
	return out_str;
}


void ShowMessage(string msg, string title = "Information")
{
	out_str(msg);										// adds the message to the command window log
	// LT_execute("type -b " + msg);					// shows the message box (LabTalk script execution)
	int enum_MB = MB_OK|MB_ICONINFORMATION;				// sets the message box type
	MessageBox(GetWindow(), msg, title, enum_MB)		// shows the message box
}


int GetNearestIndex(Dataset ds, double value)
{	
	int start_i = ds.GetLowerBound();
	int stop_i = ds.GetUpperBound();
	if ((value < ds[start_i]) || (value > ds[stop_i]) || (stop_i - start_i < 3))
	{
		return NANUM;
	}

	for (int i = start_i + 1; i <= stop_i; i++)
	{
		if ((value > ds[i-1]) && (value <= ds[i]))
		{
			return i;
		}
	}
	return NANUM;
}

int GetColumnIndex (Worksheet wks, string col_name)
{	// finds column of worksheet by name and return it's index
	for (int i=0; i<wks.Columns.Count(); i++)
	{
		if (wks.Columns(i).GetName() == col_name) { return i; }
	}
	return -1;
	// return "-1" if worksheet don't contain column with this name
}



  //=======================================================================================================================
 //		GET X-type column INDEX for input Y-type column
//======================================

int GetDataXColumnIndex(Worksheet wks, string col_name)
{// finding column with 'X'-type data for column 'col_name' containing 'Y'-type data in order to plot graph or something...
	int index;
	index = GetColumnIndex(wks, col_name);
	
	if ((index == -1) || (wks.Columns(index).GetType() == OKDATAOBJ_DESIGNATION_X))
	{
		return -1;	// -1 means no column with "X"-type data found or input column contains "X"-type data
	}
	else
	{
		for (int i = index; i >= 0; i--)
		{
			if (wks.Columns(i).GetType() == OKDATAOBJ_DESIGNATION_X) 
			{ return i; }	// return column index
		}
		return -1;			// -1 means no column with "X"-type data found
	}
}


  //=======================================================================================================================
 //		SINGLE COLUMN PROCESS
//======================================
void SingleSpectrumProcess(Dataset ds_data, Worksheet wks_data, Worksheet wks_out, int sampling_width, int samples_count, int start_index, int stop_index)
{
	/*
	Split data (one column) to samples with equal length; makes FFT of each sample; and writes output data to output worksheet

	Parameters:
	col_data		- [column object] column with input data
	wks_data		- [worksheet object] worksheet with input data (needed for "X" column search)
	wks_out 		- [worksheet object] output worksheet
	sampling_width 	- [int] length of sample (the number of array elements) 			(Default = the whole input data length)
	start_index 	- [int] data with index less than start_index will be ignored		(Default = 0)
	stop_index 		- [int] data with index greater than start_index will be ignored	(Default = end of input data)
	*/

	// creates output layer column name list
	int out_col_count = 6;
	StringArray out_col_name_list(out_col_count);
	out_col_name_list[0] = "Freq";					// frequency column "X"-type
	out_col_name_list[1] = "Amp";					// Amplitude column "Y"-type
	out_col_name_list[2] = "X";						// x-coordinates for LEFT vertical line (represent left bound of the sample)
	out_col_name_list[3] = "Line1y";				// y-coordinates of the top and bottom points of the LEFT vertical line
	out_col_name_list[4] = "X";						// x-coordinates for RIGHT vertical line (represent right bound of the sample)
	out_col_name_list[5] = "Line2y";				// y-coordinates of the top and bottom points of the RIGHT vertical line
	
	for (int sample_number = 0; sample_number < samples_count; sample_number++)
		// for all Samples
	{
		if (sample_number == samples_count - 1)
		{

		}

		for (int i = 0; i < out_col_count; i++)
		// adding columns to output worksheet
		{
			if (wks_out.Columns.Count() < i+sample_number * out_col_count+1)	{	wks_out.AddCol();	}
				
			wks_out.Columns(i + sample_number*out_col_count).SetLowerBound(0);
			wks_out.Columns(i + sample_number*out_col_count).SetUpperBound(1);
			if (i % 2 == 0)	
			{	
				wks_out.Columns(i + sample_number * out_col_count).SetType(OKDATAOBJ_DESIGNATION_X);		// set column type to "X"
			}
			else 	{wks_out.Columns(i + sample_number * out_col_count).SetType(OKDATAOBJ_DESIGNATION_Y);}	// set column type to "Y"
				
			string name = out_col_name_list[i];					// current column name
			while (wks_out.Columns(name))						// checks if the current column name is used
			{	
				name = ShortNamePostfixIncrease(name);				// change current column name postfix
			}
			
			wks_out.Columns(i + sample_number*out_col_count).SetName(name);	// set current column name
		}
		
		Dataset x_data_ds;
		int x_data_index = GetDataXColumnIndex(wks_data, ds_data.GetName());
		vector<double> v_data_x, v_data_y;
		x_data_ds.Attach(wks_data.Columns(x_data_index));
		v_data_x.SetSize(sampling_width);
		v_data_y.SetSize(sampling_width);

		for (i=0; i<sampling_width; i++)
		{
			v_data_x[i] = x_data_ds[start_index + i + sample_number * sampling_width];
			v_data_y[i] = ds_data[start_index + i + sample_number * sampling_width];
		}
		vector<double> vFreq, vAmp;
		fft_amp(v_data_x, v_data_y, vFreq, vAmp, sampling_width, true, true);
		
		Dataset TempDS;
		// FREQUENCY
		TempDS.Attach(wks_out.Columns(sample_number * out_col_count));
		TempDS.SetUpperBound(vFreq.GetSize());
		TempDS = vFreq;
		TempDS.Detach();
		
		// AMPLITUDE
		TempDS.Attach(wks_out.Columns(sample_number * out_col_count + 1));
		TempDS.SetUpperBound(vAmp.GetSize());
		TempDS = vAmp;
		TempDS.Detach();
		
		// Line_1_X
		TempDS.Attach(wks_out.Columns(sample_number * out_col_count + 2));
		TempDS[0] = x_data_ds[start_index + sample_number * sampling_width];
		TempDS[1] = TempDS[0];
		TempDS.Detach();
		
		// gets max and min values of Y
		double MaxY, MinY, data_amp_range;
		ds_data.GetMinMax(MinY, MaxY);
		data_amp_range = MaxY - MinY;
		//Line_1_Y
		TempDS.Attach(wks_out.Columns(sample_number * out_col_count + 3));
		TempDS[0] = MinY - data_amp_range * 0,05;		// sets vertical lines lower points to 5% lower than minimum 
		TempDS[1] = MaxY + data_amp_range * 0,05;		// sets vertical lines upper points to 5% higher than maximum
		TempDS.Detach();
		// Line_2_X
		TempDS.Attach(wks_out.Columns(sample_number * out_col_count + 4));
		TempDS[0] = x_data_ds[start_index + (sample_number + 1) * sampling_width - 1];
		TempDS[1] = TempDS[0];
		TempDS.Detach();
		//Line_2_Y
		TempDS.Attach(wks_out.Columns(sample_number * out_col_count + 5));
		TempDS[0] = MinY - data_amp_range * 0,05;		// sets vertical lines lower points to 5% lower than minimum 
		TempDS[1] = MaxY + data_amp_range * 0,05;		// sets vertical lines upper points to 5% higher than maximum
		TempDS.Detach();
	}
}


  //=======================================================================================================================
 //		MAIN 
//======================================
void SpectrumProcessRun(string data_wp_name, string col_name, string out_wp_name, int sampling_width = -1, int osc_count = 1, int osc_number = 1, int start_index = 0, int stop_index = -1)
{
	/* 
	3-8 input parameters:
	data_wp_name	- [string] short_name  of the data WorksheetPage
	col_name 		- [string] short_name of the column with data
	out_wp_name 	- [string] new output Worksheet_Page name 

	sampling_width  [Default = -1] - [int] data from start_index to stop_index will be split into samples of equal length
	osc_count		[Default =  1] - [int] the number of sheets which belong to the same "shoot" and contain data from different oscilloscopes		
	osc_number 		[Default =  1] - [int] the number of the sheet within one "shoot" where the target column is 
					  			   (index of target column = (Shoot_Number - 1) * osc_count + osc_number) where Shoot_Number is 1-based index 

	start_index 	[Default =  0] - [int] data with an zero-based index less than this will be cut 	
	stop_index		[Default = -1] - [int] data with an zero-based index greater than this will be cut 									
	*/

	WorksheetPage data_wp;							// source WorksheetPage
	WorksheetPage out_wp;							// output WorksheetPage
	Worksheet data_wks;								// data WorksheetPage 
	Dataset ds_data;								// data column dataset
	int samples_count, data_length, shoots_count;	// other variables
	bool forced_by_user = false;

	// creates output layer column name list
	int out_col_count = 6;
	StringArray out_col_name_list(out_col_count);
	out_col_name_list[0] = "Freq";
	out_col_name_list[1] = "Amp";
	out_col_name_list[2] = "X";
	out_col_name_list[3] = "Line1x";
	out_col_name_list[4] = "X";
	out_col_name_list[5] = "Line2x";
	
	// Start INDEX must be 1-based. But we need zero-based index:
	// if(start_index != 0) {start_index -= 1;}  // ARCHIVED
	
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

		// sampling_width check
		if ((sampling_width < 10) && (sampling_width > 0))	// if sampling_width < 1 means 
		{
			ShowMessage("Error!\nsampling_width ["+sampling_width+"] should be at least 10.");
			Error = true;
		}

		// osc_count check
		if (osc_count <= 0) 
		{
			ShowMessage("Error!\nosc_count [" + osc_count + "] < or = 0.");	
			Error = true;
		}

		// osc_number check
		if (osc_number < 0) { osc_number = -osc_number; }
		if (osc_number > osc_count) 
		{
			ShowMessage("Error!\nosc_number [" + osc_number + "] is greater than osc_count [" + osc_count + "]. ");	
			Error = true;
		}

		// start_index check
		if (start_index < 0)
		{
			out_str("Warning! Start index is " + start_index + ". Set to 0.");
			start_index = 0;
		}

		// check if out_wp_name is used
		if (!Project.WorksheetPages(out_wp_name))
		{
			out_wp.Create("origin");
			out_wp.SetName(out_wp_name);
		}
		else 
		{
			ShowMessage("Error!\nWorksheetPage '" + out_wp_name + "' already exists.\nProcess stopped.");
			Error = true;
		}

		
		if (!Error)
		{
			// gets the number of shoots in WorksheetPage
			shoots_count = data_wp.Layers.Count() / osc_count;
			out_str("Number of shoots = " + shoots_count);

			for(int shoot_num = 0; shoot_num < shoots_count; shoot_num++)
			// for all shoot in WorksheetPage
			{
				if(out_wp.Layers.Count() < shoot_num + 1) {out_wp.AddLayer("origin");}			// adds new layer to the output WorksheetPage
				Worksheet data_wks = data_wp.Layers(shoot_num * osc_count + osc_number - 1);	// attach current source layer
				out_wp.Layers(shoot_num).SetName(data_wks.GetName());							// set current output layer name equal to source layer name
				out_str("Layer	" + data_wks.GetName());										// prints current layer name to log
				Worksheet out_wks = out_wp.Layers(shoot_num);									// attach current output layer
				
				if(data_wks.Columns(col_name))
				{
					ds_data.Attach(data_wks.Columns(col_name));
					// sample length must be greater or equal 10 for FFT process
					if (start_index < ds_data.GetSize()-10)
					{
						int local_stop_index = stop_index;								// stop_index of current dataset
						if (local_stop_index > ds_data.GetSize())
						{
							local_stop_index = ds_data.GetSize() - 1;					// correct current stop index if needed
						}

						data_length = local_stop_index - start_index + 1;				// gets length of the data array
						if ((sampling_width <= 0) || (sampling_width > data_length))	// if sampling_width <= 0 then process the whole data as single sample
						{
							sampling_width = data_length;
						}

						samples_count = data_length / sampling_width;		// gets samples count

						if (!forced_by_user)								// if there are too many samples asks permission to continue (only once)
						{	
							string strMsg = "First shoot's samples count is " + samples_count + ".\nShoots count is " + shoots_count + ".\nContinue process?";
							int nMB = MB_OKCANCEL|MB_ICONQUESTION;
							if( IDCANCEL == MessageBox(GetWindow(), strMsg, "Confirmation", nMB) )
							{
							    out_str("Warning!\nThe process have been interrupted by user.");
							    throw 100; // breaks the whole process by throwing exception 
							}
							else 
							{ 
								forced_by_user = true; 
							}
						}

						// SPECTRUM
						SingleSpectrumProcess(ds_data, data_wks, out_wks, sampling_width, samples_count, start_index, stop_index);
						out_str("Spectrum process done.");
						
						// GRAPH // *reserved*

					}
					else 
					{// start_index value check
						out_str("Warning! Zero-based index start_index (" + start_index + ") > or ~= sample length (" + data_length + "). Skip current layer.");
					}
					ds_data.Detach();
				}
				else {out_str("Warning!\nCan't find column '" + col_name + "' at layer '" + data_wks.GetName() + "' of worksheetpage '" + data_wp.GetName() + "'.");}
			}
		}
		else {ShowMessage("Process stopped.");}
		
	}
	// Data WorksheetPage check
	else	{ShowMessage("Error!\nCan't find WorksheetPage '" + data_wp_name + "'.\nProcess stopped.");}
}



//=======================================================================================================================
 //		MAIN SINGLE
//======================================
void SpectrumProcessSingleRun(string data_wp_name, string data_wks_name, string col_name, string out_wp_name, double start_x, double stop_x)
{
	/* 
	3-8 input parameters:
	data_wp_name	- [string] short_name  of the data WorksheetPage
	col_name 		- [string] short_name of the column with data
	out_wp_name 	- [string] new output Worksheet_Page name 

	sampling_width  [Default = -1] - [int] data from start_index to stop_index will be split into samples of equal length
	osc_count		[Default =  1] - [int] the number of sheets which belong to the same "shoot" and contain data from different oscilloscopes		
	osc_number 		[Default =  1] - [int] the number of the sheet within one "shoot" where the target column is 
					  			   (index of target column = (Shoot_Number - 1) * osc_count + osc_number) where Shoot_Number is 1-based index 

	start_index 	[Default =  0] - [int] data with an zero-based index less than this will be cut 	
	stop_index		[Default = -1] - [int] data with an zero-based index greater than this will be cut 									
	*/

	WorksheetPage data_wp;							// source WorksheetPage
	WorksheetPage out_wp;							// output WorksheetPage
	Worksheet data_wks;								// data WorksheetPage 
	Worksheet out_wks;
	string out_wks_name;
	Dataset ds_data;								// data column dataset
	Dataset data_y_ds;
	Dataset data_x_ds;
	int samples_count, data_length, start_i, stop_i;					// other variables
	int sampling_width;

	// creates output layer column name list
	int out_col_count = 6;
	StringArray out_col_name_list(out_col_count);
	out_col_name_list[0] = "Freq";
	out_col_name_list[1] = "Amp";
	out_col_name_list[2] = "X";
	out_col_name_list[3] = "Line1x";
	out_col_name_list[4] = "X";
	out_col_name_list[5] = "Line2x";
	
	// Start INDEX must be 1-based. But we need zero-based index:
	// if(start_index != 0) {start_index -= 1;}  // ARCHIVED
	
	data_wp = Project.WorksheetPages(data_wp_name);
	if (data_wp)
	{
		out_str("Connecting WorksheetPage '" + data_wp_name + "'.....OK");

		// wks existence check
		data_wks = data_wp.Layers(data_wks_name);
		if (!data_wks) 
		{

			ShowMessage("Error!\nWorksheetPage '" + data_wp_name + "' has not layer '" + data_wks_name + "'");
			throw 100;
		}

		/*
		// sampling_width check
		if ((sampling_width < 10) && (sampling_width > 0))	// if sampling_width < 1 means 
		{
			ShowMessage("Error!\nsampling_width [" + sampling_width + "] should be at least 10.");
			throw 100;
		}
		*/

		// column check
		
		if (data_wks.Columns(col_name))
		{
			data_y_ds.Attach(data_wks.Columns(col_name));
		}
		else
		{
			ShowMessage("Error!\nWorksheetPage '" + data_wp_name + "' has not layer '" + data_wks_name + "'");
			throw 100;
		}

		int data_x_col_index = GetDataXColumnIndex(data_wks, col_name);
		if (data_x_col_index >= 0)
		{
			data_x_ds.Attach(data_wks.Columns(data_x_col_index));
		}
		else
		{
			ShowMessage("Error! Can't find column with 'X'-data for 'Y'-data in '" + col_name + "' column.");
			throw 100;
		}

		// start_x check
		double min, max;
		data_x_ds.GetMinMax(min, max);
		if ((!start_x) || (start_x < min) || (start_x >= stop_x))
		{
			start_x = min;
		}

		// stop_x check
		if ((!stop_x) || (stop_x > max) || (start_x >= stop_x))
		{
			stop_x = max;
		}

		// gets start and stop indexes for X data column from start and stop values
		start_i = GetNearestIndex(data_x_ds, start_x);
		stop_i = GetNearestIndex(data_x_ds, stop_x);
		//data_x_ds.Detach();

		if ((!start_i) || (!stop_i))
		{
			ShowMessage("Error! Wrong start/stop condition.");
			throw 100;
		}

		if (stop_i - start_x < 10)
		{
			ShowMessage("Not enough data for FFT.");
			throw 100;
		}

		// check if out_wp_name exists
		if (Project.WorksheetPages(out_wp_name))
		{
			out_wp = Project.WorksheetPages(out_wp_name);				// attach output worksheetpage
			out_wp.AddLayer("origin");									// add new layer to output worksheetpage
			out_wks = out_wp.Layers(out_wp.Layers.Count() - 1);			// attach output worksheet
			out_wks_name = data_wks_name;								// gets data worksheet name
			while (out_wp.Layers(out_wks_name))							// checks if current name is used
			{	
				out_wks_name = ShortNamePostfixIncrease(out_wks_name);	// change current name postfix
			}
			
			out_wks.SetName(out_wks_name);								// change output worksheet name 
		}
		else 
		{
			ShowMessage("Error!\nCan't find WorksheetPage '" + out_wp_name + "'.\nProcess stopped.");
			throw 100;
		}

							
		sampling_width = stop_i - start_i + 1;		// gets length of the data array
		samples_count = 1;							// gets samples count

		// SPECTRUM
		SingleSpectrumProcess(data_y_ds, data_wks, out_wks, sampling_width, samples_count, start_i, stop_i);
		out_str("Spectrum process done.");
		out_str("sampling_width = " + sampling_width);
		out_str("start_i = " + start_i);
		out_str("stop_i = " + stop_i);
		
		// GRAPH // *reserved*

		ds_data.Detach();
		
	}
	// Data WorksheetPage check
	else	{ShowMessage("Error!\nCan't find WorksheetPage '" + data_wp_name + "'.\nProcess stopped.");}
}




//=========================================================================================================================
//  --------------        UNTESTED    FUNCTIONS       ---------------------------------------------------------------------
//=========================================================================================================================


  //=======================================================================================================================
 //		GRAPH MAIN
//======================================
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


  //====================================
 //		GRAPH subprocess
//======================================

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


  //====================================
 //		GRAPH subprocess
//======================================
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
