/*------------------------------------------------------------------------------*
 * File Name:	ProcessSignals v2.0												*
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


////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.
//======================================
void ShowMessage(string m)
{
	out_str(m);
	LT_execute("type -b " + m);
}

//----------------------------------------------------------------------------------//|
//function for convertion from string AxisType to integer AxisType					//|
//return -1 for incompatible input string											//|
int AxisTypeToInt(string s)															//|
{																					//|
	int out = -1;																	//|
	if ((s == "X") || (s == "x")) {		out = 3;	}								//|
	if ((s == "Y") || (s == "y")) {		out = 0;	}								//|
	if ((s == "Z") || (s == "z")) {		out = 5;	}								//|
	if ((s == "Label") || (s == "label") || (s == "LABEL")) {		out = 4;	}	//|
	return out;																		//|
}																					//|
//----------------------------------------------------------------------------------//|

//================================================================================================
void FillColumnsData(Worksheet InputDataWks, Worksheet OutputDataWks, Worksheet OptionsWks, StringArray DefaultLabelsName, int ShootNum)
{ //filling worksheet with data
	
	Dataset InputDataDS, OutputDataDS, MultiplierDS, DelayDS, TimeCorrDS;
	StringArray type, name, longname, units, comments;
	int ColCount = InputDataWks.Columns.Count();
	double TimeCorr;
	
	//getting string arrays [0..Names.Length]
	OptionsWks.Columns(DefaultLabelsName[0]).GetStringArray(name, 0, ColCount); // get array of Names of columns
	OptionsWks.Columns(DefaultLabelsName[1]).GetStringArray(longname, 0, ColCount);
	OptionsWks.Columns(DefaultLabelsName[2]).GetStringArray(units, 0, ColCount);
	OptionsWks.Columns(DefaultLabelsName[3]).GetStringArray(comments, 0, ColCount);
	OptionsWks.Columns(DefaultLabelsName[6]).GetStringArray(type, 0, ColCount);
	
	MultiplierDS.Attach(OptionsWks.Columns(DefaultLabelsName[4])); 	// get access to "Multiplier" column
	DelayDS.Attach(OptionsWks.Columns(DefaultLabelsName[5])); 		// get access to 'Delay' column
	TimeCorrDS.Attach(OptionsWks.Columns(DefaultLabelsName[7]));		// get access to 'TimeCorr' column
	TimeCorr = TimeCorrDS[ShootNum];
	TimeCorrDS.Detach();
		
	
	//deleting existing columns in Output worksheet if needed
	if (OutputDataWks.Columns.Count() > 0)
	{
		for (int i=OutputDataWks.Columns.Count(); i>0; i--)
		{
			OutputDataWks.DeleteCol(i-1);
		}
	}
	
  //filling columns with formated data
	//for all columns in InputWorksheet but no more than we need to rename and format
	for(int i=0; i<ColCount; i++)
	{
		OutputDataWks.AddCol();

		if (name[i]!="Delete") // working with only Data that we realy need
		{
			double multiplier = 1;
			double delay = 0;
			
			if (name[i]!="")
			{
				//setting new column format
				OutputDataWks.Columns(i).SetName(name[i]);
    			OutputDataWks.Columns(i).SetLongName(longname[i]);
    			OutputDataWks.Columns(i).SetUnits(units[i]);
    			OutputDataWks.Columns(i).SetComments(comments[i]);
    			OutputDataWks.Columns(i).SetType(AxisTypeToInt(type[i]));
    			multiplier = MultiplierDS[i];
    			delay = DelayDS[i];
    			if ((type[i] == "x") || (type[i] == "X")) {delay = delay + TimeCorr;}
			}
			else 
			{
				//copying old column format
				OutputDataWks.Columns(i).SetName(InputDataWks.Columns(i).GetName());
    			OutputDataWks.Columns(i).SetLongName(InputDataWks.Columns(i).GetLongName());
    			OutputDataWks.Columns(i).SetUnits(InputDataWks.Columns(i).GetUnits());
    			OutputDataWks.Columns(i).SetComments(InputDataWks.Columns(i).GetComments());
				OutputDataWks.Columns(i).SetType(InputDataWks.Columns(i).GetType());
			}
				
			InputDataDS.Attach(InputDataWks.Columns(i)); // get access to next "DataColumn"
			OutputDataDS.Attach(OutputDataWks.Columns(i)); // get access to next "TargetColumn" to write in
			OutputDataDS.SetSize(InputDataDS.GetSize()); // settting size of column
			
			
			
			//for all cells in active column of InputWorksheet
			for (int j=0; j<InputDataDS.GetSize(); j++)
			{
				OutputDataDS[j] = InputDataDS[j] * multiplier - delay;  	//multiplying values if needed
				
			}
			InputDataDS.Detach(); // clear variables value
			OutputDataDS.Detach(); 
		}
		
	}
	MultiplierDS.Detach(); // free some memory
	DelayDS.Detach();
	
	//deleting unnecessary columns
	for (i=name.GetSize()-1; i>=0; i--)
	{
		if (name[i]=="Delete")
		{
			OutputDataWks.DeleteCol(i);
		}
	}
	
	// set name of our new Worksheet
	OutputDataWks.SetName(InputDataWks.GetName()); 
}

//===========================================================================
void ProcessSignals()
{
  //variables
	Dataset TempDS;
	Worksheet GlobalOptionsWks, TempWks;
	string OutputDataWPName, InputDataWPName;
	int OscCount;
	
  //constants
	string OptionsWPName = "DOptions";
	StringArray DefaultLabelsName(8), DefaultGlobalLabelsName(3), DefaultDOptionsLayersPrefix(2);
	DefaultLabelsName[0] = "ColName";
	DefaultLabelsName[1] = "LongName";
	DefaultLabelsName[2] = "Units";
	DefaultLabelsName[3] = "Comments";
	DefaultLabelsName[4] = "Multiplier";
	DefaultLabelsName[5] = "Delay";
	DefaultLabelsName[6] = "Type";
	DefaultLabelsName[7] = "TimeCorr";
	DefaultGlobalLabelsName[0] = "WksName";
	DefaultGlobalLabelsName[1] = "NewWksName";
	DefaultGlobalLabelsName[2] = "OscCount";
	DefaultDOptionsLayersPrefix[0] = "Options";
	DefaultDOptionsLayersPrefix[1] = "Oscillograph_";
	//------------------------------------------
    // WorksheetPage with options access check
	WorksheetPage OptionsWP(OptionsWPName);
	if (OptionsWP)
	{
		//get direct access to Worksheet with global options
		GlobalOptionsWks = OptionsWP.Layers(DefaultDOptionsLayersPrefix[0]);
		
		//column 'WksName' existence check
		if (GlobalOptionsWks.Columns(DefaultGlobalLabelsName[0]))
    	{
    		//accessing to column 'WksName' of Options WorksheetPage Layer 'Options'
    		TempDS.Attach(GlobalOptionsWks.Columns(DefaultGlobalLabelsName[0]));
    		TempDS.GetText(0, InputDataWPName); //getting the name of WorksheetPage with input data
    		TempDS.Detach();
    		
    		//column 'NewWksName' existence check
	    	if (GlobalOptionsWks.Columns(DefaultGlobalLabelsName[1]))
	    	{
	    		//accessing to column 'NewWksName' of Options WorksheetPage Layer 'Options'
	    		TempDS.Attach(GlobalOptionsWks.Columns(DefaultGlobalLabelsName[1]));
	    		TempDS.GetText(0, OutputDataWPName); //getting the name of new WorksheetPage to create and write output data in
	    		TempDS.Detach();
	    		
	//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	//SERCHING FOR ERRORS
	    		bool accessError = false;
	    		
	    	//input data WorksheetPage access check
	    		if (!Project.WorksheetPages(InputDataWPName))
	    		{	ShowMessage("Error!\nCan't connect to '"+InputDataWPName+"' WorksheetPage with input data! \n(Wrong WksName)\n"); 
	    			accessError = true;		}
	    			
	//'OscCount' column access and value check
	    		//column existence check
	    		if ((!accessError) && (GlobalOptionsWks.Columns(DefaultGlobalLabelsName[2])))
	    		{	
	    			//get access to column through Dataset
	    			TempDS.Attach(GlobalOptionsWks.Columns(DefaultGlobalLabelsName[2]));
	    			//get value
	    			OscCount = TempDS[0]; //for NANUM value OscCount will be 0
	    			TempDS.Detach();
	    			
	    		//'OscCount' column value compatibility check
	    			//value is numeric check
	    			if (OscCount < 1)
	    			{   ShowMessage("Error!\n'"+DefaultGlobalLabelsName[2]+"' at layer '"+DefaultDOptionsLayersPrefix[0]+"' of '"+OptionsWPName+"' has incompatible value.\n");
	    				accessError = true;		} 
	    			else
	    			{
	    				// multiplicity check of OscCOunt and InputWorksheetpage.Layers.Count
	    				if (Project.WorksheetPages(InputDataWPName).Layers.Count() % OscCount > 0) 
	    				{	ShowMessage("Warning!\nWorksheetPage '"+InputDataWPName+"' layer-s count("+Project.WorksheetPages(InputDataWPName).Layers.Count()+") is not a multiple of OscCount("+OscCount+").\nSome data won't be processed!\n");	}
	    				//layers with Osc_ options count sufficiency check
	    				if (OscCount >= OptionsWP.Layers.Count())
	    				{	ShowMessage("Error!\nNot enaugh layers("+(OptionsWP.Layers.Count()-1)+") with '"+DefaultDOptionsLayersPrefix[1]+"' options at '"+OptionsWPName+"' WorksheetPage for that OscCount value("+OscCount+").\n");
	    					accessError = true;		}
	    			}
	    		}
	    		else if (!accessError)
	    			{	ShowMessage("Error! Column 'OscCount' at layer 'Options' of 'DOptions' WorksheetPage is missign!\n"); 
	    				accessError = true;		}
	    		GlobalOptionsWks.Detach();
	    				
	//NewWksName value check
	    		//value is empty check
	    		if (OutputDataWPName == "")
	    		{	ShowMessage("Error!\n'NewWksName' at layer '"+DefaultDOptionsLayersPrefix[0]+"' of '"+OptionsWPName+"' has wrong value.\n(Value is empty)\n");
	    			accessError = true;		}
	    		//WorksheetPage.Name is unique check
	    		if (Project.WorksheetPages(OutputDataWPName))
	    		{	ShowMessage("Error!\n'NewWksName' at layer '"+DefaultDOptionsLayersPrefix[0]+"' of '"+OptionsWPName+"' has wrong value.\n Worksheetpage '"+OutputDataWPName+"' already exist!\n");
	    			accessError = true;		}
	    					
	//'options' columns existence check
	    		if (!accessError)
	    		{
	    			//for all oscillogrsphs
	    			for (int i=1; i<=OscCount; i++)
	    			{
	    				Worksheet wks;
	    				//layer existence check
	    				if (OptionsWP.Layers(DefaultDOptionsLayersPrefix[1] + i))
	    				{
	    					//get access to layer
	    					wks = OptionsWP.Layers(DefaultDOptionsLayersPrefix[1] + i);
	    					//for all possible columns name
	    					for (int j=0; j<DefaultLabelsName.GetSize(); j++)
	    					{
	    						//column existence check
	    						if (!wks.Columns(DefaultLabelsName[j])) 
	    						{
	    							ShowMessage("Error!\nColumn '"+DefaultLabelsName[j]+"' is missing.\nAt layer '"+DefaultDOptionsLayersPrefix[1]+i+"' of '"+OptionsWPName+"' WorksheetPage.\n");
	    							accessError = true;
	    						}
	    					}
	    				}
	    				else 
	    				{	ShowMessage("Error!\nLayer '"+DefaultDOptionsLayersPrefix[1] + i+"' is missign!\n(At '"+OptionsWPName+"' WorksheetPage)\n");
	    					accessError = true;		}
	    			}
	    		}
	    		

	//Multiplier, Delay and TimeCorr values check and replace if needed
	    		if (!accessError)
	    		{
	    			// for all 'Osc' layers of Options WorksheetPage
	    			for (int i=1; i<=OscCount; i++)
	    			{
	    				StringArray sr;
	    				Worksheet wks;
	    				wks = OptionsWP.Layers(DefaultDOptionsLayersPrefix[1] + i); //active 'Osc' layer
	    				
	    				
	    				wks.Columns(DefaultLabelsName[7]).SetLowerBound(0);
	    				int TmpShootsCount = Project.WorksheetPages(InputDataWPName).Layers.Count() / OscCount - 1;
						wks.Columns(DefaultLabelsName[7]).SetUpperBound(TmpShootsCount);
	    				
	    				StringArray tmpsa;
	    				wks.Columns(DefaultLabelsName[7]).GetStringArray(tmpsa,0,wks.Columns(DefaultLabelsName[0]).GetUpperBound());
	    				
	    				int dsfsag;
	    				//for Multiplier (m=4) and Delay (m=5) columns
	    				int UpperBound = wks.Columns(DefaultLabelsName[0]).GetUpperBound();
	    				for (int m=4; m<=6; m++)
	    				{
	    					if (m==6) {m=7; UpperBound = TmpShootsCount; }
	    					
	    					wks.Columns(DefaultLabelsName[m]).GetStringArray(sr, 0, UpperBound);
	    					//for all values in column
	    					for (int j=0; j<=UpperBound; j++)
	    					{
	    						//filling empty cells to except access errors
	    						// (accessing to empty cell through Dataset[i] leads to error)
	    						if (sr[j] == "") 
	    						{	
	    							StringArray ReplaceSR(1);
	    							ReplaceSR[0] = "s";
	    							wks.Columns(DefaultLabelsName[m]).PutStringArray(ReplaceSR, j);  
	    						}
	    					}
	    					
	    					//pissible access errors prevented above
	    					//now getting direct access to cells and replacing NANUM values with Numeric
	    					Dataset ds;
	    					ds.Attach(wks.Columns(DefaultLabelsName[m]));
	    					
	    					//for all values in column
	    					for (j=0; j<=UpperBound; j++)
	    					{
	    						//replacing nonnumeric values
	    						if (ds[j] == NANUM) 
	    						{
	    							//replace string data with numeric '1' for Multiplauer and '0' for Delay
	    							if (m==4) 
	    							{// Multiplier
	    								ds[j] = 1;
	    								out_str("Warning!\nMultiplayer value at Layer '"+wks.GetName()+"' of 'Options' WorksheetPage has incompatible type.\nValue replased with '1'.\n");
	    							}
	    							else 
	    							{// Delay and TimeCorr
	    								ds[j] = 0;
	    								out_str("Warning!\nDelay value at Layer '"+wks.GetName()+"' of 'Options' WorksheetPage has incompatible type.\nValue replased with '0'.\n");
	    							}
	    						}
	    					}
	    					ds.Detach(); 
	    				}
	    				
	    			}
	    		}
	    		
	//'Type' values check
	    		if (!accessError)
	    		{
	    			// for all 'Osc_' layers of Options WorksheetPage
	    			for (int i=1; i<=OscCount; i++)
	    			{
	    				StringArray sr;
	    				Worksheet wks;
	    				wks = OptionsWP.Layers(DefaultDOptionsLayersPrefix[1] + i); //active 'Osc' layer
	    				wks.Columns(DefaultLabelsName[6]).GetStringArray(sr, 0, wks.Columns(DefaultLabelsName[0]).GetUpperBound()); //get string array 'Type'
	    				for (int j=0; j<wks.Columns(DefaultLabelsName[0]).GetUpperBound(); j++)
	    				{
	    					//replacing incompatible values of StringArray with "Y"
	    					if (AxisTypeToInt(sr[j]) == -1)
	    					{
	    						sr[j] = "Y";
	    						out_str("Warning!\nValue #"+(j+1)+" of column '"+DefaultLabelsName[6]+"' has incompatible value. (At Layer '"+DefaultDOptionsLayersPrefix[1]+i+"' of '"+OptionsWPName+"' WorksheetPage)\nValue changed to 'Y'.");
	    					}
	    				}
	    				//replacing Type values with new StringArray
	    				wks.Columns(DefaultLabelsName[6]).PutStringArray(sr);
	    			}
	    		}
	//end of error search
	//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

	//----------------------------------------------------------------------------------
	// Process Begins
	    		if (!accessError)
	    		{
	    		//attaching data objects
	    			WorksheetPage InputDataWP(InputDataWPName);
	    			WorksheetPage OutputDataWP();
	    			OutputDataWP.Create("origin");
	    			//OutputDataWP.Rename(OutputDataWPName); //WorksheetPage.Rename(string s) - unstable function
	    			OutputDataWP.SetName(OutputDataWPName);
	    			OutputDataWP.Layers(0).Delete(); //creating new clear worksheetpage wothout any layers
	    			
	    			
	    			
	    		//adding enough layers to new WorksheetPage
	    			for (int i=0; i<=InputDataWP.Layers.Count()-OscCount; i+=OscCount)
	    			{
	    				for (int j=0; j<OscCount; j++)
	    				{
	    					OutputDataWP.AddLayer();
	    				}
	    			}
	    			//now OutputWP.Layers.Count == |InputWP.Layers.Count / OscCOunt| * OscCount
	    		
	    		//filling Output worksheetpage with data and formatting it
	    			//for all oscillograph
	    			for (int osc=1; osc<=OscCount; osc++)
	    			{
	    				//for all layers with oscillograph ¹ osc
	    				for (int i=osc-1; i<InputDataWP.Layers.Count(); i+=OscCount)
	    				{
	    					FillColumnsData(InputDataWP.Layers(i), OutputDataWP.Layers(i), OptionsWP.Layers(DefaultDOptionsLayersPrefix[1] + osc), DefaultLabelsName, i/OscCount);
	    				}
	    			}
	    			
	    		}//accessError
	    		else {ShowMessage("Error!\nProcess stopped!\n");}
	    			
	    	}//'NewWksName' column check
	    	else {ShowMessage("Error!\nColumn '"+DefaultGlobalLabelsName[1]+"' of layer '"+DefaultDOptionsLayersPrefix[0]+"' of '"+OptionsWPName+"' WorksheetPage is missign!\nProcess stopped!\n");}
	    		
    	}//'WksName' column check
    	else {ShowMessage("Error!\nColumn '"+DefaultGlobalLabelsName[0]+"' of layer '"+DefaultDOptionsLayersPrefix[0]+"' of '"+OptionsWPName+"' WorksheetPage is missign!\nProcess stopped!\n");}
    		
	}//'Options' book check
	else {ShowMessage("Error!\nWorksheetPage '"+OptionsWPName+"' is missing.\nProcess stopped!\n");}
//========================================================================================================
// END
}