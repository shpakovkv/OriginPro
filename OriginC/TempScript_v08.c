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
void ShowMessage(string m)
void StringToStringReplacement(string name, int layer, string a, string b)
void StringToNum(string WPName, int layer)
double IntToDbl(int i)
void GetHWCenter(string DataBookName, string NewBookName, int SignalsCount)

*/
///////////////////////////////////////////////////////
void ShowMessage(string m)
{// message box and CommandWindow log
// CommandWindow log appears only when it active
	out_str(m);
	LT_execute("type -b " + m);
}


void StringToStringReplacement(string name, int layer, string a, string b)
{
	WorksheetPage wp = Project.WorksheetPages(name);
	 
	
	if (wp)
	{
		Worksheet wks = wp.Layers(layer);
		if (wks)
		{
			for (int c=0; c<wks.Columns.Count(); c++)
			{
				StringArray sa;
				wks.Columns(c).GetStringArray(sa);
				for (int i=0; i<sa.GetSize(); i++)
				{
					sa[i].Replace(a, b);
				}
				wks.Columns(c).PutStringArray(sa);
			}
			wp.Refresh(false);
		}
		else {ShowMessage("Layer #" + layer + " not found!");	}
	}
	else {ShowMessage("WorksheetPage '" + name + "' not found!");	}
				
}

void StringToNum(string WPName, int layer)
{
	//	StringToNum(FrontStat2, 1)
	layer=layer - 1;
	StringToStringReplacement(WPName, layer, ".", ",");
	WorksheetPage wp = Project.WorksheetPages(WPName);
	
	if (wp)
	{
		Worksheet wks = wp.Layers(layer);
		if (wks)
		{
			for (int c=0; c<wks.Columns.Count(); c++)
			{
				Dataset ds;
				ds.Attach(wks.Columns(c));
				StringArray sa;
				ds.GetStringArray(sa);
				vector<double> dd(0);
				for (int i=0; i<=ds.GetUpperBound(); i++) 	{	dd.Add(atof(sa[i]));	}
				for (i=0; i<=ds.GetUpperBound(); i++)		{	ds[i] = dd[i];			}
				ds.Detach();
			}
		}
	}
	
}

double IntToDbl(int i)
{
	double dd;
	dd = i;
	return dd;
}


void TimeCorr(string DataBookName, string NewBookName)
{
	//script example: GetHWCenter("OutPeak", "NewBook")

	WorksheetPage outwp;
	
	Dataset tempds, datads;
	int coli = 8; 					// Column Index = 8 for "X" data of HalfWidth
	int peakcoli = 2;				// Column Index = 2 for "X" data of Peak
	int fronti = 4;					// Column Index = 4 for "X" data of Front
	int period = 10;  				// Number of columns of one signals
	
	StringArray ColNames(16);
	vector<int> CorrColNum = {2, 4, 5, 8};
	vector<double> TimeCorrValue = {-1.96, -11.42, 39.06, -10.5};
	int CorrColU = 11;
	int CorrColPeakI = 9;
	
	int SignalsCount = 5;
	
	ColNames[0] = "N";
	ColNames[1] = "PeakD1";
	ColNames[2] = "CenterD1";
	ColNames[3] = "PeakD5";
	ColNames[4] = "CenterD5";
	ColNames[5] = "PeakD13";
	ColNames[6] = "CenterD13";
	ColNames[7] = "PeakD11";
	ColNames[8] = "CenterD11";
	ColNames[9] = "PeakI";
	ColNames[10] = "CenterI";
	ColNames[11] = "UFront";
	
	ColNames[12] = "D1Corr";
	ColNames[13] = "D5Corr";
	ColNames[14] = "D13Corr";
	ColNames[15] = "D11Corr";
	
	WorksheetPage datawp(DataBookName);
	
	if (datawp)
	{
		outwp.Create("origin");
		outwp.SetName(NewBookName);
		Worksheet outwks = outwp.Layers(0);
		
		for (int i=0; i<ColNames.GetSize(); i++) 
		{	
			if (outwks.Columns.Count() < ColNames.GetSize()) {	outwks.AddCol();	}
			tempds.Attach(outwks.Columns(i));
			tempds.SetLowerBound(0);
			tempds.SetUpperBound(datawp.Layers.Count()-1);
			tempds.Detach();
			
			outwks.Columns(i).SetName(ColNames[i]);
		}
		
		for (i=0; i<datawp.Layers.Count(); i++)
		{
			Worksheet datawks = datawp.Layers(i);
			tempds.Attach(outwks.Columns(0));
			//tempds[i] = i+1;
			
			StringArray TempSA;
			tempds.GetStringArray(TempSA);
			TempSA[i] = datawks.GetName();
			tempds.PutStringArray(TempSA);
			tempds.Detach();
			
			for (int j=0; j<SignalsCount; j++)
			{
				// column Peak
				datads.Attach(datawks.Columns(j*period + peakcoli));
				tempds.Attach(outwks.Columns(1+j*2));
				tempds[i] = datads[0];
				datads.Detach();
				tempds.Detach();
				
				// column Center
				datads.Attach(datawks.Columns(j*period + coli));
				tempds.Attach(outwks.Columns(1+j*2+1));
				tempds[i] = datads[0] + (datads[1] - datads[0])/2;
				datads.Detach();
				tempds.Detach();
				
			}
			
			// column Voltage Front (UFront)
				j = SignalsCount;
				datads.Attach(datawks.Columns(j*period + fronti));
				tempds.Attach(outwks.Columns(1+j*2));
				tempds[i] = datads[0];
				datads.Detach();
				tempds.Detach();
		}
		
		//StringToStringReplacement(NewBookName, 0, ",", ".");
		
		Dataset UFrontDS, PeakIDS;
		UFrontDS.Attach(outwks.Columns(CorrColU));
		PeakIDS.Attach(outwks.Columns(CorrColPeakI));
		for (i=0; i<SignalsCount-1; i++)
		{
			datads.Attach(outwks.Columns(CorrColNum[i]));
			tempds.Attach(outwks.Columns(12+i));
			for (int j=0; j<datawp.Layers.Count(); j++)
			{
				tempds[j] = datads[j] + TimeCorrValue[i] - PeakIDS[j] + UFrontDS[j];
			}
			datads.Detach();
			tempds.Detach();
		}
		UFrontDS.Detach();
		PeakIDS.Detach();
		
	}
}


void GetPeaksForGraph(string DataBookName, string NewBookName, int SignalsCount)
{
	//script example: GetHWCenter("OutPeak", "NewBook")

	WorksheetPage outwp;
	WorksheetPage datawp(DataBookName);
	Dataset tempds, datads;
	int coli = 3; 					// Column Index (zero-based) = 3 for "Y" data of Peak
	int period = 10;  				// Number of columns of one signals
	int colnum = 2;
	StringArray ColNames(2), ColTypes(2);
	
	ColNames[0] = "X";
	ColNames[1] = "Peak";
	
	ColTypes[0] = "x";
	ColTypes[1] = 'y';
	
	
	if (datawp)
	{
		outwp.Create("origin");
		outwp.SetName(NewBookName);

		for (int i=0; i<datawp.Layers.Count(); i++)
		{
			Worksheet datawks = datawp.Layers(i);
			if (outwp.Layers.Count() < i+1)	{outwp.AddLayer();}
			Worksheet outwks = outwp.Layers(i);
			outwks.SetName(datawks.GetName());
			
			for (int i=0; i<colnum; i++) 
			{	
				if (outwks.Columns.Count() < colnum) {	outwks.AddCol();	}
				tempds.Attach(outwks.Columns(i));
				tempds.SetLowerBound(0);
				tempds.SetUpperBound(SignalsCount-1);
				tempds.Detach();
				outwks.Columns(i).SetName(ColNames[i]);
				if (ColTypes[i] == 'x') {outwks.Columns(i).SetType(OKDATAOBJ_DESIGNATION_X);}
					else if (ColTypes[i] == 'y') {outwks.Columns(i).SetType(OKDATAOBJ_DESIGNATION_Y);}
			}
			
			tempds.Attach(outwks.Columns(0));
			for (i=0; i<SignalsCount; i++)
			{
				tempds[i] = i*10;
			}
			tempds.Detach();
			
			tempds.Attach(outwks.Columns(1));
			for (int j=0; j<SignalsCount; j++)
			{
				
				datads.Attach(datawks.Columns(j*period + coli));
				tempds[j] = -datads[0];
				datads.Detach();
			}
			tempds.Detach();
		}
	}
}
