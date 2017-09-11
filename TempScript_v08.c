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


void GraphAllInOne(string TemplateGraphPageName, int CurvesCount, int StartLayer, string postfix, vector<int> skip)
{
	StringArray GraphName(17), ColName(17), BookName(17), BookName02(3), ColName02(3);
	
	/*
	// 2017 05 SERIES
	GraphName[0] = "";	ColName[0] = "";	BookName[0] = "";
	GraphName[1] = "AllInOne00" + postfix;	ColName[1] = "PMTD1";	BookName[1] = "OutDPO7054";
	GraphName[2] = "AllInOne10" + postfix;	ColName[2] = "PMTD2";	BookName[2] = "OutDPO7054";
	GraphName[3] = "AllInOne20" + postfix;	ColName[3] = "PMTD3";	BookName[3] = "OutDPO7054";
	GraphName[4] = "AllInOne30" + postfix;	ColName[4] = "PMTD4";	BookName[4] = "OutDPO7054";
	GraphName[5] = "AllInOne40" + postfix;	ColName[5] = "PMTD5";	BookName[5] = "OutHMO3004";
	GraphName[6] = "AllInOne50" + postfix;	ColName[6] = "PMTD6";	BookName[6] = "OutHMO3004";
	GraphName[7] = "AllInOne60" + postfix;	ColName[7] = "PMTD7";	BookName[7] = "OutHMO3004";
	GraphName[8] = "AllInOne70" + postfix;	ColName[8] = "PMTD8";	BookName[8] = "OutTDS2024C";
	GraphName[9] = "AllInOne80" + postfix;	ColName[9] = "PMTD9";	BookName[9] = "OutTDS2024C";
	GraphName[10] = "AllInOne90" + postfix;	ColName[10] = "PMTD10";	BookName[10] = "OutTDS2024C";
	GraphName[11] = "AllXRay" + postfix;		ColName[11] = "PMTS2";	BookName[11] = "OutTDS2024C";
	GraphName[12] = "AllInOne25" + postfix;	ColName[12] = "Hama";	BookName[12] = "OutHMO3004";
	
	
	BookName02[0] = "OutLeCroy";	ColName02[0] = "Voltage";
	BookName02[1] = "OutLeCroy";	ColName02[1] = "Current";
	BookName02[2] = "";	ColName02[2] = "";
	*/
	
	/*
	// 2016 SERIES
	GraphName[0] = "";	ColName[0] = "";	BookName[0] = "";
	GraphName[1] = "AllInOneD1" + postfix;	ColName[1] = "PMTD1";	BookName[1] = "OutDPO7054";
	GraphName[2] = "AllInOneD2" + postfix;	ColName[2] = "PMTD2";	BookName[2] = "OutDPO7054";
	GraphName[3] = "AllInOneD3" + postfix;	ColName[3] = "PMTD3";	BookName[3] = "OutDPO7054";
	GraphName[4] = "AllInOneD4" + postfix;	ColName[4] = "PMTD4";	BookName[4] = "OutDPO7054";
	GraphName[5] = "AllInOneD5" + postfix;	ColName[5] = "PMTD5";	BookName[5] = "OutLeCroy";
	GraphName[6] = "AllInOneD6" + postfix;	ColName[6] = "PMTD6";	BookName[6] = "OutLeCroy";
	GraphName[7] = "AllInOneD7" + postfix;	ColName[7] = "PMTD7";	BookName[7] = "OutLeCroy";
	GraphName[8] = "AllInOneD8" + postfix;	ColName[8] = "PMTD8";	BookName[8] = "OutLeCroy";
	GraphName[9] = "AllInOneD9" + postfix;	ColName[9] = "PMTD9";	BookName[9] = "OutTDS3054BL";
	GraphName[10] = "AllInOneD10" + postfix;	ColName[10] = "PMTD10";	BookName[10] = "OutTDS3054BL";
	GraphName[11] = "AllInOneD11" + postfix;	ColName[11] = "PMTD11";	BookName[11] = "OutTDS3054BL";
	GraphName[12] = "AllInOneD12" + postfix;	ColName[12] = "PMTD12";	BookName[12] = "OutTDS3054BL";
	GraphName[13] = "AllInOneD13" + postfix;	ColName[13] = "PMTD13";	BookName[13] = "OutTDS3054BR";
	GraphName[14] = "AllInOneS1" + postfix;	ColName[14] = "PMTS1";	BookName[14] = "OutTDS3054BR";
	GraphName[15] = "AllInOneS2" + postfix;	ColName[15] = "PMTS2";	BookName[15] = "OutTDS3054BR";
	GraphName[16] = "AllInOneS3" + postfix;	ColName[16] = "PMTS3";	BookName[16] = "OutTDS3054BR";
	
	
	BookName02[0] = "OutTDS2024C";	ColName02[0] = "Voltage";
	BookName02[1] = "OutTDS2024C";	ColName02[1] = "PreI";
	BookName02[2] = "";	ColName02[2] = "";	// this variable's values will be loop throgh above BookName & ColName arrays
	*/
	
	// 2015 04 SERIES
	GraphName[0] = "";	ColName[0] = "";	BookName[0] = "";
	GraphName[1] = "AllInOnePMTD2" + postfix;	ColName[1] = "D2Pb3";	BookName[1] = "OutAxial";
	GraphName[2] = "AllInOnePMTD4" + postfix;	ColName[2] = "D4Pb10";	BookName[2] = "OutAxial";
	GraphName[3] = "AllInOnePMTD3" + postfix;	ColName[3] = "D3Pb30";	BookName[3] = "OutAxial";
	GraphName[4] = "AllInOnePMTD1" + postfix;	ColName[4] = "D1Fe40";	BookName[4] = "OutAxial";

	GraphName[5] = "AllInOnePMTD6" + postfix;	ColName[5] = "D6Pb3";	BookName[5] = "OutLatheral";
	GraphName[6] = "AllInOnePMTD8" + postfix;	ColName[6] = "D8Pb10";	BookName[6] = "OutLatheral";
	GraphName[7] = "AllInOnePMTD7" + postfix;	ColName[7] = "D7Pb30";	BookName[7] = "OutLatheral";
	GraphName[8] = "AllInOnePMTD5" + postfix;	ColName[8] = "D5Fe40";	BookName[8] = "OutLatheral";

	GraphName[9] = "AllInOnePMTD11" + postfix;	ColName[9] = "D11Pb50";	BookName[9] = "OutHama";
	GraphName[10] = "AllInOnePMTD12" + postfix;	ColName[10] = "D12Pb50";	BookName[10] = "OutHama";
	GraphName[11] = "AllInOnePMTD10" + postfix;	ColName[11] = "D10";	BookName[11] = "OutHama";
	GraphName[12] = "AllInOnePMTD9" + postfix;	ColName[12] = "D9";	BookName[12] = "OutHama";

	GraphName[13] = "AllInOnePMTDS1" + postfix;	ColName[13] = "DS1";	BookName[13] = "OutXRay";
	GraphName[14] = "AllInOnePMTDS2" + postfix;	ColName[14] = "DS2Pb50";	BookName[14] = "OutXRay";
	
	BookName02[0] = "OutXRay";	ColName02[0] = "U";
	BookName02[1] = "OutXRay";	ColName02[1] = "I";
	
	// REVERSED:
	//BookName02[0] = "Neutron";	ColName02[0] = "I";
	//BookName02[1] = "Neutron";	ColName02[1] = "U";
	BookName02[2] = "";	ColName02[2] = "";	// this variable's values will be loop throgh above BookName & ColName arrays
	
	/*
	// 2014 11 SERIES
	GraphName[0] = "";	ColName[0] = "";	BookName[0] = "";
	GraphName[1] = "AllInOnePMT1" + postfix;	ColName[1] = "FEU1";	BookName[1] = "PMT1";
	GraphName[2] = "AllInOnePMT2" + postfix;	ColName[2] = "FEU2";	BookName[2] = "PMT2";
	GraphName[3] = "AllInOnePMT3" + postfix;	ColName[3] = "FEU3";	BookName[3] = "PMT3";
	GraphName[4] = "AllInOneXRay" + postfix;	ColName[4] = "XRay";	BookName[4] = "Neutron";
	
	// BookName02[0] = "Neutron";	ColName02[0] = "U";
	// BookName02[1] = "Neutron";	ColName02[1] = "I";
	
	// REVERSED:
	BookName02[0] = "Neutron";	ColName02[0] = "I";
	BookName02[1] = "Neutron";	ColName02[1] = "U";
	BookName02[2] = "";	ColName02[2] = "";	// this variable's values will be loop throgh above BookName & ColName arrays
	*/
	
	/*
	// 2013 SERIES
	GraphName[0] = "";	ColName[0] = "";	BookName[0] = "";
	GraphName[1] = "AllInOneXRay" + postfix;	ColName[1] = "XRay";	BookName[1] = "Out1";
	GraphName[2] = "AllInOneN" + postfix;		ColName[2] = "Neutron";	BookName[2] = "Out1";
	
	BookName02[0] = "Out2";	ColName02[0] = "U";
	BookName02[1] = "Out2";	ColName02[1] = "I";
	//// REVERSED:
	//BookName02[0] = "Out2";	ColName02[0] = "I";
	//BookName02[1] = "Out2";	ColName02[1] = "U";
	
	BookName02[2] = "";	ColName02[2] = "";	// this variable's values will be loop throgh above BookName & ColName arrays
	*/
	
	int GraphNumber = 14;
	int GraphLayersCount = 3;
	int StartGraphLayer = 2; // zero-based index
	GraphPage TemplateGraphPage;
	
	
	TemplateGraphPage = Project.GraphPages(TemplateGraphPageName);
	Tree trTemplateFormat;
	trTemplateFormat = TemplateGraphPage.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
	
	StartLayer = StartLayer - 1;	// convert one-based index to zero-based
	if (StartLayer < 0) { StartLayer = 0; }
	
	for (int i=1; i<=GraphNumber; i++)
	{
		GraphPage NewGraphPage; //create new GraphPage window
		NewGraphPage.Create("origin");
		
		NewGraphPage.SetName(GraphName[i]);
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
		
		BookName02[2] = BookName[i];	
		ColName02[2] = ColName[i];
		//filling Layers of NewGraphPage with Curves
		for (int j=0; j<GraphLayersCount; j++)
		{//for all GraphLayers in NewGraphPage
			FillAllInOneGraphLayer(NewGraphPage.Layers(j), BookName02[j], ColName02[j], CurvesCount, StartLayer, skip);
		}
		
		
		
		bool FormatUpdErr = NewGraphPage.ApplyFormat( trTemplateFormat, true, true, true ); //changing format
		if (!FormatUpdErr) {out_str("Warning!\n GraphPage[index=" + i + "] ('"+NewGraphPage.GetName()+"') format update failed!\n");}
		
		NewGraphPage.Show = 0; //hiding new graph window
		NewGraphPage.Detach();
	}
}


