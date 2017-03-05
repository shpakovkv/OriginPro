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
#include <Array.h>
#include <graph.h>
////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.

struct Front
{
	double RiseLow;
	double RiseUp;
	double FallLow;
	double FallUp;
	
	int RiseLowIndex;
	int RiseUpIndex;
	int FallLowIndex;
	int FallUpIndex;
};

struct DoublePoint
{
	double x1,x2,y1,y2;
};

//==========================================================================================================================================
//	\\-- FUNCTIONS --//
//	 \\=============//

Front FrontConstructor()
{
	Front* F = new Front;
	F->RiseLow = 0;
	F->RiseUp = 0;
	F->FallLow = 0;
	F->FallUp = 0;
	
	F->RiseLowIndex = 0;
	F->RiseUpIndex = 0;
	F->FallLowIndex = 0;
	F->FallUpIndex = 0;
	
	return F;
}

DoublePoint DoublePointConstructor()
{
	DoublePoint* dp = new DoublePoint;
	dp->x1 = 0;
	dp->x2 = 0;
	dp->y1 = 0;
	dp->y2 = 0;
	return dp;
}

//--------------------------------------------------------
void ShowMessage(string m)
{// message box and CommandWindow log
// CommandWindow log appears only when it active
	out_str(m);
	LT_execute("type -b " + m);
}


void FillDatasetWthNumRows(Dataset ds, int LowerBound, int UpperBound)
{
	ds.SetLowerBound(LowerBound);
	ds.SetUpperBound(UpperBound);
	for (int i=LowerBound; i<=UpperBound; i++)
	{
		ds[i] = i;
	}
}

  //=======================================================================================================================
 //		NOISE FILTER 
//======================================

vector<double> NoiseFiltration(vector<double> data, int FilterExtent)
{
	vector<double> h;
	int start;
	
	if (FilterExtent == 1)
	{
		// Fc ~= 100 MHz
		vector<double> h1 = {0,   0.000011510267176,   0.000049277083182,   0.000106532773351,   0.000126346201842,  0,  -0.000371474334318,  -0.000952423122961,  -0.001450849287325,  -0.001321612732363,   0,   0.002649912239950,   0.005875567674448,   0.007909647494990,   0.006482612593007,  0,  -0.011019372946171,  -0.022997330718705,  -0.029604976736698,  -0.023638292503258,   0,   0.041396106048532,   0.094227435326779,   0.146834815368294,   0.185685267443231,   0.200002603734039,   0.185685267443231,   0.146834815368294,   0.094227435326779,   0.041396106048532,   0,  -0.023638292503258,  -0.029604976736698,  -0.022997330718705,  -0.011019372946171,  0,   0.006482612593007,   0.007909647494990,   0.005875567674448,   0.002649912239950,   0,  -0.001321612732363,  -0.001450849287325,  -0.000952423122961,  -0.000371474334318,  0,   0.000126346201842,   0.000106532773351,   0.000049277083182,   0.000011510267176,   0};
		start = 25;
		h = h1;
	}
	else if (FilterExtent == 2)
	{
		// Fc ~= 50 MHz
		vector<double> h2 = {0.000004551602719,   0.000018746084318,   0.000042192361496,   0.000066272463426,   0.000066859541032,  0,  -0.000196575782539,  -0.000592488345108,  -0.001242256108799,  -0.002152431680217,  -0.003243503319672,  -0.004315753711652,  -0.005030818776290,  -0.004920474778187,  -0.003430451381534,   0,   0.005831201942810,   0.014306299470198,   0.025348575846779,   0.038498274422096,   0.052909754707675,   0.067419364171203,   0.080680059723455,   0.091343767979186,   0.098260427118263,   0.100656812898682,   0.098260427118263,   0.091343767979186,   0.080680059723455,   0.067419364171203,   0.052909754707675,   0.038498274422096,   0.025348575846779,   0.014306299470198,   0.005831201942810,   0,  -0.003430451381534,  -0.004920474778187,  -0.005030818776290,  -0.004315753711652,  -0.003243503319672,  -0.002152431680217,  -0.001242256108799,  -0.000592488345108,  -0.000196575782539,  0,   0.000066859541032,   0.000066272463426,   0.000042192361496,   0.000018746084318,   0.000004551602719};
		start = 25;
		h = h2;
	}
	else if (FilterExtent > 2)
	{
		// Fc ~= 20 MHz
		vector<double> h3 = {0,   0.000003613300076,   0.000018970023076,   0.000060707557283,   0.000152454275162,   0.000328865389835,   0.000636919481552,   0.001135990707492,   0.001896261357743,   0.002995174789980,   0.004511845476966,   0.006519621444056,   0.009077301194561,   0.012219798431914,   0.015949275390220,   0.020227884694951,   0.024973237047094,   0.030057532477068,   0.035310963853793,   0.040529553767360,   0.045487071610118,   0.049950162728964,   0.053695376998310,   0.056526475515182,   0.058290270820057,   0.058889343334375,   0.058290270820057,   0.056526475515182,   0.053695376998310,   0.049950162728964,   0.045487071610118,   0.040529553767360,   0.035310963853793,   0.030057532477068,   0.024973237047094,   0.020227884694951,   0.015949275390220,   0.012219798431914,   0.009077301194561,   0.006519621444056,   0.004511845476966,   0.002995174789980,   0.001896261357743,   0.001135990707492,   0.000636919481552,   0.000328865389835,   0.000152454275162,   0.000060707557283,   0.000018970023076,   0.000003613300076,   0};
		start = 25;
		h = h3;
	}
	
	vector<double> output(data.GetSize()-start);
	
	for (int i=start; i<data.GetSize(); i++)
	{
		double sum = 0;
		int len = h.GetSize()-1;
		if (i<h.GetSize()) 	{	len = i;	}
			
		for (int n=0; n<=len; n++)
		{
			sum = sum + h[n]*data[i-n];
		}
		
		output[i-start] = sum;
	}
	return output;
}

  //====================================================================================================================
 // 		MAXIMUM VALUE
//=======================================

double GetMaxDSValue(Dataset y, int LowerBoundIndex, int UpperBoundIndex)
{
	double max;
	max = y[LowerBoundIndex];
	for (int i=LowerBoundIndex; i<=UpperBoundIndex; i++)
	{
		if (max < y[i]) {max = y[i];}
	}
	return max;
}



  //=========================================================================================================
 //		INDEX of MAXIMUM VALUE
//=====================================

int GetDSMaxValIndex(Dataset y, double max, int LowerBoundIndex, int UpperBoundIndex, bool reverse)
{
	if (reverse)
	{
		for (int i=UpperBoundIndex; i>=LowerBoundIndex; i--)
		{
			if (y[i] == max) {return i;}
		}
	}
	else
	{
		for (int i=LowerBoundIndex; i<=UpperBoundIndex; i++)
		{
			if (y[i] == max) {return i;}
		}
	}
	
	return -1;
}



  //====================================================================================================================================
 //		FRONT CALCULATION
//====================================

Front FrontCalc(Dataset ds, int MaxIndex, int LowerBoundIndex, int UpperBoundIndex, double LowerTreshold, double UpperTreshold, bool reverse)
{
	double max = ds[MaxIndex];
	Front& F = FrontConstructor();
	
	int i = MaxIndex;
	int di = -1;
	if (reverse) {di = 1;}
	
// LOWER:

	while ((i>=LowerBoundIndex) && (i<=UpperBoundIndex) && (ds[i] >= LowerTreshold))
	{
		i=i+di;
	}
	i=i-di; // lower point = Y(i) >= Treshold (lower)
	
	F.RiseLow = ds[i];
	F.RiseLowIndex = i;

// UPPER:
	
	while ((i>=LowerBoundIndex) && (i<=UpperBoundIndex) && (ds[i] <= UpperTreshold))
	{
		i=i-di;
	} // Upper point = Y(i) >= Treshold (upper)
	if ((i<LowerBoundIndex) ||(i>UpperBoundIndex))	{	i = i+di;	}	// for left/right bound condition
	
	F.RiseUp = ds[i];
	F.RiseUpIndex = i;
	
	return F;
}



  //===============================================================================================================
 //		HALF WIDTH CALCULATION
//======================================

DoublePoint HalfWidthCalc(Dataset X, Dataset Y, int MaxIndex)
{
	double MaxX = X[MaxIndex];
	double MaxY = Y[MaxIndex];
	DoublePoint& hw = DoublePointConstructor();
	hw.y1 = MaxY/2;
	hw.y2 = MaxY/2;
	
// Left point
	for (int i=MaxIndex; i>=Y.GetLowerBound(); i--)
	{
		if (i == Y.GetLowerBound())
		{
			hw.x1 = X[i];
			break;
		}
		
		else if ((i>Y.GetLowerBound()) && (Y[i] == MaxY/2))
		{
			hw.x1 = X[i];
			break;
		}
		
		else if ((i>Y.GetLowerBound()) && (Y[i] < MaxY/2))
		{
			/*		y = a*x +b
					a = (Y2-Y1)/(X2-X1)
					b = Y1 - a*X1
					Xhw = (Yhw - b)/a = (Yhw - Y1 + a*X1)/a = X1 + (Yhw - Y1)/a =>
					X1 + (Yhw - Y1)*(X2-X1)/(Y2-Y1)
					
					Y1 = Y[i]		X1 = X[i]
					Y2 = Y[i+1]		X2 = X[i+1]
					Y2 > Y1			X2 > X1
			*/
			hw.x1 =	X[i] + (hw.y1-Y[i])*(X[i+1] - X[i])/(Y[i+1] - Y[i]);
			break;
		}
	}
	
// Right point
	for (i=MaxIndex; i<=Y.GetUpperBound(); i++)
	{
		if (i == Y.GetUpperBound())
		{
			hw.x2 = X[i];
			break;
		}
		
		else if ((i<=Y.GetUpperBound()) && (Y[i] == MaxY/2))
		{
			hw.x2 = X[i];
			break;
		}
		
		else if ((i<=Y.GetUpperBound()) && (Y[i] < MaxY/2))
		{
			/*
					Xhw = X1 + (Yhw - Y1)*(X2-X1)/(Y2-Y1)
					
					Y1 = Y[i-1]		X1 = X[i-1]
					Y2 = Y[i]		X2 = X[i]
					Y2 > Y1			X2 > X1
			*/
			hw.x2 =	X[i-1] + (hw.y1-Y[i-1])*(X[i] - X[i-1])/(Y[i] - Y[i-1]);
			break;
		}
	}
	
	return hw;
}



  //=======================================================================================================================
 //		GET COLUMN INDEX 
//======================================

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



  //=======================================================================================================================
 //		GET DIGITS OF INTEGER 
//======================================

int GetDigit(int x)
{
	int tempD, D, i;
	tempD = 1;
	i=1;
	while (((i<=9) && (x >= tempD)) || ((x==0)&&(i==1)))
	{
		tempD=tempD*10;
		i++;
	}
	D=i-1;
	return D;
}



  //=======================================================================================================================
 //		STRING POSTFIX INCREASE
//========================================

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


//!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+
  //========================================================================================================================
 //		GRAPH FORMAT
//=====================

void SA_SetDefaultGraphFormat(GraphPage gp)
{
	GroupPlot gplot = gp.Layers(0).Groups(0);
		
	vector<int> vNester;
    
    // the Nester is an array of types of objects to do nested cycling in the group
    vNester.SetSize(4);  // four types of setting to do nested cycling in the group
    vNester[0] = 0;  // cycling line color in the group
    vNester[1] = 3;  // cycling symbol type in the group
    vNester[2] = 2;  // cycling line style in the group
    vNester[3] = 8;  // cycling symbol interior in the group
    
    gplot.Increment.Nester.nVals = vNester;  // set Nester of the grouped plot
    
    int nNumPlots = gp.Layers(0).DataPlots.Count();  // nNumPlots = 4 in this example
    
    vector<int> vLineColor, vSymbolShape, vLineStyle, vSymbolInterior;
    
    // four vectors for setting line color, symbol shape, line style and symbol interior of grouped plots
    vLineColor.SetSize(nNumPlots);
    vSymbolShape.SetSize(nNumPlots);
    vLineStyle.SetSize(nNumPlots);
    vSymbolInterior.SetSize(nNumPlots);
    
    // setting corresponding values to four vectors
    for(int nPlot=0; nPlot<nNumPlots; nPlot++)
    {
        if(nNumPlots>=4)
        {
            switch (nPlot)
            {
            	case 0:  // setting for plot 0
                    vLineColor[nPlot] = SYSCOLOR_BLACK;
                    vSymbolShape[nPlot] = 0;  // no symbol
                    vLineStyle[nPlot] = 0;    // solid
                    vSymbolInterior[nPlot] = 0;  // solid
                    break;
                case 1:  // setting for plot 0
                    vLineColor[nPlot] = SYSCOLOR_RED;
                    vSymbolShape[nPlot] = 9;  // star
                    vLineStyle[nPlot] = 0;    // solid
                    vSymbolInterior[nPlot] = 0;  // solid
                    break;
                case 2:  // setting for plot 1
                    vLineColor[nPlot] = SYSCOLOR_GREEN;
                    vSymbolShape[nPlot] = 3;  // up triangle
                    vLineStyle[nPlot] = 0;    // solid
                    vSymbolInterior[nPlot] = 0;  // solid
                    break;
                case 3:  // setting for plot 2
                    vLineColor[nPlot] = SYSCOLOR_CYAN;
                    vSymbolShape[nPlot] = 4;  // down triangle
                    vLineStyle[nPlot] = 0;    // solid
                    vSymbolInterior[nPlot] = 0;  // solid
                    break;
                default:  // setting for other plot
                    vLineColor[nPlot] = SYSCOLOR_BLUE;
                    vSymbolShape[nPlot] = 1;  // square
                    vLineStyle[nPlot] = 0;    // solid
                    vSymbolInterior[nPlot] = 0;  // solid
                    break;
            }
        }
    }
    
    Tree trFormat;
    trFormat.Root.Increment.LineColor.nVals = vLineColor;  // set line color to theme tree
    trFormat.Root.Increment.Shape.nVals = vSymbolShape;  // set symbol shape to theme tree
    trFormat.Root.Increment.LineStyle.nVals = vLineStyle;  // set line style to theme tree
    trFormat.Root.Increment.SymbolInterior.nVals = vSymbolInterior;  // set symbol interior to theme tree
    
    // apply theme tree
    if(0 == gplot.UpdateThemeIDs(trFormat.Root) ) {  bool bb = gplot.ApplyFormat(trFormat, true, true); }
    
    // now change format of layer axes
    trFormat = gp.Layers(0).GetFormat();
    trFormat.Root.Axes.X.Ticks.TopTicks.show.nVal = 1;		//turn on top axes
    trFormat.Root.Axes.X.Ticks.BottomTicks.show.nVal = 1;
    trFormat.Root.Axes.Y.Ticks.RightTicks.show.nVal = 1;	//turn on right axes
    trFormat.Root.Axes.Y.Ticks.LeftTicks.show.nVal = 1;
    
    trFormat.Root.Axes.X.Titles.TopTitle.show.nVal = 1;
    
	// apply theme tree
    if(0 == gp.Layers(0).UpdateThemeIDs(trFormat.Root) ) { bool bb = gp.Layers(0).ApplyFormat(trFormat, true, true);   }   
    
    //this change settings can be done after needed axes are already turned on
    trFormat.Root.Axes.X.Ticks.BottomTicks.Major.nVal = 1;	//set major ticks to 'In' for X and Y axes
    trFormat.Root.Axes.X.Ticks.BottomTicks.Minor.nVal = 1;
    trFormat.Root.Axes.Y.Ticks.LeftTicks.Major.nVal = 1;	//set minor ticks to 'In' for X and Y axes
    trFormat.Root.Axes.Y.Ticks.LeftTicks.Minor.nVal = 1;
    
    // apply theme tree
    if(0 == gp.Layers(0).UpdateThemeIDs(trFormat.Root) ) {  bool bb = gp.Layers(0).ApplyFormat(trFormat, true, true); } 
    
}

   //!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+!@#$%^&*()_+
  //========================================================================================================================
 //		GRAPH
//=====================
void AddTextToGraphLayer(GraphPage gp, int GraphLayerIndex, string text, int FontSize, int left, int top)
{
	GraphLayer gl = gp.Layers(GraphLayerIndex);
	
	//creatting Text object
	GraphObject TextObj = gl.CreateGraphObject(GROT_TEXT); 
	TextObj.Text = text;
	TextObj.Attach = 0;  //attach to Layer
	
	
	Tree trFormat;
	trFormat.Root.Dimension.Units.nVal = UNITS_PAGE; // units in % of page
	trFormat.Root.Font.Size.nVal = FontSize;
	if (0 == TextObj.UpdateThemeIDs(trFormat.Root)) { TextObj.ApplyFormat(trFormat, true, true);	}
    
	if (left == -1) // -1 for center object along X axis
	{ 
		trFormat.Root.Dimension.Left.nVal = 50; // units in % of page
		if (0 == TextObj.UpdateThemeIDs(trFormat.Root)) { TextObj.ApplyFormat(trFormat, true, true);	}
		int TextLeft = TextObj.Left; //in page (i.e. paper, logical) coordinate units
		
		int TextWidth = TextObj.Width;
		
		TextLeft = TextLeft - TextObj.Width/2;
		TextObj.Left = TextLeft;
	}
	else	
	{	
		trFormat = TextObj.GetFormat(FPB_ALL, FOB_ALL, true, true);
		trFormat.Root.Dimension.Left.nVal = left;	
		if (0 == TextObj.UpdateThemeIDs(trFormat.Root)) { TextObj.ApplyFormat(trFormat, true, true);	}
	}
		
	if (top == -1) 	// -1 for center object along Y axis
	{
		trFormat.Root.Dimension.Top.nVal = 50; // units in % of page
		if (0 == TextObj.UpdateThemeIDs(trFormat.Root)) { TextObj.ApplyFormat(trFormat, true, true);	}
		int TextTop = TextObj.Top; //in page (i.e. paper, logical) coordinate units
		TextTop = TextTop - TextObj.Height/2;
	}
	else	
	{	
		trFormat = TextObj.GetFormat(FPB_ALL, FOB_ALL, true, true);
		trFormat.Root.Dimension.Top.nVal = top;	
		if (0 == TextObj.UpdateThemeIDs(trFormat.Root)) { TextObj.ApplyFormat(trFormat, true, true);	}
	}
		
	
		
	//out_str("TextWidth = " + TextObj.Width);
	//out_str("TextLeft = " + TextObj.Left);
	//out_str("Height = " + TextObj.Height);
}


void SA_GraphProcess(Worksheet wks, int DataColumnIndex, int CurveCount, int StartX, int StopX)
{
	GraphPage NewGraph; //create new GraphPage window
	NewGraph.Create("origin");
						    	
	for (int i=0; i<CurveCount; i++)
	{
		Curve cr;
		cr.Attach(wks.Columns(DataColumnIndex + i*2));
		NewGraph.Layers(0).AddPlot(cr, IDM_PLOT_LINESYMB);
		
	}
	
	NewGraph.Layers(0).GroupPlots(0, CurveCount-1);	
	SA_SetDefaultGraphFormat(NewGraph);
	
	// changeing Y scale by default
	NewGraph.Layers(0).Rescale();
	legend_update(NewGraph.Layers(0));
	
	// Changeing X Scale
	Tree trFormat;
	trFormat.Root.Axes.X.Scale.From.nVal = StartX;
	trFormat.Root.Axes.X.Scale.To.nVal = StopX;
	if(0 == NewGraph.Layers(0).UpdateThemeIDs(trFormat.Root) ) {  bool bb = NewGraph.Layers(0).ApplyFormat(trFormat, true, true); }
		    
	// Adding Title to graph
	string text = wks.GetName() + " (" + wks.Columns(DataColumnIndex).GetName() + ")";
	AddTextToGraphLayer(NewGraph, 0, text, 26, -1, 0); // -1 for 'center' at GraphPage
	
	NewGraph.Show = 0; // hiding graph for better performance
	
	//Tree trFormat;
	//trFormat.Root.Legend.show.nVal = 1;
    //trFormat.Root.Axes.X.Scale.From.nVal = 0;
    //trFormat.Root.Axes.X.Scale.To.nVal = 1000;
    //trFormat.Root.Axes.X
	
}


//=======================================================================================================================
 //		TEST FUNCTION
//======================================

void savetree(string name)
{
	
	GraphPage gp; //create new GraphPage window
	gp.Create("origin");
	
	WorksheetPage wp;
	wp = Project.WorksheetPages("Out");
	Worksheet wks = wp.Layers(0);
	
						    	
	for (int i=1; i<=9; i=i+2)
	{
		Curve cr;
		cr.Attach(wks.Columns(i));
		gp.Layers(0).AddPlot(cr, IDM_PLOT_LINESYMB);
		
	}
	
	gp.Layers(0).GroupPlots(0, 5);	
	SA_SetDefaultGraphFormat(gp);
	
	gp.Layers(0).Rescale();
	legend_update(gp.Layers(0));
	gp.Show = 0;
	
	Tree trFormat;
	gp.Layers(0).UpdateThemeIDs(trFormat.Root);
	trFormat.Root.Axes.X.Titles.AddNode("TopTitle");
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Text").strVal = "TopTitle";
	
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Dimension");
		trFormat.Root.Axes.X.Titles.TopTitle.Dimension.AddNode("XOffset").nVal = 0.;
		trFormat.Root.Axes.X.Titles.TopTitle.Dimension.AddNode("YOffset").nVal = 0.;
		
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Font");
		trFormat.Root.Axes.X.Titles.TopTitle.Font.AddNode("Face").nVal = 1391921696;
		trFormat.Root.Axes.X.Titles.TopTitle.Font.AddNode("Size").nVal = 26;
		
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Background");
		trFormat.Root.Axes.X.Titles.TopTitle.Background.AddNode("Dimension");
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Dimension.AddNode("Left").nVal = 12.4;
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Dimension.AddNode("Top").nVal = 12.4;
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Dimension.AddNode("Right").nVal = 12.4;
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Dimension.AddNode("Bottom").nVal = 12.4;
		trFormat.Root.Axes.X.Titles.TopTitle.Background.AddNode("Shadow");
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Shadow.AddNode("Style").nVal = 0;
		trFormat.Root.Axes.X.Titles.TopTitle.Background.AddNode("Border");
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Border.AddNode("Color").nVal = -4;
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Border.AddNode("Style").nVal = 0;
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Border.AddNode("Width").nVal = 0.;
		trFormat.Root.Axes.X.Titles.TopTitle.Background.AddNode("Fill");
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Fill.AddNode("Pattern");
				trFormat.Root.Axes.X.Titles.TopTitle.Background.Fill.Pattern.AddNode("Style").nVal = 0;
			trFormat.Root.Axes.X.Titles.TopTitle.Background.Fill.AddNode("Color").nVal = -4;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Show").nVal = 1;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Color").nVal = 0;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Angle").nVal = 0.;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Text").strVal = "TOP_TITLE";
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("TextVerbatim").nVal = 0;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("LinkToVars").nVal = 0;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("SysFont").nVal = 0;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("BehindData").nVal = 0;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("KeepInside").nVal = 0;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("States").nVal = 0;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Event").nVal = 0;
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Script").strVal = "";
	
	if(0 == gp.Layers(0).UpdateThemeIDs(trFormat.Root) )  {	bool bb = gp.Layers(0).ApplyFormat(trFormat, true, true);  }
	else {	out_str("Failed to update Theme IDs! (TopTotle)");}	
	
	trFormat.Root.Axes.X.Titles.TopTitle.show.nVal = 1;
	if(0 == gp.Layers(0).UpdateThemeIDs(trFormat.Root) ) 
	{  
		bool bbTrue = true;
		bool bbFalse = false;
		bool bb = gp.Layers(0).ApplyFormat(trFormat, true, true);
		int sgagaerr = 0;
	}  
	//else {out_str("Failed to update Theme IDs! (Add Text)");}
		
	trFormat.Root.Axes.X.Titles.TopTitle.AddNode("Text").strVal = "TOP_TITLE";
	if(0 == gp.Layers(0).UpdateThemeIDs(trFormat.Root) )  
	{	
		bool bbTrue = true;
		bool bbFalse = false;
		bool bb = gp.Layers(0).ApplyFormat(trFormat, true, true);  
		int sgfjhfg = 0;
	}
	else {	out_str("Failed to update Theme IDs! (TopTotle)");}	
		
	//Tree tr;
	//tr = gp.Layers(0).GetFormat(FPB_ALL, FOB_ALL, true, true);
	//tr.Save("c:\\Tektronix\TheTreeNEWWWWWWW.xml");
}

void OnlyForTest(string name)
{
	GraphPage gp = Project.GraphPages("Graph20");
	AddTextToGraphLayer(gp, 0, name, 26, -1, 0);
	
}


void SignalFrontAndDuration(string DatabookName, string TableName, int SignalsCount)
{
	int ColCount = 6;
	StringArray DefaultColNames(ColCount), DefaultDataColNames(4), TempDataColNames;
	vector<int> DefaultColIndexes(8);
	
	DefaultColNames[0] = "FRLowX";
	DefaultColNames[1] = "FRMidleX";
	DefaultColNames[2] = "PeakX";
	DefaultColNames[3] = "PeakY";
	DefaultColNames[4] = "HW";
	DefaultColNames[5] = "FFLowX";
	
	DefaultDataColNames[0] 	= "Peak";
	DefaultDataColNames[1] 	= "FR";
	DefaultDataColNames[2] 	= "FF";
	DefaultDataColNames[3] 	= "HW";
	
	//RiseFront low point [X] 
	DefaultColIndexes[2] = 4;
	//RiseFront middle point [X] (from HW)
	DefaultColIndexes[3] = 8;
	//Peak [X]
	DefaultColIndexes[4] = 2;
	//Peak [Y]
	DefaultColIndexes[5] = 3;
	// HalfWidth
	DefaultColIndexes[6] = 8;
	// FrontFall low point [X]
	DefaultColIndexes[7] = 6;
	
	
	
	//string name2, Front, Duration;
	WorksheetPage DataWP;
	DataWP = Project.WorksheetPages(DatabookName);
	
	Worksheet DataWks, OutWks;
	
	if (DataWP)
	{
		if (!Project.Pages(TableName))
		{
			WorksheetPage OutWP();
			OutWP.Create("origin");
			OutWP.SetName(TableName);
			OutWks = OutWP.Layers(0);
			
			// ADDING COLUMNS
			while (OutWks.Columns.Count() < (2 + ColCount*SignalsCount))
			{
				OutWks.AddCol();
			}
			
			// RENAMEING new COLUMNS
			OutWks.Columns(0).SetName("N");
			OutWks.Columns(0).SetLowerBound(0);
			OutWks.Columns(0).SetUpperBound(DataWP.Layers.Count());
			OutWks.Columns(1).SetName("Layer");
			OutWks.Columns(1).SetLowerBound(0);
			OutWks.Columns(1).SetUpperBound(DataWP.Layers.Count());
			
			DataWks = DataWP.Layers(0);
			
			for (int i=0; i<SignalsCount; i++)
			{
				for (int j=0; j<ColCount; j++)
				{
					string name = DefaultColNames[j];
		    		while (OutWks.Columns(name))
		    		{
		    			name = StringPostfixIncrease(name);
		    		}
		    		OutWks.Columns(2 + i*ColCount + j).SetName(name);
		    		OutWks.Columns(2 + i*ColCount + j).SetLowerBound(0);
					OutWks.Columns(2 + i*ColCount + j).SetUpperBound(DataWP.Layers.Count()-1);
				}
			}
			DataWks.Detach();
			
			// FILLING COLUMNS with DATA
			for (i=0; i<DataWP.Layers.Count(); i++)
			{
				DataWks = DataWP.Layers(i);
				
				Dataset TempDS, TempDS2;
				// SET ROW NUMBER
				TempDS.Attach(OutWks.Columns(0));
				TempDS[i] = i+1;
				TempDS.Detach();
				
				// SET LAYER NAME
				StringArray TempSA(1);
				TempSA[0] = DataWP.Layers(i).GetName();
				OutWks.Columns(1).PutStringArray(TempSA, i);
				
				for(int s=0; s<SignalsCount; s++)
				{
					//TempDataColNames = DefaultDataColNames;
					for (int j=0; j<ColCount; j++)
					{
						TempDS.Attach(OutWks.Columns(2 + j + s*ColCount));
						TempDS2.Attach(DataWks.Columns(DefaultColIndexes[2 + j] + s*10));
						
						if (j+2 != 6 )	{	TempDS[i] = TempDS2[0];	} // DefaultColIndexes[6] == HalfWidth
						else {	TempDS[i] = TempDS2[1] - TempDS2[0];	}
						TempDS.Detach();
						TempDS2.Detach();
					}
				}
				DataWks.Detach();
			}
			
		}
		else {ShowMessage("Wrong OUTPUT WorksheetPages name!/n('" + TableName +"' already exist!)");}
	}
	else {ShowMessage("Wrong DATA WorksheetPages name!/n('" + DatabookName + "')");}
		
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
void StringToStringReplacement(string name, string layer, string a, string b)
{
	if (!Project.WorksheetPages(name)) 
	{
		ShowMessage("WorksheetPage '" + name + "' not found!");
	}
	else
	{
		Worksheet wks = Project.WorksheetPages(name).Layers(layer);
		if (!wks) 
		{ 
			ShowMessage("Layer '" + layer + "' not found!");
		}
		else
		{
			StringToStringReplacement(name, wks.GetIndex(), a, b)
		}
	}		
}


void CommaToDotReplacement(string name, int layer)
{
	StringToStringReplacement(name, layer, ",", ".")			
}
void DotToCommaReplacement(string name, int layer)
{
	StringToStringReplacement(name, layer, ".", ",")			
}
void CommaToDotReplacement(string name, string layer)
{
	StringToStringReplacement(name, layer, ",", ".")
}
    //FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
   //GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG
  //=======================================================================================================================
 //		MAIN PROCESS			
//===============================

void SignalAnalysisProcess()
{
	//variables
	Dataset TempDS;
	string DataBookName, NewDataBookName;
	vector<string> DataBookNames(0);
	vector<int> OscCounts(0), DataBookNumbers(0), BookNumbers(0);
	int OscCount, SignalsCount, PrefixLength, ProcessDirection, FilterExtent;
	WorksheetPage OptionsBook;
	Worksheet TempWks, OptionsWks, SettingsWks;
	
	//constants
	string OptionsBookName = "SAOptions";
	StringArray DefaultGlobalOptionsLabel(4), DefaultOptionsLayersPrefix(2), DefaultOptionsLabel(15),DefaultOutputColumnName(8), OutputColumnNames(5);
	
	DefaultGlobalOptionsLabel[0] = "DataBookName";
	DefaultGlobalOptionsLabel[1] = "OscCount";
	DefaultGlobalOptionsLabel[2] = "NewBookName";
	DefaultGlobalOptionsLabel[3] = "PrefixLength"; 	// number of char from Layer's name to get and
													// and to put in NewBook Layer's name as prefix
	DefaultOptionsLayersPrefix[0] = "GlobalOptions";
	DefaultOptionsLayersPrefix[1] = "Settings";
	
	DefaultOptionsLabel[0] = "Osc";
	DefaultOptionsLabel[1] = "ColName";
	DefaultOptionsLabel[2] = "Invert";			//Rise and Fal; Max and Min replacement
	DefaultOptionsLabel[3] = "Peaks";		//number of Maximum to find
	DefaultOptionsLabel[4] = "Differ";	//minimum relative difference between heights of peaks (in percentage of one), (to exept signal roughness based mistake)
	DefaultOptionsLabel[5] = "RiseLow";	// Lower point of RiseFront in percentage of one (one is the Maximum)
	DefaultOptionsLabel[6] = "RiseUp";	// Upper point of RiseFront
	DefaultOptionsLabel[7] = "FallUp";	// Lower point of FallFront 
	DefaultOptionsLabel[8] = "FallLow";	// Upper point of FallFront
	DefaultOptionsLabel[9] = "HalfW";		// (<>-1 no; ==-1 yes) Have program to calculate HalfHeight and Width on HalfHeight
	DefaultOptionsLabel[10] = "Start";			// Localize area of "X" to analyze
	DefaultOptionsLabel[11] = "Stop";
	DefaultOptionsLabel[12] = "Noise";	// Applies digital filter to signal
	DefaultOptionsLabel[13] = "Graph";
	DefaultOptionsLabel[14] = "Book";
	
	DefaultOutputColumnName[0] = "XPeak";
	DefaultOutputColumnName[1] = "YPeak";
	DefaultOutputColumnName[2] = "XRise";
	DefaultOutputColumnName[3] = "YRise";
	DefaultOutputColumnName[4] = "XFall";
	DefaultOutputColumnName[5] = "YFall";
	DefaultOutputColumnName[6] = "XHalf";
	DefaultOutputColumnName[7] = "YHalf";
	
	OutputColumnNames[0] = "Y";
	OutputColumnNames[1] = "Peak";
	OutputColumnNames[2] = "FR";
	OutputColumnNames[3] = "FF";
	OutputColumnNames[4] = "HW";
	
	OptionsBook = Project.WorksheetPages(OptionsBookName);
	
	// trying connect to 'Options' WorksheetPage
	if (OptionsBook)
	{
		bool AccessError = false;
		
		//GlobalOptions Worksheet existence check
		for (int i=0; i<DefaultOptionsLayersPrefix.GetSize(); i++)
		{
			if (OptionsBook.Layers(DefaultOptionsLayersPrefix[i]))
			{	//get direct access to GlobalOptions Worksheet of Options Worksheetpage
				switch(i)
				{
					case 0:
						OptionsWks = OptionsBook.Layers(DefaultOptionsLayersPrefix[i]);
						break;
					case 1:
						SettingsWks = OptionsBook.Layers(DefaultOptionsLayersPrefix[i]);
						break;
				}		
			}
			else 
			{	ShowMessage("Error!\n'"+DefaultOptionsLayersPrefix[i]+"' layer of '"+OptionsBookName+"' worksheetpage is missing.\n");
				AccessError = true;		}
		}
		
	
  //----------------------------------------------------------------------------------------
 // GLOBAL OPTIONS CHECK
//----------------------------
	
		if (!AccessError)
		{	for (int i=0; i<DefaultGlobalOptionsLabel.GetSize(); i++)
			{
				//existence check
				if (OptionsWks.Columns(DefaultGlobalOptionsLabel[i]))
				{
					TempDS.Attach(OptionsWks.Columns(DefaultGlobalOptionsLabel[i]));
					TempDS.SetLowerBound(0);
				
		//OSC COUNT
					if (i == 1)		
					{	
						//OscCount = TempDS[0]; //save oscillographs count
						
						//get access to 'OscCount' column
						for (int j=0; j<=TempDS.GetUpperBound(); j++)
						{
							OscCounts.Add(TempDS[j]); //save oscilligraphs count
						}
						TempDS.Detach();
						
						//OscCount value compatibility check (for Dataset NANUM value OscCount will be 0)
						for (j=0; j<OscCounts.GetSize(); j++)
						{
							if (OscCounts[j] < 1)
							{	
								ShowMessage("Error!\n '"+DefaultGlobalOptionsLabel[1]+"' in '"+OptionsBookName+"' WorksheetPage has incompatible value.\n");
								AccessError = true;		
							}
						}
						
						if (DataBookNames.GetSize() != OscCounts.GetSize())
							{
								ShowMessage("Error!\nThe number of rows in\n"+DefaultGlobalOptionsLabel[0]+" and "+DefaultGlobalOptionsLabel[1]+"columns are different!");
								AccessError = true;
							}
					}
				
		//DATA BOOK NAME
					else if (i == 0)
						{	
							//getting DataWorksheetpage names
							
							for (int j=0; j<=TempDS.GetUpperBound(); j++)
							{
								string tempstr;
								TempDS.GetText(j, tempstr);
								DataBookNames.Add(tempstr);
								if (!Project.WorksheetPages(tempstr))
									{
										ShowMessage("WorksheetPage '"+tempstr+"' do not exist.\nCheck DataBookName.");
										AccessError = true;
									}
							}
							TempDS.Detach();
						}

		//NEW BOOK NAME
					else if (i == 2) 
						{	TempDS.GetText(0,NewDataBookName);
							if (Project.WorksheetPages(NewDataBookName))
							{
								ShowMessage("WorksheetPage '"+NewDataBookName+"' already exist.\nEnter different NewBookName.");
								AccessError = true;
							}
						}
		// PREFIX
					else if (i == 3)
						{		
							if (TempDS[0] == NANUM)		{	PrefixLength = 0;			}
							else if (TempDS[0] < 0)		{	PrefixLength = 0;			}
							else 						{	PrefixLength = TempDS[0];	}
						}

					TempDS.Detach();
				}
				else 
				{	ShowMessage("Error!\n'"+DefaultGlobalOptionsLabel[i]+"' column of "+OptionsBookName+" worksheet Page is missing!\n");
					AccessError = true;		}
				
			}//OptionsWks columns check	(For_End)
			
			
	// OSC COUNT VALUE 
	
			if (!AccessError)
			{
				for (int j=0; j<OscCounts.GetSize(); j++)
				{
					//OscCount value compatibility check (for Dataset NANUM value OscCount will be 0)
					if (OscCounts[j] < 1)
						{	ShowMessage("Error!\n '"+DefaultGlobalOptionsLabel[i]+"' in '"+OptionsBookName+"' WorksheetPage has incompatible value.\n");
							AccessError = true;		}
					else
					{	// multiplicity check of OscCOunt and Input Data WorksheetPage.Layers.Count
						if (Project.WorksheetPages(DataBookNames[j]).Layers.Count() % OscCounts[j] > 0) 
						{	ShowMessage("Error!\nWorksheetPage '"+DataBookNames[j]+"' layer-s count("+Project.WorksheetPages(DataBookNames[j]).Layers.Count()+") is not a multiple of OscCount("+OscCounts[j]+").\nSome data won't be processed!\n");
							AccessError = true;		}
					}
				}
			}
			
		}//'GlobalOptions' layer's columns check

		
  //--------------------------------------------------------------------------		
 // SETTINGS CHECK
//-------------------------

		if (!AccessError)
		{	
			//'ColumnName' existence check 
	    			if (!SettingsWks.Columns(DefaultOptionsLabel[1])) 		// Columns[1] = Columns("Column Name")
	    			{//if missing
	    				ShowMessage("Error!\n 'ColumnName' column access error in "+OptionsBookName+".Layer[2].\n");
	    				AccessError = true;
	    			}
	    			else 
	    			{//if existing
	    				//getting access to 'ColumnName' column
	    				TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[1]));		// Columns[1] = Columns("Column Name")
	    				
	    				//get signals Count
	    				TempDS.SetLowerBound(0);
	    				SignalsCount = TempDS.GetUpperBound()+1;

	    				//for all cells in column
	    				for (int k=0; k<=SignalsCount-1; k++)
	    				{
	    					string s;
	    					TempDS.GetText(k,s);
	    					//emty values check
	    					if (s == "")
	    					{
	    						ShowMessage("Error!\nValue ColumnName["+(k+1)+"] in '"+DefaultOptionsLayersPrefix[1]+"' of '"+OptionsBookName+"' WorkshhetPage is empty.\n");
	    						AccessError = true;
	    					}
	    				}
	    				TempDS.Detach();
	    			}//-----------------------------------
		//Settings check
	    	if (!AccessError)
	    	{
	    		
	    		// getting DataBookNumbers
	    		TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[14]));
	    		for (int r=0; r<SignalsCount; r++)				//DefaultOptionsLabel[14] ~= "Book Number"
				{
					
					DataBookNumbers.Add(TempDS[r]);
				}
				TempDS.Detach();
				
				
				for (int i=0; i<DefaultOptionsLabel.GetSize(); i++)
				{
					if (i != 1)
					{
	
						if (SettingsWks.Columns(DefaultOptionsLabel[i]))		//DefaultOptionsLabel[1] ~= "Column Name"
						{
							TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[i]));
							TempDS.SetLowerBound(0);
							TempDS.SetUpperBound(SignalsCount-1);
							
							for (int r=0; r<SignalsCount; r++)
							{
							
		// EMPTY and NANUM VALUE CHECK
		
								if (TempDS[r] == NANUM)
								{
									if ((i == 2) || (i == 9) || (i == 12) || (i == 13))
									{
										// Invert[2]; HalfWidth[9]; NoiseSuppress[12]; GraphNeeded[13]
										TempDS[r]=0;
									}
									else if ((i == 5) ||(i == 6) ||(i == 7) ||(i == 8))
									{
										// RiseLow[5]; RiseUp[6]; FallUp[7]; FallLow[8]

										string val;
										TempDS.GetText(r, val);
										if ((val.Right(1) == "m") || (val.Right(1) == "M"))
										{
											val.TrimRight("mM");
											double vald = atof(val, false);
											if (vald == NANUM)
											{
												ShowMessage("Error!\nValue "+DefaultOptionsLabel[i]+"["+(r+1)+"] in '"+DefaultOptionsLayersPrefix[1]+"' of '"+OptionsBookName+"' WorkshhetPage has incompatible format.\n");
	    										AccessError = true;
											}
										}
										else
										{
											ShowMessage("Error!\nValue "+DefaultOptionsLabel[i]+"["+(r+1)+"] in '"+DefaultOptionsLayersPrefix[1]+"' of '"+OptionsBookName+"' WorkshhetPage has incompatible format.\n");
	    									AccessError = true;	
										}
										
									}
									else
									{
	    								ShowMessage("Error!\nValue "+DefaultOptionsLabel[i]+"["+(r+1)+"] in '"+DefaultOptionsLayersPrefix[1]+"' of '"+OptionsBookName+"' WorkshhetPage has incompatible format.\n");
	    								AccessError = true;
									}
	    						}
	    					
	    //  VALUE SMART CHECK
	    						if (!AccessError)
	    						{
	    							
	    						//OscNumber value check
	    							if (i == 0)		//DefaultOptionsLabel[0] ~= "Osc Number"
	    							{
	    								if (DataBookNumbers[r] > OscCounts.GetSize())
	    								{
	    									ShowMessage("Error at Layer '"+DefaultOptionsLayersPrefix[1]+"' of '"+OptionsBookName+"' WorksheetPage!\n"+DefaultOptionsLabel[14]+"["+DataBookNumbers[r]+"] value is greather than DataBooks count [" + OscCounts.GetSize() + "].");
	    									AccessError = true;
	    								}
																// DataBookNumbers = [1:SignalsCount] and OscCounts index is [0:DataBooksCount-1]
	    								if ((TempDS[r] > OscCounts[DataBookNumbers[r]-1]) || (TempDS[r] < 0))
	    								{	ShowMessage("Error at Layer '"+DefaultOptionsLayersPrefix[1]+"' of '"+OptionsBookName+"' WorksheetPage!\n"+DefaultOptionsLabel[i]+"["+(r+1)+"] value is greather than 'OscCount' or less than '0'.");
	    									AccessError = true;
	    								}
	    							}
	    					//PeakCount value check
	    							else if (i == 3)		//DefaultOptionsLabel[3] ~= "Peaks Count"
	    							{	int temp = TempDS[r];
	    								if (temp < 1)
	    								{
	    									TempDS[r] = 1;
	    									ShowMessage("Warning!\nPeakCount has incompatible value.\nValue changed to '1'");
	    								}
	    							}
	    					//PeakDifference value check
	    							else if (i == 4)		//DefaultOptionsLabel[4] ~= "Peak Difference"
	    							{	double d;
	    								d = TempDS[r];
	    								if (d < 0)		{	TempDS[r] = abs(d);	}
	    							}
	    					//Rise and Fall front treshold values check
	    							else if ((i == 5) && (i == 6) && (i == 7) && (i == 8)) 	// Front Rise and Fall
	    							{	double d;
	    								d = TempDS[r];
	    								if ((d > 1) || (d < 0))
	    								{ 	ShowMessage("Error at Layer '"+DefaultOptionsLayersPrefix[1]+"' of '"+OptionsBookName+"' WorksheetPage!\n"+DefaultOptionsLabel[i]+"["+(i+1)+"] value is greather than '1' or lesser than '0'.");
	    									AccessError = true;
	    								}
	    							}
	    					//HalfWidth value check
	    							else if (i == 9)				//DefaultOptionsLabel[9] ~= "Half Width"
	    							{
	    								if ((TempDS[r] != 1) && (TempDS[r] != 0))
	    								{	TempDS[r] = 0;
	    									ShowMessage("Warning!\nHalfWidth has incompatible value.\nValue changed to '0'");
	    								}
	    							}
	    					// Graph needed
	    							else if (i == 13)				//DefaultOptionsLabel[13] ~= "Graph Needed"
	    							{	int temp = TempDS[r];
	    								if ((temp != 1) && (temp != 0))
	    								{
	    									TempDS[r] = 1;
	    									ShowMessage("Warning!\nGraphNeeded has incompatible value.\nValue changed to '0'");
	    								}
	    							}
	    					// NOISE FILTER
									else if (i == 12)				//DefaultOptionsLabel[12] ~= "Noise Filter"
	    							{	
	    								int temp = TempDS[r];
	    								if ((temp < 0) || (temp == NANUM))
	    								{
	    									TempDS[r] = 0;
	    									ShowMessage("Warning!\nNoiseFilter has incompatible value.\nValue changed to '0'");
	    								} 
	    								FilterExtent = TempDS[r];	
	    							}
	    						}
							}
							
							TempDS.Detach();
							
						}
						else
						{	ShowMessage("Error!\n'"+DefaultOptionsLabel[i]+"' column of layer '"+DefaultOptionsLayersPrefix[1]+"' from "+OptionsBookName+" worksheet Page is missing!\n");
						AccessError = true;		}
					}
				}
	    	}
		}//'Settings' layer's columns check
	
		
  //------------------------------------------------------------------------------------------
 // Complex Error Search
//-----------------------------

	    		if (!AccessError)
	    		{
	    			
	    		//lower and upper point of Rise front check
	    			Dataset TempDS2;
	    			TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[5]));			// Column[5] ~= "Rise Front Treshold (lower)"
	    			TempDS2.Attach(SettingsWks.Columns(DefaultOptionsLabel[6]));		// Column[6] ~= "Rise Front Treshold (upper)"
	    			for (int i=0; i<SignalsCount; i++)
	    			{
	    				string dss1, dss2;
	    				TempDS.GetText(i, dss1);
	    				TempDS2.GetText(i, dss2);
	    				dss1.TrimRight("mM");
	    				dss2.TrimRight("mM");
	    				
	    				if (atof(dss1, false) == atof(dss2, false))
	    				{
	    					ShowMessage("Error!\nRiseTrehold1 is equal to RiseTreshold2.");
	    					AccessError = true;
	    				}
	    			}
	    			TempDS.Detach();
	    			TempDS2.Detach();
	    			
	    		//lower and upper point of Fall front check
	    			TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[7]));			// Column[7] ~= "Fall Front Treshold (upper)"
	    			TempDS2.Attach(SettingsWks.Columns(DefaultOptionsLabel[8]));		// Column[8] ~= "Fall Front Treshold (lower)"
	    			for (i=0; i<SignalsCount; i++)
	    			{
	    				string dss1, dss2;
	    				TempDS.GetText(i, dss1);
	    				TempDS2.GetText(i, dss2);
	    				dss1.TrimRight("mM");
	    				dss2.TrimRight("mM");
	    				
	    				if (atof(dss1, false) == atof(dss2, false))
	    				{
	    					ShowMessage("Error!\nFallTrehold1 is equal to FallTreshold2.");
	    					AccessError = true;
	    				}
	    			}
	    			TempDS.Detach();
	    			TempDS2.Detach();
	    			
	    		//start and stop 'X' values check
	    			TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[10]));			// Column[10] ~= "Start X" (left bound)
	    			TempDS2.Attach(SettingsWks.Columns(DefaultOptionsLabel[11]));			// Column[11] ~= "Stop X" (right bound)
	    			for (i=0; i<SignalsCount; i++)
	    			{
	    				ProcessDirection = 1;
	    				if (TempDS[i] < TempDS2[i])	{	ProcessDirection = -1;	}
	    				else if (TempDS[i] == TempDS2[i]) 
	    				{	
	    					ShowMessage("Error!\nStart is equal to Stop.");
	    					AccessError = true;
	    				}
	    			}
	    			TempDS.Detach();
	    			TempDS2.Detach();
	    			
	    	//attaching data objects
	    			//WorksheetPage DataBook(DataBookNames[0]);
	    			int ShootsCount = Project.WorksheetPages(DataBookNames[0]).Layers.Count()/OscCounts[0];
	    			
	    		//Data Columns existence check (by ColumnName)
	    			if (!AccessError)
	    			{
	    				TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[0]));			// Column[0] ~= "Osc Number"
			    		StringArray NamesSA;
			    		SettingsWks.Columns(DefaultOptionsLabel[1]).GetStringArray(NamesSA);		// Column[1] ~= "Column Names"
			    		
	    				for (int q=0; q<ShootsCount; q++)
	    				{
	    					
			    			Worksheet wks;
	    					for (i=0; i<SignalsCount; i++)
	    					{											// DataBookNumbers = [1:SignalsCount] and DataBookNames index is [0:SignalsCount-1]
		    					wks = Project.WorksheetPages(DataBookNames[DataBookNumbers[i]-1]).Layers(q*OscCounts[DataBookNumbers[i]-1] + TempDS[i] - 1);
		    					if (!wks.Columns(NamesSA[i]))
	    						{
		    						ShowMessage("Error!\nColumn '"+NamesSA[i]+"' at Layer '"+wks.GetName()+"' of DataBook '"+DataBookNames[DataBookNumbers[i]-1]+"' does not exist.");
		    						AccessError = true;
		    						wks.Detach();
		    						break;
	    						}
	    						wks.Detach();
	    					}	
	    					if (AccessError) {break;}
	    				}
	    				TempDS.Detach();
	    			}// Data Column check
	  
	  //-----------------------\\	
	 // AccessErrorSearch_END	\\
	//---------------------------\\
   //   Process Start			  \\		
  //=====================================================================================================================================
			
	    			if (!AccessError)
	    			{
		    			WorksheetPage OutputDataWP();
		    			OutputDataWP.Create("origin");
		    			//OutputDataWP.Rename(OutputDataWPName); //WorksheetPage.Rename(string s) - unstable function
		    			OutputDataWP.SetName(NewDataBookName);
		    			
		    			//creating new clear worksheetpage wothout any layers
		    			for (int i=0; i<OutputDataWP.Layers.Count(); i++)		{	OutputDataWP.Layers(0).Delete();	}
	    		
		    		//filling Output worksheetpage with layers
		    			Dataset OscNumberDS;
		    			OscNumberDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[0]));			// Column[0] ~= "Osc Number"
		    			
		    			StringArray ColumnNameSA;
		    			SettingsWks.Columns(DefaultOptionsLabel[1]).GetStringArray(ColumnNameSA);
		    			
		    	//		Dataset ColumnNameDS;
		    	//		ColumnNameDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[1]));			// Column[1] ~= "Column Names"
		    				
		    			for (int L=0; L<ShootsCount; L++)
		    			{
		    				OutputDataWP.AddLayer();
		    				
		    		/* 
		    				string LayerName;
		    				LayerName = DataBook.Layers(L*OscCount).GetName();
		    				if (PrefixLength > 0)
		    				{
		    					if (PrefixLength < LayerName.GetLength()) {	LayerName = LayerName.Left(PrefixLength);	}
		    				}
		    				// if PrefixLength == 0 then new LayerName will be Old LayerName (from DataBook)
		    				
		    				OutputDataWP.Layers(L).SetName(LayerName);
		    		*/
		    		
				// 			RENAMING layers of new databook
		    				//OutputDataWP.Layers(L).SetName(DataBook.Layers(L*OscCounts[0]).GetName());
		    				OutputDataWP.Layers(L).SetName(Project.WorksheetPages(DataBookNames[0]).Layers(L*OscCounts[0]).GetName());
		    				Worksheet wks = OutputDataWP.Layers(L);
		    				
		    				
		    				for (int s=0; s<SignalsCount; s++)
		    				{
		    					bool GraphNeeded = false;
		    					TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[13])); // column GRAHP NEEDED
		    					if (TempDS[s] == 1) {GraphNeeded = true;}
		    					OutputColumnNames[0] = ColumnNameSA[s];
		    					
		    					for (int c=0; c<10; c++)
		    					{
		    						if (!wks.Columns(s*10+c))	{	wks.AddCol();	}
		    							
		    						wks.Columns(s*10+c).SetLowerBound(0);
		    						wks.Columns(s*10+c).SetUpperBound(1);
		    						if (c % 2 == 0)	
		    							{	
		    								wks.Columns(s*10+c).SetName("X" + (s*10+c/2+1));
		    								wks.Columns(s*10+c).SetType(OKDATAOBJ_DESIGNATION_X);
		    							}
		    						else 
		    							{	
		    								wks.Columns(s*10+c).SetType(OKDATAOBJ_DESIGNATION_Y);
		    								string name;
		    								name = OutputColumnNames[(c-1)/2];
		    								while (wks.Columns(name))
		    								{
		    									name = StringPostfixIncrease(name);
		    								}
		    								wks.Columns(s*10+c).SetName(name);
		    							}
		    					}
		    					
  //----------------------------------------------------------------
 // Attaching DATA columns (X,Y)
//------------------------
																						// DataBookNumbers = [1:SignalsCount] and DataBookNames index is [0:SignalsCount-1]
		    					Worksheet DataWks = Project.WorksheetPages(DataBookNames[DataBookNumbers[s]-1]).Layers(L*OscCounts[DataBookNumbers[s]-1] + OscNumberDS[s] - 1);
		    					Dataset DataX,DataY;
		    					
		    					DataX.Attach(wks.Columns(s*10));
		    					DataY.Attach(wks.Columns(s*10+1));
		    					
		    					int invert;
		    					TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[2])); // Column "Invert"
		    					if ( TempDS[s] == 0) {invert = 1;}
		    						else {invert = -1;}
		    					TempDS.Detach();
		    					
		    					TempDS.Attach(DataWks.Columns(ColumnNameSA[s]));
		    					// Reverseing data
		    					DataY.SetLowerBound(TempDS.GetLowerBound());
		    					DataY.SetUpperBound(TempDS.GetUpperBound());
		    					for (int sell=DataY.GetLowerBound(); sell<=DataY.GetUpperBound(); sell++)
		    					{
		    						DataY[sell] = invert*TempDS[sell];
		    					}
		    					TempDS.Detach();
		    					
		    					int DataXColumnIndex = GetDataXColumnIndex(DataWks, ColumnNameSA[s]);
		    					if (DataXColumnIndex == -1)
		    					{
		    						FillDatasetWthNumRows(DataX,0,DataY.GetUpperBound());
		    					}
		    					else
		    					{
		    						TempDS.Attach(DataWks.Columns(DataXColumnIndex));
		    						DataX.SetLowerBound(TempDS.GetLowerBound());
		    						DataX.SetUpperBound(TempDS.GetUpperBound());
		    						for (sell=DataX.GetLowerBound(); sell<=DataX.GetUpperBound(); sell++)
		    						{
		    							DataX[sell] = TempDS[sell];
		    						}
		    						TempDS.Detach();
		    					}
		    					
		    					DataWks.Detach();
		    					
		    					
		    					
  //---------------------------------------------------------------------------------------------------------------------
 // Filter Function
//-----------------------
								if (FilterExtent > 0)
								{
									vector<double> Vdata;
									Vdata = NoiseFiltration(DataY, FilterExtent);
									wks.Columns(s*10).SetUpperBound(Vdata.GetSize()-1);
									wks.Columns(s*10+1).SetUpperBound(Vdata.GetSize()-1);
									for (int i=0; i<Vdata.GetSize(); i++)
									{
										DataY[i] = Vdata[i];
									}
									
								}

  //------------------------------------------------------------------
 //		GET PARAMETERS
//------------------------
		    					bool reverse;
		    					reverse = false;
		    					//	TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[??])); // Column "Reverse"
		    					//	if ( TempDS[s] == 0) {reverse = false;}
		    					//		else {reverse = true;}
		    					//	TempDS.Detach();
		    						
		    						
		    						
		    					Front& FT = FrontConstructor();
		    					string tempstr;
	// FRONTs TRESHOLDs
		    					TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[5]));		// Column[5] ~= "Rise Front Treshold (lower)"	
		    					TempDS.GetText(s, tempstr);
		    					if ((tempstr.Right(1) == "m") || (tempstr.Right(1) == "M"))
		    					{
		    						tempstr.TrimRight("mM");
		    						FT.RiseLowIndex = 1;
		    					}
		    					else {	FT.RiseLowIndex = 0;	}
		    					FT.RiseLow = atof(tempstr, false);
		    					TempDS.Detach();
		    					
		    					
		    					TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[6]));		// Column[6] ~= "Rise Front Treshold (upper)"
		    					TempDS.GetText(s, tempstr);
		    					if ((tempstr.Right(1) == "m") || (tempstr.Right(1) == "M"))
		    					{
		    						tempstr.TrimRight("mM");
		    						FT.RiseUpIndex = 1;
		    					}
		    					else {	FT.RiseUpIndex = 0;	}
		    					FT.RiseUp = atof(tempstr, false);
		    					TempDS.Detach();
		    					
		    					
		    					TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[7]));		// Column[7] ~= "Fall Front Treshold (upper)"
		    					TempDS.GetText(s, tempstr);
		    					if ((tempstr.Right(1) == "m") || (tempstr.Right(1) == "M"))
		    					{
		    						tempstr.TrimRight("mM");
		    						FT.FallUpIndex = 1;
		    					}
		    					else {	FT.FallUpIndex = 0;	}
		    					FT.FallUp = atof(tempstr, false);
		    					TempDS.Detach();
		    					
		    					
		    					TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[8]));		// Column[8] ~= "Fall Front Treshold (lower)"
		    					TempDS.GetText(s, tempstr);
		    					if ((tempstr.Right(1) == "m") || (tempstr.Right(1) == "M"))
		    					{
		    						tempstr.TrimRight("mM");
		    						FT.FallLowIndex = 1;
		    					}
		    					else {	FT.FallLowIndex = 0;	}
		    					FT.FallLow = atof(tempstr, false);
		    					TempDS.Detach();
		    					
	//	HALF-WIDTH [yes/no]
		    					TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[9]));		// Column[9] ~= "Half Width" [yes/no]
		    					bool HalfWidth = false;
		    					if (TempDS[s] == 1) {HalfWidth = true;}
		    					TempDS.Detach();
		    					
	//		START-X
		    					double StartX, StopX;
		    					int StartIndex, StopIndex;
		    					TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[10]));	// Column[10] ~= "Start X" (left bound)
		    					StartX = TempDS[s];
		    					TempDS.Detach();
		    					if (StartX < DataX[DataX.GetLowerBound()]) 	
		    					{	
		    						StartX = DataX[DataX.GetLowerBound()];
		    						StartIndex = DataX.GetLowerBound();
		    					}
		    					else
		    					{	
		    						for (int i=DataX.GetLowerBound(); i<=DataX.GetUpperBound(); i++) 
		    						{
		    							if (StartX <= DataX[i]) 
		    							{
		    								StartIndex = i;
		    								break;
		    							}
		    						}
		    					}
		    					
	//		STOP-X
		    					TempDS.Attach(SettingsWks.Columns(DefaultOptionsLabel[11]));	// Column[11] ~= "Stop X" (right bound)
		    					StopX = TempDS[s];
		    					TempDS.Detach();
		    					if ((StopX > DataX[DataX.GetUpperBound()]) && (DataX[DataX.GetUpperBound()] != NANUM))
		    					{	
		    						StopX = DataX[DataX.GetUpperBound()];
		    						StopIndex = DataX.GetUpperBound();
		    					}
		    					else
		    					{
		    						for (i=DataX.GetUpperBound(); i>=DataX.GetLowerBound(); i--) 
		    						{
		    							if ((StopX >= DataX[i]) && (DataX[i] != NANUM))
		    							{
		    								StopIndex = i;
		    								break;
		    							}
		    						}
		    					}
		    					
		    					
    //===============================================================================\\	    					
   //						CALCULATIONs   											  \\
  //-----------------------------------------------------------------------------------\\
 //		PEAK
//--------------------------		    					
		    					double MaxY,MaxX,MaxI;
		    					MaxY = GetMaxDSValue(DataY, StartIndex, StopIndex);
		    					MaxI = GetDSMaxValIndex(DataY, MaxY, StartIndex, StopIndex, reverse);
		    					MaxX = DataX[MaxI];
		    					
		    					TempDS.Attach(wks.Columns(s*10+2));
		    					TempDS.SetUpperBound(0);
		    					TempDS[0] = MaxX;
		    					TempDS.Detach();
		    					
		    					TempDS.Attach(wks.Columns(s*10+3));
		    					TempDS.SetUpperBound(0);
		    					TempDS[0] = MaxY;
		    					TempDS.Detach();
		    					
  //-----------------------------------------------------------------------------------
 //		FRONT
//--------------------------
		    					
		    					Front& F = FrontConstructor();
		    					Front& FrontFall = FrontConstructor();
		    					double UpBound, LowBound;
		    					// if index==1 then Treshhold == % of maximum [in scale of 1]
		    					// if index==0 then Treshhold == final value [double]
		    					
		    					if (FT.RiseLowIndex == 0) 	{LowBound 	= FT.RiseLow;}
		    					else 						{LowBound 	= FT.RiseLow*MaxY;}
		    					if (FT.RiseUpIndex == 0) 	{UpBound 	= FT.RiseUp;}
		    					else 						{UpBound 	= FT.RiseUp*MaxY;}
		    							
		    					F 			= FrontCalc(DataY,MaxI,StartIndex,StopIndex,LowBound,UpBound,false);
		    					
		    					if (FT.FallLowIndex == 0) 	{LowBound 	= FT.FallLow;}
		    					else 						{LowBound 	= FT.FallLow*MaxY;}
		    					if (FT.FallUpIndex == 0) 	{UpBound 	= FT.FallUp;}
		    					else 						{UpBound 	= FT.FallUp*MaxY;}
		    						
		    					FrontFall 	= FrontCalc(DataY,MaxI,StartIndex,StopIndex,LowBound,UpBound,true);
		    					
		    					F.FallUp 		= FrontFall.RiseUp;
		    					F.FallLow 		= FrontFall.RiseLow;
		    					F.FallUpIndex 	= FrontFall.RiseUpIndex;
		    					F.FallLowIndex 	= FrontFall.RiseLowIndex;
		    					
		    					// writeing results
		    					TempDS.Attach(wks.Columns(s*10+4));
		    					TempDS[0] = DataX[F.RiseLowIndex];
		    					TempDS[1] = DataX[F.RiseUpIndex];
		    					TempDS.Detach();
		    					
		    					TempDS.Attach(wks.Columns(s*10+5));
		    					TempDS[0] = F.RiseLow;
		    					TempDS[1] = F.RiseUp;
		    					TempDS.Detach();
		    					
		    					TempDS.Attach(wks.Columns(s*10+6));
		    					TempDS[0] = DataX[F.FallLowIndex];
		    					TempDS[1] = DataX[F.FallUpIndex];
		    					TempDS.Detach();
		    					
		    					TempDS.Attach(wks.Columns(s*10+7));
		    					TempDS[0] = F.FallLow;
		    					TempDS[1] = F.FallUp;
		    					TempDS.Detach();
		    					
		    					
  //-------------------------------------------------------------------------------------
 //		HALF-WIDTH
//----------------------
		    					
//		    					DoublePoint& FR = DoublePointConstructor();
//		    					DoublePoint& FF = DoublePointConstructor();
//		    					FR.x1 = DataX[F.RiseLowIndex];
//		    					FR.y1 = F.RiseLow;
//		    					FR.x2 = DataX[F.RiseUpIndex];
//		    					FR.y2 = F.RiseUp;
//		    					FF.x1 = DataX[F.FallLowIndex];
//		    					FF.y1 = F.FallLow;
//		    					FF.x2 = DataX[F.FallUpIndex];
//		    					FF.y2 = F.FallUp;
		    					
		    					
		    					DoublePoint& HW = HalfWidthCalc(DataX, DataY, MaxI);
		    					
		    					TempDS.Attach(wks.Columns(s*10+8));
		    					TempDS[0]= HW.x1;
		    					TempDS[1]= HW.x2;
		    					TempDS.Detach();
		    					
		    					TempDS.Attach(wks.Columns(s*10+9));
		    					TempDS[0]= HW.y1;
		    					TempDS[1]= HW.y2;
		    					TempDS.Detach();
		    				
  //-----------------------------------------------------------------
 //		RE-REVERSE
//--------------------------
		    					
		    					if (invert == -1)
		    					{
		    						//	for (int sell=DataY.GetLowerBound(); sell<=DataY.GetUpperBound(); sell++)
		    						//	{
		    						//		DataY[sell] = -DataY[sell];
		    						//	}

		    						for (c=1; c<=9; c=c+2)
		    						{
		    							TempDS.Attach(wks.Columns(s*10+c));
		    							for (int sell=TempDS.GetLowerBound(); sell<=TempDS.GetUpperBound(); sell++)
		    							{
		    								TempDS[sell] = -TempDS[sell];
		    							}
		    						}
		    						
		    					}
		    					
  //===========================================================================
 //			GRAPH 
//=========================
		    		  			
		    		  			if (GraphNeeded) 
		    		  			{
		    		  				SA_GraphProcess(wks, s*10+1, 5, StartX, StopX);
		    		  			}
		    					
		    				} // for all signals(rows) in SettingsWks
		    				
		    			} // for all Layers in DataBook
		    			
		    			if (!AccessError) 
						{
							string name = NewDataBookName;
							name.TrimRight("0123456789");
							while (Project.WorksheetPages(name))
							{
								name = StringPostfixIncrease(name);
							}
							SignalFrontAndDuration(NewDataBookName, name, SignalsCount);
							CommaToDotReplacement(name, 0);
						}
	    			}// Main Process
	    
		    	}//accessError
		    	else {ShowMessage("Error!\nProcess stopped!\n");}
			
	}//'OptionsBook' check
	else {ShowMessage("Error!\n Can't connect to table 'Options'!\n");}
}
// SignalAnalysisProcess_END
//==================================================================================================================================
