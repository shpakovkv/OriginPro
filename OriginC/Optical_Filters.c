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


////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.
void CalculateTransmission()
{
	Dataset FilterNames, Thickness, TempDataDataset, TempTargetDataset, L;
	int DataLength, i, j, line;
	string DataTableName("Absorption"), TargetTableName("Transmission"), FilterName;
	double d;
	
	if (FilterNames.Attach("Calc_EnName"))
    {
    	DataLength=FilterNames.GetSize();
    	if (Thickness.Attach("Calc_d"))
    	{
    		if (L.Attach(DataTableName+"_L"))
    		{
    			if (DataLength<=Thickness.GetSize())
    			{
    				matrix<double> Data(1,L.GetSize());
    				for (i=0; i<DataLength; i++)
    				{
    					FilterNames.GetText(i,FilterName);
    					if (TempDataDataset.Attach(DataTableName+"_"+FilterName))
    					{
    						if (TempTargetDataset.Attach(TargetTableName+"_"+FilterName))
    						{
    							d=Thickness[i];
    							line=i+1;
								
    							if (d==0) {printf("Warning! Wrong Thickness (d) in line "+line);}
    							Data=Calc_Transmission(TempDataDataset, d);
    							for (j=0; j<L.GetSize(); j++)
    							{
    								TempTargetDataset[j]=Data[0][j]*100;
    							}
    						}
    						else{printf("Warning! Cannot find '"+FilterName+"' column in Transmission DataTable!");}
    					}
    					else{printf("Warning! Cannot find '"+FilterName+"' filter data!");}
    				}
    				printf("Calculations complete.\n");
    			}
    			else{printf("Error! Enter Thickness (d)!");}
    		}
    		else{printf("Error! Data table loading fail!");}
    	}
    	else{printf("Error! Cannot find 'd' column!\n");}
    }
    else {printf("Error! Cannot find 'Name' column!\n");}
	printf("----------------------\n");
}

void SetThicknessDefault()
{
	Dataset defaultD, D;
	int i;
	
	if (D.Attach("Calc_d"))
	{
		if (defaultD.Attach("Calc_Default"))
		{
			D.SetSize(defaultD.GetSize());
			for(i=0; i<D.GetSize(); i++)
			{D[i]=defaultD[i];}
			printf("Default thickness loaded.\n");
		}
		else {printf("Error! Cannot find 'Default' column!");}
	}
	else {printf("Error! Cannot find 'd' column!");}
	printf("-------------------------\n");
}

matrix<double> Calc_Transmission(Dataset ds, double d)
{
	int i, size; 
	size=ds.GetSize(); 
	
	matrix<double> Result(1,size); 
	for (i=0; i<size; i++)
	{
		Result[0][i]=exp(-d*ds[i]);
	}
	
	return Result; 
}


void MultiFilter()
{
	printf("//Start//");
	Dataset FilterName;
	Dataset Thickness;
	Dataset Result;
	Dataset L;
	string DataTableName("Absorption");
	bool Error=false;
	string temp;
    int i;

    if (FilterName.Attach("MultiFilter_Name"))
    {printf("Filter_Names Loaded\n"); printf("- Number of lines = "+FilterName.GetSize()+"\n");}
    else {printf("Filter_Names loading fail!\n"); Error=true;}
    
    if (Thickness.Attach("MultiFilter_d"))
    {printf("Filter Thikness (d) Loaded\n");}
    else {printf("Error! Filter Thikness (d) loading fail!\n"); Error=true;}
    	
    if (Result.Attach("MultiFilter_Result"))
    {printf("Result column found\n");}
    else {printf("Error! annot find Result column!\n"); Error=true;}
    
    if (L.Attach(DataTableName+"_L"))
    {printf("Absorption DataTabel Loaded:\n");}
    else {printf("Error! Absorption DataTabel loading fail!\n"); Error=true;}
    	
    if (FilterName.GetSize()>Thickness.GetSize())
    {printf("Error! Enter Thickness (d) for all the filters!\n"); Error=true;}
    matrix<double> Data(FilterName.GetSize()+1,L.GetSize());
	
    for(i=0; i<L.GetSize(); i++)
    {Data[FilterName.GetSize()][i]=1;}
    
    if (!Error)
    {
    	Dataset ds;
    	int j, NumberOfFilters=0;
    	for(i=0; i<FilterName.GetSize(); i++)
    	{
    		
    		FilterName.GetText(i,temp);
    		if (ds.Attach(DataTableName+"_"+temp))
    		{
    			double doubleThickness;
    			printf("- Filter '"+temp+"' loaded");
    			matrix<double> TempTransmissionArray(1,ds.GetSize());
    			doubleThickness=Thickness[i];
    			TempTransmissionArray=Calc_Transmission(ds,doubleThickness);
    			if (doubleThickness==0)
    			{printf(" - Warning! Wrong Thikness (d) found!");}
				for(j=0; j<ds.GetSize(); j++)
    			{
    			    Data[i][j]=TempTransmissionArray[0][j];
    				Data[FilterName.GetSize()][j]=Data[FilterName.GetSize()][j]*Data[i][j];
    			}
    			printf("\n");
    			NumberOfFilters++;
    		}
    		else{printf("- Warning! Cannot find '"+temp+"' filter data!\n");}
    	}
    	Result.SetSize(L.GetSize());
    	for(j=0; j<L.GetSize(); j++)
    	{
    		Result[j]=Data[FilterName.GetSize()][j]*100;
    	}
    	printf("Done.\n");
    }
    else{printf("Data loading error!\n");}
    printf("---------------------------\n");
}