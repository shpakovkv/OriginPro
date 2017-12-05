//Discription:
// ������ ������� ��� ������� �� ��������� (��. ���������� ExShoots) ��������� ��������� Workbook'�, �� ������ ���� ������� ����������� ������ �� ������������� ��������� Workbook'� (��. ���������� aa). 

//���������� ������ ��� ����������:
// aa - ������ ���� WorkBook'��
// ExShoots - ������ ������� ���������, ������� ��� ����������
// NamePrefix$ - ������� ����� �������� WorkBook'��

StringArray aa;       // create empty string array
aa.Add("OutDPO7054");	// add DataBook Names to array
aa.Add("OutLeCroy");
aa.Add("OutTDS3054BR");
aa.Add("OutTDS3054BL");
aa.Add("OutTDS2024C");
aa.Add("OutTDS3032C");

Dataset ExShoots={5,14,15,21,27,32,49};	// Exeption shoots indexes that we only need

string NamePrefix$ = "Shoot";			// Workbook Name prefix

//==================================================================
string ss$;		// variable

for (ii=1; ii<=ExShoots.GetSize(); ii++)  
{
    //type "$(ii)";
	
	ss$ = NamePrefix$ + Format(ExShoots(ii), "#4")$;	// create Name of new Workbook
	//type("Creating Book");
	newbook name:=ss$ sheet:=0 option:=lsname;	// create new WorkBook with 0 Worksheets and with Short Name = ss$
	// now active window is newbook
	for (jj=1; jj<=aa.GetSize(); jj++)
	{
		window -a %(aa.GetAt(jj)$);				// activating DataBook
		page.active = ExShoots(ii);				// activating Data worksheet
		StringArray ColNames, ColTypes;
		for (cc=1; cc<=wks.nCols; cc++)
		{
			wks.col = cc;						// activating data column
			ColNames.Add(wks.col.name$);		// getting column Short Name
			if (wks.col.type == 4) 				// getting column Type
			{
				ColTypes.Add("4");
			}
			else 
			{
				ColTypes.Add("1");
			}
		}
		
		//type("Activating Book");
		window -a %(ss$);				// activating newbook
		//type("Creating Sheet");
		string Lname$;
		Lname$ = Format(ExShoots(ii), "#4")$ + aa.GetAt(jj)$;
		newsheet name:=Lname$;	// create new worksheet in newbook

		// adress formula: [WorkBookName]SheetNameOrIndex!ColumnNameOrIndex[CellIndex] 
		type "Copying sheet %(aa.GetAt(jj)$)[$(ExShoots(ii))] to %(ss$)['%(aa.GetAt(jj)$)']";
		wrcopy iw:=[%(aa.GetAt(jj)$)]$(ExShoots(ii))! ow:=[%(ss$)]%(Lname$)! label:=LUC;
		
		page.active = jj;								// activating worksheet
		for (cc=1; cc<=wks.nCols; cc++)
		{
			wks.col = cc;								// activating column
			wks.col.name$ = ColNames.GetAt(cc)$;		// setting column Short Name
			wks.col.type = %(ColTypes.GetAt(cc)$);		// setting column Type
		}
	}

}
type "Done.";