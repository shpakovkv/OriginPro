//Discription:
// Скрипт создает для каждого из выбранных (см. переменную ExShoots) выстрелов отдельные Workbook'и, на каждом слое которых расположены данные из определенного исходного Workbook'а (см. переменную aa). 

//Необходимо задать три переменные:
// aa - массив имен WorkBook'ов
// ExShoots - массив номеров выстрелов, которые нас интересуют
// NamePrefix$ - префикс имени будующих WorkBook'ов

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
ss$ = NamePrefix$ + "0000";	// create Name of new Workbook
	//type("Creating Book");
newbook name:=ss$ sheet:=0 option:=lsname;	// create new WorkBook with 0 Worksheets and with Short Name = ss$
	// now active window is newbook
for (ii=1; ii<=ExShoots.GetSize(); ii++)  
{
    //type "$(ii)";
	
	
	for (jj=1; jj<=aa.GetSize(); jj++)
	{
		window -a %(aa.GetAt(jj)$);				// activating DataBook
		page.active = ExShoots(ii);				// activating Data worksheet
		StringArray ColNames;
		for (cc=1; cc<=wks.nCols; cc++)
		{
			wks.col = cc;						// activating data column
			ColNames.Add(wks.col.name$);		// getting column Short Names
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
		
		page.active = jj + (ii-1)*aa.GetSize();								// activating worksheet
		for (cc=1; cc<=wks.nCols; cc++)
		{
			wks.col = cc;								// activating column
			wks.col.name$ = ColNames.GetAt(cc)$;		// setting column Short Names
		}
	}

}
type "Done.";