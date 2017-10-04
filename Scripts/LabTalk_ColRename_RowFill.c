StringArray aa;
aa.Add("PeaksSeries1");
aa.Add("PeaksSeries2");
aa.Add("PeaksSeries3");
aa.Add("PeaksSeries4");

for (bb = 1; bb<=aa.GetSize(); bb++)
{
	window -a %(aa.GetAt(bb)$);
	count = page.nLayers;
	for (ii = 1; ii<=count; ii++)
	{
		page.active = ii;

		// for 11 dedectors:
		// col(1) = {0, 9, 18, 27, 36, 45, 54, 63, 72, 81, 90};

		// for 10 detectors
		col(1) = {0, 10, 20, 30, 40, 50, 60, 70, 80, 90};
		
		wks.col2.name$ = "PeakX";
		wks.col3.name$ = "PeakY";
		wks.col4.name$ = "sqrL";
		wks.col5.name$ = "sqrR";
		wks.col6.name$ = "sqrSum";
		rowNumber = wks.col3.nRows;
		for (iRow = 1; iRow <= rowNumber; iRow++)
		{
			Cell(iRow, 3) = -Cell(iRow, 3);
		}
		double xx = Cell(4, 2);
		double yy = Cell(4, 3);
		double xx2 = Cell(5, 2);
		double yy2 = Cell(5, 3);
		Cell(4, 2) = xx2;
		Cell(4, 3) = yy2;
		Cell(5, 2) = xx;
		Cell(5, 3) = yy;
	}
}