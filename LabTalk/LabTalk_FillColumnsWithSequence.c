/*
The script fills the sapecified (by index) column on all layers 
of the specified (by short name) Books with the DataSequence 
(a sequence of numbers).

How to use:
1) Specify the Book(s) name(s).
2) Specify the column index (NOTE: first column index is 1)
3) Specify the sequence of numbers.
4) Paste the script to the Command Window and press Enter.
*/

// add the short names of the target books
StringArray BookNames;
BookNames.Add("PeaksSeries1");
BookNames.Add("PeaksSeries2");
BookNames.Add("PeaksSeries3");
BookNames.Add("PeaksSeries4");
BookNames.Add("PeaksSeries5");
BookNames.Add("PeaksSeries6");
BookNames.Add("PeaksSeries7");

// index of column to fill with sequence (NOTE: first column index is 1)
int ColNumber = 1;  

// the sequence of numbers for column filling
Dataset DataSequence = {0, 10, 20, 25, 30, 40, 50, 60, 70, 80, 90};

for (iBook = 1; iBook<=BookNames.GetSize(); iBook++)
{
	window -a %(BookNames.GetAt(iBook)$);
	count = page.nLayers;
	for (iLayer = 1; iLayer<=count; iLayer++)
	{
		page.active = iLayer;
		wcol(ColNumber) = DataSequence;
	}
}