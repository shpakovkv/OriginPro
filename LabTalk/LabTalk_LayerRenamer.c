/*
The script changes workbook's layers names for specified WorksheetPages. There are two scenario:

1. Fills layers names with zero padded layer number starting from 1. Or you may specify increment, and layers will be numbered from the specified number (increment + 1). 
For this scenario set grNames = {""};

2. You can set the numbering of subgroups of layers by specifying grNames array.
Then the script will number the first N layers with the same number X and adds a certain postfix to it (different for each layer). 
Then the script will number the second N layers with X+1 number (with different postfix), etc.
Where N - number of layer in each group == size of grNames array
X - group number (first group number == increment + 1)

WARNING! The script will process all layers of the WorksheetPages (WorkBooks) specified in bookNames array.

How to use:
1) Specify the names of books (worksheet pages) in the bookNames array.
2) Specify the increment (the first layer/group will be renamed as increment + 1, the second = increment + 2, etc.)
3) Specify the number of digits for the layer number. The layer/group number will be zero padded number.
4) Specify the names of group members (layers) in the grNames array OR set grNames = {""};
5) Paste the script to the Command Window of an OriginPro project and press Enter.
*/
//=======================================
// SETUP

// the list of the WorksheetPages whose layers need to be renamed
StringArray bookNames = {"Book1", "Book2"};  

// by default the number of the first layer/group is 1
// but you can increase/decrease it by this parameter
int increment = 0;

// the name of the layers will be a number with leading zeros 
// with this total number of digits
int layerNumberDigits = 4;  

// group members (layers) names
StringArray grNames = {"CH1", "CH2", "CH3"};  

//=======================================
// initialization
string name$ = "";
string memberName$ = "";  // group member (layer) name
int numberOfChannels = grNames.GetSize();

// main loop
for (bookNumber = 1; bookNumber<=bookNames.GetSize(); bookNumber++)
{
	window -a %(bookNames.GetAt(bookNumber)$);
	int count = page.nLayers;	
	for (layerNumber = 1; layerNumber<=count / numberOfChannels; layerNumber++)
	{
		for (grNumber = 1; grNumber<=numberOfChannels; grNumber++)
		{
			page.active = (layerNumber - 1) * numberOfChannels + grNumber;
			name$ = Format(layerNumber + increment, "#$(layerNumberDigits)")$;
			memberName$ = grNames.GetAt(grNumber)$;
			layer.name$ = "%(name$)%(memberName$)";
			// type "layerNumber = $(layerNumber); grNumber = $(grNumber); active page = $(page.active)"  // debug
		}
	}
}
