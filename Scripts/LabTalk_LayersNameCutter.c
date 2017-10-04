StringArray aa;       
aa.Add("PeaksSeries1");  // WorksheetPage name
aa.Add("PeaksSeries2");
aa.Add("PeaksSeries3");
aa.Add("PeaksSeries4");
int start_idx = 5;  // one-based index of the first character to be left
// first (start_idx - 1) characters will be deleted

int char_count = 11; // the number of characters to be left (all others will be deleted)

//=======================================================
for (bb = 1; bb<=aa.GetSize(); bb++)
{
	window -a %(aa.GetAt(bb)$);
	int count = page.nLayers;	
	for (ii = 1; ii<=count; ii++)
	{
		page.active = ii;
		string ss$ = layer.name$;
		ss$ = ss.Mid(start_idx, char_count)$;
		layer.name$ = ss$;
	}
}