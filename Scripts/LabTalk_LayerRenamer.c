StringArray aa;       
aa.Add("DPO7054");
aa.Add("HMO3004");
aa.Add("TDS2024C");	
aa.Add("LeCroy");
int increment = 0;

for (bb = 1; bb<=aa.GetSize(); bb++)
{
	window -a %(aa.GetAt(bb)$);
	int count = page.nLayers;	
	for (ii = 1; ii<=count; ii++)
	{
		page.active = ii;
		layer.name$ = Format(ii+increment, "#4")$;
	}
}