StringArray aa;       
aa.Add("GOptions1");	
aa.Add("GOptions2");
aa.Add("GOptions3");
aa.Add("GOptions4");
aa.Add("GOptions5");
aa.Add("GOptions6");

for(ii=1; ii<=aa.GetSize(); ii++)
{
window -a %(aa.GetAt(ii)$);
window -r %(aa.GetAt(ii)$) GOptions;
ProcessGraph();
window -r GOptions %(aa.GetAt(ii)$);
}