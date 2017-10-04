StringArray aa;       
aa.Add("DOptions6");	
aa.Add("DOptions7");
aa.Add("DOptions8");
aa.Add("DOptions9");
aa.Add("DOptions10");
aa.Add("DOptions11");

for(ii=1; ii<=aa.GetSize(); ii++)
{
window -a %(aa.GetAt(ii)$);
window -r %(aa.GetAt(ii)$) DOptions;
ProcessSignals();
window -r DOptions %(aa.GetAt(ii)$);
}