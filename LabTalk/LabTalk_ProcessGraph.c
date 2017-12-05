/* Script for multiple start of plotting processing (by ProcessGraph)
Use it, if you have multiple GOptions tables with parameters, and want to plot all graphs at once.

How to use
----------

1. Add ProcessGraph.c to your project in Code Builder and compile all.
2. Add ProcessGraph.opj to your project.
3. Create template graph and configure graph parameters at GOptions table (see ProcessGraph manual for details). 
4. If you need to plot several series of graphs with different parameters repeat this steps as many times as you need:
    * make new teamplate graph, 
    * duplicate GOptions table, 
    * fill that table with new values
5. Add short names of GOptions tables to the string array TNames.
    ..note: the project should not have a window named "GOptions".
6. Copy this script, paste it to Command Window of OriginPro, and press Enter.


Troubleshooting
---------------

If the process stops with an error due to incorrect parameters, 
during the process one of the parameter tables can be named "GOptions", 
which means that a new script can not be started. 
Rename the table as you specified in the string array TNames, before starting script again.

*/

// PARAMETERS
StringArray TNames;       
TNames.Add("GOptions1");	
TNames.Add("GOptions2");
TNames.Add("GOptions3");
TNames.Add("GOptions4");
TNames.Add("GOptions5");
TNames.Add("GOptions6");

// -------------------------------------
for(ii=1; ii<=TNames.GetSize(); ii++)
{
window -a %(TNames.GetAt(ii)$);
window -r %(TNames.GetAt(ii)$) GOptions;
ProcessGraph();
window -r GOptions %(TNames.GetAt(ii)$);
}