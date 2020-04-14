// SETUP
StringArray booksWithOptions = {"DOptions1", "DOptions2", "DOptions3"};

// PROCESS
for(ii=1; ii<=booksWithOptions.GetSize(); ii++)
{
window -a %(booksWithOptions.GetAt(ii)$);
window -r %(booksWithOptions.GetAt(ii)$) DOptions;
ProcessSignals();
window -r DOptions %(booksWithOptions.GetAt(ii)$);
}