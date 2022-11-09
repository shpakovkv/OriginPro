Usefull hints
=============

#### Text Label Options

You can use it to automatically label graph axes. 

How to find help: Help -> Programming -> LabTalk, then search "Text Label Options". 
Alternative: Help -> Programming -> LabTalk, then search "Substitution Notation", then look for a link to "Text Label Options" on the page under Worksheet Information Substitution table.

Example:
```
**"%(1,@LA), %(1,@LU)"** => **"Y_Long_Name, Y_Units"** *if there is no LongName then ShortName will be used*
**"%(1X,@LA), %(1X,@LU)"** => **"X_Long_Name, X_Units"** *if there is no LongName then ShortName will be used*
```
