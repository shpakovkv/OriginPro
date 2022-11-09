Usefull hints
=============

#### Text Label Options

You can use it to automatically label graph axes.

How to find help: Help -> Programming -> LabTalk, then search "Text Label Options".
Alternative: Help -> Programming -> LabTalk, then search "Substitution Notation", then look for a link to "Text Label Options" on the page under Worksheet Information Substitution table.

Example:
* **`%(1,@LA), %(1,@LU)`** will be converted to **"Y_Long_Name, Y_Units"**
* **`%(1X,@LA), %(1X,@LU)`** will be converted to **"X_Long_Name, X_Units"**

*For* `@LA` *options: if there is no LongName, then ShortName will be used instead*
