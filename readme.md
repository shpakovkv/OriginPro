
SignalProcess project
=====================

The program is intended for initial processing and analysis of experimental data (digital signals of detectors).

The project consists of the following tools:
* **ProcessSignals** - initial data processing
* **ProcessGraph** and **AllInOneGraph** - automatization of plotting process
* **SignalAnalysisProcess** - search for peaks
* **Process_FFT** - automation of calculation of FFT of different parts of the signal
* **Optical_Filters** - spectral characteristics of color optical glass (GOST standart 9411-81)
* **TempScript_v08.c** - temporary file 

All parts are tested and working properly (except TempScript).  
All full manuals are in the "Manuals" folder.

How to use OriginC scripts:
1. Open a project with data.
2. Append .opj file of the tool you need to your project.
3. In the table that appears, enter all the required parameters.
4. Open *Code Builder*, open .c file of the tool you need and add it to the *Workspace*.
5. Compile all in the *Code Builder*.
6. Return to the OriginPro project and use the appropriate script to start the process or just click the "Process" button in the tool parameters table (this button will not work if you changed the short name of the parameters table).

How to use LabTalk scripts:
1. Change the values of the script parameters (in the parameters section) according to your needs (follow the instructions in the comments).
2. Paste changed script to the *Command Window* in your OriginPro project and press Enter.

---

ProcessSignals:
---------------

The script is intended for initial processing of experimental data (digital signals of detectors).
It allows to multiply the data (samples by time and amplitude axes) by the specified multiplier and subtract the specified time delay independently for each signal. And also the script allows you to set the offset value along the time axis independently for each *shot*. A *shot* is one record (one result), which can consist of several signals and even files, including different formats, due to the use of several different recording devices during the experiment.


Consists of: 
* **ProcessSignals.c** - OriginC script
* **ProcessSignals.opj** - OriginPro project with tables for entering parameters.

Usefull LabTalk scripts:
* **LabTalk_ProcessSignals.c**
* **LabTalk_LayerRenamer.c**
* **LabTalk_LayersNameCutter.c**

ProcessGraph
------------

\**description in development*\*

Consists of: 
* **ProcessGraph.c** - OriginC script
* **ProcessGraph.opj** - OriginPro project with tables for entering parameters.

AllInOneGraph
-------------

\**description in development*\*

Consists of: 
* **AllInOneGraph.c** - OriginC script 
* **LabTalk_GraphAllInOne_single.c** - LabTalk editable script for entering parameters (for one graph)
* **LabTalk_GraphAllInOne_multiple.c** - LabTalk editable script for entering parameters (for several different graphs)

SignalAnalysisProcess
---------------------

\**description in development*\*

Process_FFT
-----------

\**description in development*\*

---

LabTalk scripts
---------------

\**description in development*\*

Includes:
* **LabTalk_3in1Script.c**                
* **LabTalk_ColRename.c**                 
* **LabTalk_ColRename_RowFill.c**         
* **LabTalk_CopyLayersFromBooks.c**       
* **LabTalk_FillColumnsWithSequence.c**   
* **LabTalk_GetLayersToMultipleBooks.c**  
* **LabTalk_GetLayersToOneNewBook.c**     
* **LabTalk_GraphAllInOne_multiple.c**    
* **LabTalk_GraphAllInOne_single.c**
* **LabTalk_InvertColumnValues.c**
* **LabTalk_LayerRenamer.c**
* **LabTalk_LayersNameCutter.c**
* **LabTalk_ProcessGraph.c**
* **LabTalk_ProcessSignals.c**
* **LabTalk_SaveLayersToCSVs.c**
* **LabTalk_SwapTwoRows.c**

---

Optical_Filters
---------------

OriginPro project contains table with spectral characteristics of various color optical glass from GOST standart 9411-81 list.
Allow you to easily create spectral absorption or transmission graphs for a set of several different filters of a given thickness.