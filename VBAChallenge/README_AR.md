# VBA-challenge
By. A. Rijpkema 10/16/2019

In this folder you'll find the script used to generate screenshots VBAChallenge 2014, VBAChallenge 2015 and VBAChallenge 2016.
The VBS script knows 3 macros that can be run. 

When running Run_VBAStock all worksheets will be processed.
When running Fill_Summary_Tables only the currently active sheet will be processed.
When running Fill_Min_Max_Tables a second summary table on the active sheet will be filled as well.

Within the Fill_Min_Max_Tables the Create_Summary_Table_Headers() will run again to create and format the tables to be filled.