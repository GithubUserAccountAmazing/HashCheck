![image](https://github.com/originates/HashCheck-and-Run/blob/main/hashcheck.png)

----

This uses AHK to call PowerShell to call VBscript to open a single instance excel file.

An interesting property of this script is that the excel file [excel.xlsm] will be opened as [excel1.xlsx]. This quirck it will allow you to run VBA macros regardless of your macro security settings. 
This can be useful when using excel as a 'display manager' for a program and you don't want the user to have access to standard excel functionality.

Since this opens a template of the original xlsm file—if you need to save data consider using vba to write the data to a seperate csv file and use power query within the excel file to read the csv file.

Very convoluted and very niche, but this set up is very powerful when used correctly.


