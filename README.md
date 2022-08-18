![image](https://github.com/originates/HashCheck-and-Run/blob/main/hashcheck.png)

----

This uses AHK to call PowerShell to call VBscript to open a single instance excel file (after verifying the sha256 of the excel file!)

An interesting property of this script is that the excel file [excel.xlsm] will be opened as [excel1.xlsx]. This quirck will allow you to run VBA macros regardless of your macro security settings. 
This can be useful when using excel as a 'display manager' for a program and you don't want the user to have access to standard excel functionality.

Since this opens a template of the original xlsm file—if you need to save data consider using vba to write the data to a seperate csv file and use power query within the excel file to read the csv file.

Very convoluted and very niche, but this set up is very powerful when used correctly.

----

Once you are satisfied with your version of this script use "Convert .ahk to .exe" to (you guessed it) convert the ahk file into a exe file. 

Due to the 'powerful nature' of this script: If you are running this program for other users on a shared network—consider putting this exe file (and excel file) in a place that can only be modified by trusted users. 😉

----


#### Disclaimer: For educational purposes only
The material embodied in this software is provided to you "as-is" and without warranty of any kind, express, implied or otherwise, including without limitation, any warranty of fitness for a particular purpose. In no event shall the GitHub user 'originates' be liable to you or anyone else for any direct, special, incidental, indirect or consequential damages of any kind, or any damages whatsoever, including without limitation, loss of profit, loss of use, savings or revenue, or the claims of third parties, whether or not 'originates' has been advised of the possibility of such loss, however caused and on any theory of liability, arising out of or in connection with the possession, use or performance of this software.


