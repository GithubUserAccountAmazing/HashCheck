![image](https://github.com/originates/HashCheck-and-Run/blob/main/hashcheck.png)

This is a program that leverages AutoHotkey (AHK), PowerShell and VBScript to securely open a single-instance Excel file with a verified SHA256 hash. The program bypasses the user macro settings and executes a Workbook_Open VBA script embedded in the Excel file.

This program has some interesting features and benefits:

- It opens the excel file as a template of the original xlsm file, which allows you to run VBA macros regardless of your macro security settings. This can be useful when using excel as a 'display manager' for a program and you don't want the user to have access to standard excel functionality. Also this prevents a user from modifying the Excel file.
- It checks the sha256 of the excel file before opening it, which ensures that the file has not been tampered with or corrupted.
- It imports a certificate to the TrustedPublisher store, which prevents any security warnings or prompts when running the VBA macros.

You may be concerned that allowing users to override user macro settings is a security risk; however, many organizations grant their users full control over VBA. This tool allows you to restrict users to specific workbooks and their associated VBA scripts while still keeping Macros disabled. This program is not a perfect solution, but it is more secure than allowing users to run any macro and also ensures that users can trust the source and integrity of the Excel file due to the SHA256 hash check.

## Installation

To use this program, you need to have AHK, PowerShell, and VBscript installed on your system. You also need to have an excel file with a Workbook_Open vba script and a certificate file.

To install this program, follow these steps:

1. Download the HashCheck-And-Run.ahk file from this repository and save it in a folder of your choice.
2. Edit the HashCheck-And-Run.ahk file with a text editor and change the following variables according to your needs:

    - fileLocation: The path to your excel file
    - fileHash: The sha256 hash of your excel file
    - SplashImage: The path to an image file that will be displayed as a splash screen while the program runs
    - certLocation: The path to your certificate file

3. Save the HashCheck-And-Run.ahk file,
4. Open the Start Menu and go to the Apps list and go to AutoHotKey in the Apps’ list, and select Convert .ahk to .exe.
5. Click Convert

    - Optional: Use a custom icon for the executable to distiguish the program you are creating.
 
6. Run the executable

## Usage

Once you run the HashCheck-And-Run.ahk file, you will see a splash screen with an image of your choice. The program will then perform the following actions:

- It will use certUtil.exe to check the sha256 hash of your excel file and compare it with the one you specified in the fileHash variable. If they match, it will proceed to the next step. If they don't match, it will display an error message and exit.
- It will use Import-Certificate to import your certificate file to the TrustedPublisher store of your current user. This will allow you to run your VBA macros without any security warnings or prompts.
- It will use MSScriptControl.ScriptControl.1 to create an object that can execute VBscript code. It will then create an Excel application object, disable its alerts and visibility, and open your excel file as a template. It will then make the Excel application visible so you can see and interact with it.

At this point, your excel file should be opened and ready to run its Workbook_Open vba script.


