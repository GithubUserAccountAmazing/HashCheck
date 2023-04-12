; This program is free software: you can redistribute it and/or modify
; it under the terms of the GNU General Public License as published by
; the Free Software Foundation, either version 3 of the License, or
; (at your option) any later version.
;
; This program is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
; GNU General Public License for more details.
;
; You should have received a copy of the GNU General Public License
; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;----------------------------------------------------------------------------------------------
;
;     db   db  .d8b.  .d8888. db   db      .o88b. db   db d88888b  .o88b. db   dD 
;     88   88 d8' `8b 88'  YP 88   88     d8P  Y8 88   88 88'     d8P  Y8 88 ,8P' 
;     88ooo88 88ooo88 `8bo.   88ooo88     8P      88ooo88 88ooooo 8P      88,8P   
;     88~~~88 88~~~88   `Y8b. 88~~~88     8b      88~~~88 88~~~~~ 8b      88`8b   
;     88   88 88   88 db   8D 88   88     Y8b  d8 88   88 88.     Y8b  d8 88 `88. 
;     YP   YP YP   YP `8888Y' YP   YP      `Y88P' YP   YP Y88888P  `Y88P' YP   YD 
;
;
; This script checks the integrity of a file and a certificate using their hashes, 
; and then runs an Excel workbook_open vba macro from the file if they are valid.
; It also displays a splash image when starting the script

; Things to consider:
; Any time you update the excel file you must update the associated SHA256 file hash and recompile.
; If you forget to update the hash the file will not open and a message box will alert the user
; Convert this file into an EXE and store in an area that only trusted users have write-access to!

; Define variables for excel file location, excel file hash
fileLocation = \path\to\file
fileHash = xxxxxxxxxxxxxxxxxxxxxxx

; if you do not want to add a certificate to the user's trusted publishers:
;   remove the certLocation and certHash lines 
;   and then remove the following from psScript: -And (Get-FileHash -Algorithm SHA256 -Path \"%certLocation%\").Hash -eq '%certHash%')

certLocation = \path\to\certfile.cer
certHash = xxxxxxxxxxxxxxxxxxxxxxx

; Display a splash image with a name and a transparent background
SplashImage, \path\to\splashimage,b,,,nameofimage
Winset, TransColor, Black, nameofimage

; Define a PowerShell script as a variable
; Compare the hashes of the file and the certificate with the expected values
; If the hashes match, import the certificate to the trusted publisher store and run the Excel macro from the file
; If the hashes do not match, display a message box that warns the user about the integrity of the file
psScript =
(
	if ((Get-FileHash -Algorithm SHA256 -Path \"%fileLocation%\").Hash -eq '%fileHash%' -And (Get-FileHash -Algorithm SHA256 -Path \"%certLocation%\").Hash -eq '%certHash%') {
		start-job {  
			Import-Certificate -FilePath \"%certLocation%\" -CertStoreLocation \"Cert:\CurrentUser\TrustedPublisher\" 
			$sc = New-Object -ComObject MSScriptControl.ScriptControl.1
			$sc.Language = 'VBScript'
			$sc.AddCode('
				Set objXL = CreateObject(\"Excel.Application\")
				objXL.DisplayAlerts = False
				objXL.Visible = False
				Set objWkbk = objXL.Workbooks.add(\"%fileLocation%\")
				objXL.Visible = True
			')
		} -runas32 | wait-job | receive-job
	}else{
		start-job {  
			$sc = New-Object -ComObject MSScriptControl.ScriptControl.1
			$sc.Language = 'VBScript'
			$sc.AddCode('
				msgbox(\"The integrity of this file is suspect. Aborting...\")
			')
		} -runas32 | wait-job | receive-job
	}
)

; Run the PowerShell script in a hidden window and hide any output
Run, powershell.exe -windowstyle hidden -Command "& {%psScript%}",,Hide

; Wait for 9 seconds before closing the splash image
Sleep, 9000
SplashImage, Off
