;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;
;
;                                     db   db  .d8b.  .d8888. db   db    .o88b. db   db d88888b  .o88b. db   dD 
;                                     88   88 d8' `8b 88'  YP 88   88   d8P  Y8 88   88 88'     d8P  Y8 88 ,8P' 
;                                     88ooo88 88ooo88 `8bo.   88ooo88   8P      88ooo88 88ooooo 8P      88,8P   
;                                     88~~~88 88~~~88   `Y8b. 88~~~88   8b      88~~~88 88~~~~~ 8b      88`8b   
;                                     88   88 88   88 db   8D 88   88   Y8b  d8 88   88 88.     Y8b  d8 88 `88. 
;                                     YP   YP YP   YP `8888Y' YP   YP    `Y88P' YP   YP Y88888P  `Y88P' YP   YD 
;                                                                        
;                                                                        
;                                              .d8b.  d8b   db d8888b.     d8888b. db    db d8b   db 
;                                             d8' `8b 888o  88 88  `8D     88  `8D 88    88 888o  88 
;                                             88ooo88 88V8o 88 88   88     88oobY' 88    88 88V8o 88 
;                                             88~~~88 88 V8o88 88   88     88`8b   88    88 88 V8o88 
;                                             88   88 88  V888 88  .8D     88 `88. 88b  d88 88  V888 
;                                             YP   YP VP   V8P Y8888D'     88   YD ~Y8888P' VP   V8P 
;
;
;
;                                 check and verify the SHA256 of a excel file before running the workbook_open vba
;
;
;
;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;     This is the location of your file. (Quotes are not needed here)
;    	Any time you update the file you must update the associated SHA256 file hash and recompile.
;           If you forget to update the hash the file will not open and a message box will alert
;           the user that The integrity of this program is suspect and the program will abort.
;    	        To get the SHA256 hash of your file open powershell and use the following command:
;        	echo $(Get-FileHash -Algorithm SHA256 -Path \"%fileLocation%\").Hash


fileLocation = \path\to\file
fileHash = xxxxxxxxxxxxxxxxxxxxxxx


;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;    This is only needed if you want to add a certificate to the user's trusted publishers.
;    	remove this part from psScript if not needed: " -And (Get-FileHash -Algorithm SHA256 -Path \"%certLocation%\").Hash -eq '%certHash%')


certLocation = \path\to\certfile.cer
certHash = xxxxxxxxxxxxxxxxxxxxxxx


;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;    starts a splash image before starting powershell script - comment out if not needed


SplashImage, \path\to\splashimage,b,,,nameofimage
Winset, TransColor, Black, nameofimage


;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;    the script that will run in powershell - this will open the workbook in a new instance using the workbook as a template.
;        change 'objXL.Workbooks.add' to 'objXL.Workbooks.Open' if you don't want to use the file as a template.


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


;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;    uncomment for debugging purposes


;clipboard := psScript
;Run, powershell.exe -NoExit -Command "& {%psScript%}"


;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;    runs the powershell script - comment out if debugging


Run, powershell.exe -windowstyle hidden -Command "& {%psScript%}",,Hide


;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
;
;    removes splash image after 9 seconds - modify to change time or comment out if no splash image


Sleep, 9000
SplashImage, Off


;--------------------------------------------------------------------------------------------------------------------------------------------------------------------
