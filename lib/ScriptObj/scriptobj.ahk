/**
 * =============================================================================================== *
 * @Author           : RaptorX   <graptorx@gmail.com>
 * @Script Name      : Script Object
 * @Script Version   : 0.1.0
 * @Homepage         :
 *
 * @Creation Date    : November 09, 2020
 *
 * @Description      :
 * -------------------
 * This is an object used to have a few common functions between scripts
 * Those are functions related to script information, upgrade and configuration.
 *
 * =============================================================================================== *
 */

; SuperGlobal variables
global null:="",sec:=1000,min:=60*sec,hour:=60*min

class script
{
	name        := ""
	version     := ""
	author      := ""
	email       := ""
	homepage    := ""
	sdbg		:= false
	dbgFile     := ""

	/**
	 * Function: Update
	 * Checks for the current script version
	 * Downloads the remote version information
	 * Compares and automatically downloads the new script file and reloads the script.
	 *
	 * Parameters:
	 * vfile	-	Version File
	 *				Remote version file to be validated against.
	 * rfile	-	Remote File
	 *				Remote script file to be downloaded and installed if a new version is found.
	 *				It should be a zip file that will be unzipped by the function
	 *
	 * Notes:
	 * The versioning file should only contain a version string and nothing else.
	 * The matching will be performed against a SemVer format and only the three
	 * major components will be taken into account.
	 *
	 * e.g. '1.0.0'
	 *
	 * For more information about SemVer and its specs click here: <https://semver.org/>
	 */
	update(vfile, rfile){
		; Error Codes
		static codes := "ERR_INVALIDVFILE|ERR_INVALIDRFILE|ERR_NOCONNECT|ERR_NORESPONSE|"
					 .  "ERR_CURRENTVER|ERR_MSGTIMEOUT|ERR_USRCANCEL"

		loop parse, codes, |
			%a_loopfield% := a_index

		; A URL is expected in this parameter, we just perform a basic check
		; TODO make a more robust match
		if (!regexmatch(vfile, "^((http(s)?|ftp):\/\/)?(([a-z0-9_\-]+\.)*)"))
			return ERR_INVALIDVFILE

		; This function expects a ZIP file
		if (!regexmatch(rfile, "\.zip"))
			return ERR_INVALIDRFILE

		Progress, 50, 50/100, % "Checking for updates", % "Updating"
		; Check if we are connected to the internet
		runwait %a_comspec% /c "Ping -n 2 google.com" ,, Hide
		if (errorlevel)
		 	return ERR_NOCONNECT

		; Download remote version file
		http := comobjcreate("WinHttp.WinHttpRequest.5.1")
		http.Open("GET", vfile, true)
		http.Send(), http.WaitForResponse()

		if !(http.responseText)
			return ERR_NORESPONSE

		; Make sure SemVer is used
		regexmatch(this.version, "\d+\.\d+\.\d+", loVersion)
		regexmatch(http.responseText, "\d+\.\d+\.\d+", remVersion)

		Progress, 100, 100/100, % "Checking for updates", % "Updating"
		Progress, OFF

		; Compare against current stated version
		if (loVersion >= remVersion)
			return ERR_CURRENTVER
		else
		{
			; If new version ask user what to do
			; Yes/No | Icon Question | System Modal
			msgbox % 0x4 + 0x20 + 0x1000
				   , % "New Update Available"
				   , % "There is a new update available for this application.`n"
					 . "Do you wish to upgrade to v" remVersion "?"
				   , 10	; timeout

			ifmsgbox timeout
				return ERR_MSGTIMEOUT
			ifmsgbox no
				return ERR_USRCANCEL

			; Create temporal dirs
			filecreatedir % tmpDir := a_temp "\" regexreplace(a_scriptname, "\..*$")
			filecreatedir % zipDir := tmpDir "\uzip"

			; Create lock file
			fileappend % a_now, % lockFile := tmpDir "\lock"

			; Download zip file
			urldownloadtofile % rfile, % tmpDir "\temp.zip"

			; Extract zip file to temporal folder
			oShell := ComObjCreate("Shell.Application")
			oDir := oShell.NameSpace(zipDir), oZip := oShell.NameSpace(tmpDir "\temp.zip")
			oDir.CopyHere(oZip.Items), oShell := oDir := oZip := ""

			filedelete % tmpDir "\temp.zip"

			/*
			******************************************************
			* Wait for lock file to be released
			* Copy all files to current script directory
			* Cleanup temporal files
			* Run main script
			* EOF
			*******************************************************
			*/
			if (a_iscompiled){
				tmpBatch =
				(Ltrim
					:lock
					if not exist "%lockFile%" goto continue
					timeout /t 10
					goto lock
					:continue

					xcopy "%zipDir%\*.*" "%a_scriptdir%\" /E /C /I /Q /R /K /Y
					if exist "%a_scriptfullpath%" cmd /C "%a_scriptfullpath%"

					cmd /C "rmdir "%tmpDir%" /S /Q"
					exit
				)
				fileappend % tmpBatch, % tmpDir "\update.bat"
				run % a_comspec " /c """ tmpDir "\update.bat""",, hide
			}
			else
			{
				tmpScript =
				(Ltrim
					while (fileExist("%lockFile%"))
						sleep 10

					filecopy %zipDir%\*, %a_scriptdir%, true
					fileremovedir %tmpDir%, true

					if (fileExist("%a_scriptfullpath%"))
						run %a_scriptfullpath%
					else
						msgbox `% 0x10 + 0x1000
							, "Update Error"
							, "There was an error while running the updated version.``n"
							. "Try to run the program manually."
							, 10
						exitapp
				)
				fileappend % tmpScript, % tmpDir "\update.ahk"
				run % a_ahkpath " " tmpDir "\update.ahk"
			}
			filedelete % lockFile
			exitapp
		}
	}

	/**
	* Function: Autostart
	* This Adds the current script to the autorun section for the current
	* user.
	*
	* Parameters:
	* status 	-	Autostart status
	*				It can be either true or false.
	*				Setting it to true would add the registry value.
	*				Setting it to false would delete an existing registry value.
	*/
	autostart(status){
		if (status)
			RegWrite, REG_SZ, HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, %A_ScriptName%
																				, %A_ScriptFullPath%
		else
			RegDelete, HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, %A_ScriptName%
	}

	/**
	* Function: Splash
	* Shows a custom image as a splash screen with a simple fading animation
	*
	* Parameters:
	* img 		-	Image file to be displayed
	* speed 	-	How fast the fading animation will be. Higher value is faster.
	* pause 	-	How long in seconds the image will be paused after fully displayed.
	*/
	splash(img="", speed=10, pause=2){
		global

		Gui, Splash: -Caption +LastFound +Border +AlwaysOnTop +Owner
		$hwnd := WinExist(), alpha := 0
		WinSet, Transparent, 0

		Gui, Splash: add, Picture, x0 y0 vpicImage, % img
		GuiControlGet, picImage, Splash:pos
		Gui, Splash: show, w%picImageW% h%picImageH%

		SetBatchLines 3
		Loop, 255
		{
			if (alpha >= 255)
				break
			alpha += speed
			WinSet, Transparent, %alpha%
		}

		; pause duration in seconds
		sleep pause * 1000

		Loop, 255
		{
			if (alpha <= 0)
				break
			alpha -= speed
			WinSet, Transparent, %alpha%
		}
		SetBatchLines -1

		Gui, Splash: destroy
		return
	}

	/**
	* Funtion: Debug
	* Sends formatted debug messages either to a file or debugger app
	*
	* Parameters:
	* msg 			-	Message to be sent
	* delimiter 	-	Which message delimiter to use. There are 3 tipes:
	* 					specifying 1 would have a message *preceded* by a delimiter
	* 					specifying 2 would have a message *followed* by a delimiter
	* 					specifying 3 would have a message *enclosed* by delimiters
	*/
	debug(msg,delimiter = false){

		static ft := true   ; First time

		t := delimiter = 1 ? msg := "* ------------------------------------------`n" msg
		t := delimiter = 2 ? msg := msg "`n* ------------------------------------------"
		t := delimiter = 3 ? msg := "* ------------------------------------------`n" msg
								 .  "`n* ------------------------------------------"

		script.sdbg && ft ? (msg := "* ------------------------------------------`n"
								 .  "* " script.name " Debug ON`n* " script.name "[Start]`n" msg
								 , ft := 0)

		if (script.sdbg && !script.dbgFile)
			OutputDebug, %msg%
		else if (script.sdbg)
			FileAppend, %msg%`n, % script.dbgFile
	}
}
