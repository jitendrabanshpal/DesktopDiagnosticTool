#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=DesktopDiagnosticTool_x86.exe
#AutoIt3Wrapper_Outfile_x64=DesktopDiagnosticTool_x64.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=Desktop Diagnostic Tool
#AutoIt3Wrapper_Res_Description=Desktop Diagnostic Tool
#AutoIt3Wrapper_Res_Fileversion=1.0.0.2
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;#AutoIt3Wrapper_Icon=tool-icon-27.ico
;~ ###	Author:  		Jitendra Banshpal
;~ ###	FileName:  		DesktopDiagnosisTool.au3
;~ ###	Purpose: 		Desktop Diagnostic Tool provides a clear picture of the currency (or presence) of all relevant desktop components,
;~ ### 					via a text and html output file, which can be sent back to the Service teams for analysis.
;~ ### Date Created:  	04/03/2019

#include <File.au3>
#include <Array.au3>
#include <Date.au3>
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <IE.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>

$filedate = StringReplace(_NowCalcDate(), "/", "-")
$filetime = StringReplace(_NowTime(), ":", "-")
Global $FileName = "Desktop_Diagnostic_Log_" & $filedate & "_" & $filetime & ".txt"
Global $FileHandler = @DesktopDir & "\Desktop_Diagnostic_Log_" & $filedate & "_" & $filetime & ".txt"
FileOpen($FileHandler, 2)
FileWrite($FileHandler, "")
FileClose($FileHandler)
FileOpen($FileHandler, 1)
FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Log Started <<<<<<<<<<<<<<<<<<<" & @CRLF)
FileWrite($FileHandler, "Report Generation Date and Time is: " & _NowCalcDate() & " " & _NowTime() & @CRLF)
FileWrite($FileHandler, "==========================================================================================" & @CRLF)

;JAVA_SECURITY_DETAILS function helps to find java security level
Func JAVA_SECURITY_DETAILS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Java Security Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	$JavaDeploymentPath = @LocalAppDataDir & "Low\Sun\Java\Deployment\deployment.properties"
	Global $hFileOpen = FileOpen($JavaDeploymentPath, $FO_READ)
	If $hFileOpen = -1 Then
		FileWrite($FileHandler, "deployment.properties doesn't exist " & @CRLF)
	Else
		Global $sFileRead = FileReadLine($hFileOpen, 3)
		$findStringlenghtUptoEqual = StringInStr($sFileRead, "=")
		$fullStringLength = StringLen($sFileRead)
		$leftString = StringLeft($sFileRead, $findStringlenghtUptoEqual)
		If $leftString <> "deployment.security.level=" Then
			FileWrite($FileHandler, "Java Security Level cannot be identified. Please check manually. Go to Control Panel >> Java >> Security." & @CRLF)
		Else
			$exactStringLenght = $fullStringLength - $findStringlenghtUptoEqual
			$securityLevel = StringRight($sFileRead, $exactStringLenght)
			If $securityLevel == "later" Then
				$securityLevel = "HIGH"
			EndIf
			FileWrite($FileHandler, "Java Security Level = " & $securityLevel & @CRLF)
			FileClose($hFileOpen)
		EndIf
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>JAVA_SECURITY_DETAILS

;JAVA_SECURITY_TRUSTED_SITES_DETAILS function helps to find java security trusted site details
Func JAVA_SECURITY_TRUSTED_SITES_DETAILS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Java Security Trusted Site Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	$JavaExcpetionPath = @LocalAppDataDir & "Low\Sun\Java\Deployment\security\exception.sites"
	Global $hFileOpen = FileOpen($JavaExcpetionPath, $FO_READ)
	If $hFileOpen = -1 Then
		FileWrite($FileHandler, "exception.sites doesn't exist " & @CRLF)
	Else
		Global $sFileRead = FileRead($hFileOpen, $FO_READ)
		Global $aInput
		_FileReadToArray($JavaExcpetionPath, $aInput)
		If UBound($aInput) > 0 Then
			For $i = 1 To UBound($aInput) - 1
				FileWrite($FileHandler, $aInput[$i] & @CRLF)
			Next
			FileClose($hFileOpen)
		Else
			FileWrite($FileHandler, "No Trusted sites present." & @CRLF)
		EndIf
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>JAVA_SECURITY_TRUSTED_SITES_DETAILS

;OPERATING_SYSTEM_WITH_ARCH function helps to find client machine's OS version, OS Architecture, OS Build, OS Service Pack and OS Language.
Func OPERATING_SYSTEM_WITH_ARCH()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Operating System Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	FileWrite($FileHandler, "OS Version is: " & @OSVersion & @CRLF)
	FileWrite($FileHandler, "OS Architect is: " & @OSArch & @CRLF)
	FileWrite($FileHandler, "OS Build is: " & @OSBuild & @CRLF)
	If @OSLang == '0809' Then
		FileWrite($FileHandler, "OS Language is: " & "English - United Kingdom" & @CRLF)
	ElseIf @OSLang == '0409' Then
		FileWrite($FileHandler, "OS Language is: " & "English - United States" & @CRLF)
	Else
		FileWrite($FileHandler, "OS Language is: " & @OSLang & @CRLF)
	EndIf
	If @OSServicePack = "" Then
		FileWrite($FileHandler, "OS Service Pack is not present" & @CRLF)
	Else
		FileWrite($FileHandler, "OS Service Pack is: " & @OSServicePack & @CRLF)
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>OPERATING_SYSTEM_WITH_ARCH

;REG_READ_SOFTWARE_DETAILS function is generic function and gets parameter as Registry Path, Display Name and Alternate Display name (Optional) and trace through the registry and provide the details.

Func REG_READ_SOFTWARE_DETAILS($regapth, $displayname, $altdisplayname)
	$SoftVersionNo = Null
	Local $returnvalues[3]
	If $altdisplayname == Null Then
		For $j = 1 To 500
			$AppKey = RegEnumKey($regapth, $j)
			If @error <> 0 Then ExitLoop
			If StringInStr(RegRead($regapth & "\" & $AppKey, "DisplayName"), $displayname) Then
				$SoftVersionNo = RegRead($regapth & "\" & $AppKey, "DisplayVersion")
				$returnvalues[0] = $SoftVersionNo
				$SoftVersionDetails = RegRead($regapth & "\" & $AppKey, "DisplayName")
				$returnvalues[1] = $SoftVersionDetails
				$InstallationDate = RegRead($regapth & "\" & $AppKey, "InstallDate")
				$returnvalues[2] = $InstallationDate
				Return ($returnvalues)
			EndIf
		Next
		If $SoftVersionNo == Null Then
			$returnvalues[0] = Null
			Return ($returnvalues)
		EndIf
	Else
		For $j = 1 To 500
			$AppKey = RegEnumKey($regapth, $j)
			If @error <> 0 Then ExitLoop
			If StringInStr(RegRead($regapth & "\" & $AppKey, "DisplayName"), $displayname) Or StringInStr(RegRead($regapth & "\" & $AppKey, "DisplayName"), $altdisplayname) Then
				$SoftVersionNo = RegRead($regapth & "\" & $AppKey, "DisplayVersion")
				$returnvalues[0] = $SoftVersionNo
				$SoftVersionDetails = RegRead($regapth & "\" & $AppKey, "DisplayName")
				$returnvalues[1] = $SoftVersionDetails
				$InstallationDate = RegRead($regapth & "\" & $AppKey, "InstallDate")
				$returnvalues[2] = $InstallationDate
				Return ($returnvalues)
			EndIf
		Next
		If $SoftVersionNo == Null Then
			$returnvalues[0] = Null
			Return ($returnvalues)
		EndIf
	EndIf
EndFunc   ;==>REG_READ_SOFTWARE_DETAILS


;JAVA32BIT_AND_JRE_DETAILS function to get details of Java version based 32 bit installation.

Func JAVA32BIT_AND_JRE_DETAILS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> JRE/JDK Version Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	Global $C_version32
	Global $HP_version32
	Global $javaversion32
	Global $Full_Version32
	Global $B_version32
	RegRead("HKEY_LOCAL_MACHINE\Software\JavaSoft\Java Runtime Environment", "CurrentVersion")
	If @error == 0 Then
		$C_version32 = RegRead("HKEY_LOCAL_MACHINE\Software\JavaSoft\Java Runtime Environment", "CurrentVersion")
		$HP_version32 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\" & $C_version32, "JavaHome")
		$B_version32 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment", "BrowserJavaVersion")
		If $B_version32 == "" Then
			$B_version32 = "No Browser Version Present"
		EndIf
		FileWrite($FileHandler, "HKEY_LOCAL_MACHINE\Software\JavaSoft\Java Runtime Environment" & @CRLF)
		FileWrite($FileHandler, "Current JRE Version " & $C_version32 & @CRLF)
		FileWrite($FileHandler, "Current JRE Home path " & $HP_version32 & @CRLF)
		FileWrite($FileHandler, "Current JRE Browser Version : " & $B_version32 & @CRLF)
	Else
		RegRead("HKEY_LOCAL_MACHINE\Software\JavaSoft\Java Development Kit", "CurrentVersion")
		If @error == 0 Then
			$C_version32 = RegRead("HKEY_LOCAL_MACHINE64\Software\JavaSoft\Java Development Kit", "CurrentVersion")
			$HP_version32 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Development Kit\" & $C_version32, "JavaHome")
			$B_version32 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Development Kit", "BrowserJavaVersion")
			If $B_version32 == "" Then
				$B_version32 = "No Browser Version Present"
			EndIf
			FileWrite($FileHandler, "HKEY_LOCAL_MACHINE\Software\JavaSoft\Java Development Kit\" & @CRLF)
			FileWrite($FileHandler, "Current JDK Version " & $C_version32 & @CRLF)
			FileWrite($FileHandler, "Current JDK Home path " & $HP_version32 & @CRLF)
			FileWrite($FileHandler, "Current JDK Browser Version : " & $B_version32 & @CRLF)
		Else
			$C_version32 = "32 bit JRE/JDK is not Installed."
			FileWrite($FileHandler, "JRE/JDK is not Installed on following registry path." & @CRLF)
			FileWrite($FileHandler, "HKEY_LOCAL_MACHINE\Software\JavaSoft\Java Runtime Environment" & " or " & "HKEY_LOCAL_MACHINE\Software\JavaSoft\Java Development Kit\" & @CRLF)
		EndIf
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>JAVA32BIT_AND_JRE_DETAILS


;JAVA64BIT_AND_JRE_DETAILS function to get details of Java version based 64 bit installation.
Func JAVA64BIT_AND_JRE_DETAILS()
	Global $C_version64
	Global $HP_version64
	Global $javaversion64
	Global $Full_Version64
	Global $B_version64
	RegRead("HKEY_LOCAL_MACHINE\Software\Wow6432Node\JavaSoft\Java Runtime Environment", "CurrentVersion")
	If @error == 0 Then
		$C_version64 = RegRead("HKEY_LOCAL_MACHINE64\Software\Wow6432Node\JavaSoft\Java Runtime Environment", "CurrentVersion")
		$HP_version64 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\JavaSoft\Java Runtime Environment\" & $C_version64, "JavaHome")
		$B_version64 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\JavaSoft\Java Runtime Environment", "BrowserJavaVersion")
		If $B_version64 == "" Then
			$B_version64 = "No Browser Version Present"
		EndIf
		FileWrite($FileHandler, "HKEY_LOCAL_MACHINE\Software\Wow6432Node\JavaSoft\Java Runtime Environment" & @CRLF)
		FileWrite($FileHandler, "Current JRE Version " & $C_version64 & @CRLF)
		FileWrite($FileHandler, "Current JRE Home path " & $HP_version64 & @CRLF)
		FileWrite($FileHandler, "Current JRE Browser Version : " & $B_version64 & @CRLF)
	Else
		RegRead("HKEY_LOCAL_MACHINE\Software\Wow6432Node\JavaSoft\Java Development Kit", "CurrentVersion")
		If @error == 0 Then
			$C_version64 = RegRead("HKEY_LOCAL_MACHINE64\Software\Wow6432Node\JavaSoft\Java Development Kit", "CurrentVersion")
			$HP_version64 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\JavaSoft\Java Development Kit\" & $C_version64, "JavaHome")
			$B_version64 = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\JavaSoft\Java Development Kit", "BrowserJavaVersion")
			If $B_version64 == "" Then
				$B_version64 = "No Browser Version Present"
			EndIf
			FileWrite($FileHandler, "HKEY_LOCAL_MACHINE\Software\Wow6432Node\JavaSoft\Java Development Kit\" & @CRLF)
			FileWrite($FileHandler, "Current JDK Version " & $C_version64 & @CRLF)
			FileWrite($FileHandler, "Current JDK Home path " & $HP_version64 & @CRLF)
			FileWrite($FileHandler, "Current JDK Browser Version : " & $B_version64 & @CRLF)
		Else
			$C_version64 = "64 bit JRE/JDK is not Installed."
			FileWrite($FileHandler, "JRE/JDK is not Installed on following registry path." & @CRLF)
			FileWrite($FileHandler, "HKEY_LOCAL_MACHINE\Software\Wow6432Node\JavaSoft\Java Runtime Environment" & " or " & "HKEY_LOCAL_MACHINE\Software\Wow6432Node\JavaSoft\Java Development Kit\" & @CRLF)
		EndIf
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>JAVA64BIT_AND_JRE_DETAILS

;BROWSER_CHROME_DETAILS function to get details of chrome browser installation and version.

Func BROWSER_CHROME_DETAILS()
	Global $Chrome
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Chrome Browser Version Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	$FilePath = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", "")
	If FileExists($FilePath) Then
		$Chrome = FileGetVersion($FilePath)
		If $Chrome <> "0.0.0.0" Then
			FileWrite($FileHandler, "Chrome version is: " & $Chrome & @CRLF)
		Else
			FileWrite($FileHandler, "Chrome is not installed." & @CRLF)
			$Chrome = "Chrome is not installed."
		EndIf
	ElseIf FileExists(@LocalAppDataDir & "\Google\Chrome\Application\chrome.exe") Then
		$Chrome = FileGetVersion(@LocalAppDataDir & "\Google\Chrome\Application\chrome.exe")
		If $Chrome <> "0.0.0.0" Then
			FileWrite($FileHandler, "Chrome version is: " & $Chrome & @CRLF)
		Else
			FileWrite($FileHandler, "Chrome is not installed." & @CRLF)
			$Chrome = "Chrome is not installed."
		EndIf
	ElseIf FileExists(@ProgramFilesDir & "\Google\Chrome\Application\chrome.exe") Then
		$Chrome = FileGetVersion(@LocalAppDataDir & "\Google\Chrome\Application\chrome.exe")
		If $Chrome <> "0.0.0.0" Then
			FileWrite($FileHandler, "Chrome version is: " & $Chrome & @CRLF)
		Else
			FileWrite($FileHandler, "Chrome is not installed." & @CRLF)
			$Chrome = "Chrome is not installed."
		EndIf
	Else
		FileWrite($FileHandler, "Chrome Browser is not installed." & @CRLF)
		$Chrome = "Chrome is not installed."
	EndIf

	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>BROWSER_CHROME_DETAILS

;BROWSER_FIREFOX_DETAILS function to get details of Firefox browser installation and version.

Func BROWSER_FIREFOX_DETAILS()
	Global $FireFox
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> FireFox Browser Version Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	$FilePath = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe", "")
	If FileExists($FilePath) Then
		$FireFox = FileGetVersion($FilePath)
		If $FireFox <> "0.0.0.0" Then
			FileWrite($FileHandler, "FireFox version is: " & $FireFox & @CRLF)
		Else
			FileWrite($FileHandler, "FireFox is not installed." & @CRLF)
			$FireFox = "FireFox is not installed."
		EndIf
	ElseIf FileExists(@LocalAppDataDir & "\Mozilla Firefox\firefox.exe") Then
		$FireFox = FileGetVersion(@LocalAppDataDir & "\Mozilla Firefox\firefox.exe")
		If $FireFox <> "0.0.0.0" Then
			FileWrite($FileHandler, "FireFox version is: " & $FireFox & @CRLF)
		Else
			FileWrite($FileHandler, "FireFox is not installed." & @CRLF)
			$FireFox = "FireFox is not installed."
		EndIf
	Else
		FileWrite($FileHandler, "FireFox Browser is not installed." & @CRLF)
		$FireFox = "FireFox is not installed."
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>BROWSER_FIREFOX_DETAILS

;BROWSER_INTERNET_EXPLORER_DETAILS function to get details of Internet browser installation and version.

Func BROWSER_INTERNET_EXPLORER_DETAILS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Internet Explorer Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	Global $svcVersion = (RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer", "svcVersion"))
	Local $svcKBNumber = (RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer", "svcKBNumber"))
	Local $svcUpdateVersion = (RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer", "svcUpdateVersion"))
	If $svcVersion <> "" Then
		FileWrite($FileHandler, "Internet Explorer version is: " & $svcVersion & @CRLF)
	Else
		Local $Version = (RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer", "Version"))
		If $Version <> "" Then
			FileWrite($FileHandler, "Internet Explorer version is: " & $Version & @CRLF)
		Else
			FileWrite($FileHandler, "Internet Explorer is not present." & @CRLF)
		EndIf
	EndIf
	If $svcKBNumber <> "" Then
		FileWrite($FileHandler, "Internet Explorer KB version is: " & $svcKBNumber & @CRLF)
	Else
		FileWrite($FileHandler, "Internet Explorer KB version is not present." & @CRLF)
	EndIf
	If $svcUpdateVersion <> "" Then
		FileWrite($FileHandler, "Internet Explorer updated version is: " & $svcUpdateVersion & @CRLF)
	Else
		FileWrite($FileHandler, "Internet Explorer updated version is not present." & @CRLF)
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>BROWSER_INTERNET_EXPLORER_DETAILS

;LOCAL_LIST_OF_ALL_PRINTERS function to get details of installed printers.

Func LOCAL_LIST_OF_ALL_PRINTERS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> List of Installed Printers <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	Local $arrPrinterList[10000]
	If StringInStr(@OSVersion, "WIN_20") Then
		FileWrite($FileHandler, "No Printer List for servers" & @CRLF)
	Else
		For $i = 1 To 10000
			$reg = RegEnumVal("HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Devices", $i)
			If @error = -1 Then ExitLoop
			$arrPrinterList[$i] = $reg
			FileWrite($FileHandler, "Printer No : " & $i & " = " & $arrPrinterList[$i] & @CRLF)
		Next
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>LOCAL_LIST_OF_ALL_PRINTERS

;USERS_MACHINE_ENVIRONMENT_PATH_DETAILS function to get details of Environment path from client machine.

Func USERS_MACHINE_ENVIRONMENT_PATH_DETAILS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Environment Path Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	Local $sEnvVar = EnvGet("PATH")
	FileWrite($FileHandler, "The environment variable %PATH% has the value of: " & @CRLF & $sEnvVar & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>USERS_MACHINE_ENVIRONMENT_PATH_DETAILS


;DOT_NET_VERSION_DETAILS function to get details of .net Frameworks from client machine.
Func DOT_NET_VERSION_DETAILS($bOnlyInstalled = False)
	Local $i = 1, $iClientInstall, $iFullInstall, $iNum, $aVersions[100][4], $iTotal = 0, $bVer4Found = 0
	Local $sKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP", $sSubKey
	; Detect v1.0 (special key)
	RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\.NETFramework\Policy\v1.0", "3705")
	; If value was read (key exists), and is of type REG_SZ (1 as defined in <Constants.au3>), v1.0 is installed
	If @error = 0 And @extended = 1 Then
		$iTotal += 1
		$aVersions[$iTotal][0] = 1.0
		$aVersions[$iTotal][1] = 'v1.0.3705'
		$aVersions[$iTotal][2] = 1
		$aVersions[$iTotal][3] = 1
	EndIf
	While 1
		$iClientInstall = 0
		$iFullInstall = 0
		$sSubKey = RegEnumKey($sKey, $i)
		If @error Then ExitLoop
		$i += 1
		; 'v4.0' is a deprecated version.  Since it comes after 'v4' (the newer version) while enumerating,
		;   a simple check if 'v4' has already been found is sufficient
		If $sSubKey = 'v4.0' And $bVer4Found Then ContinueLoop
		$iNum = Number(StringMid($sSubKey, 2)) ; cuts off at any 2nd decimal points (obviously)
		; Note - one of the SubKeys is 'CDF'. Number() will return 0 in this case [we can safely ignore that]
		If $iNum = 0 Then ContinueLoop
		If $iNum < 4 Then
			$iClientInstall = RegRead($sKey & '\' & $sSubKey, 'Version')
			If $iClientInstall Then $iFullInstall = 1 ; older versions were all-or-nothing I believe
		Else
			; Version 4 works with one or both of these keys.  One can only hope new versions keep the same organization
			$iFullInstall = RegRead($sKey & '\' & $sSubKey & '\Full', 'Version')
			If @error Then $iFullInstall = 0
			$iClientInstall = RegRead($sKey & '\' & $sSubKey & '\Client', 'Version')
			If $iNum < 5 Then $bVer4Found = True
		EndIf
		If @error Then $iClientInstall = 0
		If $bOnlyInstalled And $iClientInstall = 0 And $iFullInstall = 0 Then ContinueLoop
		$iTotal += 1
		$aVersions[$iTotal][0] = $iNum
		$aVersions[$iTotal][1] = $sSubKey
		$aVersions[$iTotal][2] = $iClientInstall
		$aVersions[$iTotal][3] = $iFullInstall
	WEnd
	If $iTotal = 0 Then Return SetError(-3, @error, '')
	$aVersions[0][0] = $iTotal
	ReDim $aVersions[$iTotal + 1][4]
	Return $aVersions
EndFunc   ;==>DOT_NET_VERSION_DETAILS

;INTERNET_EXPLORER_TRUSTED_SITES_DETAILS function to get list of trusted sites from Internet browser where user need to replace nhs.uk.

Func INTERNET_EXPLORER_TRUSTED_SITES_DETAILS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Internet Explorer Trusted Sites Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	For $j = 1 To 50
		$internetExplorerTrustedSitesRegPath = RegEnumKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nhs.uk", $j)
		If @error <> 0 And $j <= 1 Then
			FileWrite($FileHandler, "Trusted sites are not configured" & @CRLF)
			ExitLoop
		EndIf
		$internetExplorerTrustedSitesRegFullPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\nhs.uk" & "\" & $internetExplorerTrustedSitesRegPath
		$returnValueOfKey = RegRead($internetExplorerTrustedSitesRegFullPath, "https")
		If @error <> 0 And $j <= 1 Then
			FileWrite($FileHandler, "Trusted sites are not configured" & @CRLF)
			ExitLoop
		EndIf
		If $returnValueOfKey == 2 Then
			FileWrite($FileHandler, "Trusted Sites Are : " & $internetExplorerTrustedSitesRegPath & @CRLF)
		EndIf
	Next
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>INTERNET_EXPLORER_TRUSTED_SITES_DETAILS


;REPORT_DOT_NET_VERSION_DETAILS function to report the all details of .net framework versions.

Func REPORT_DOT_NET_VERSION_DETAILS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Multiple Version .NET Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	$adotNetVersions = DOT_NET_VERSION_DETAILS()
	Local $iRows = UBound($adotNetVersions, $UBOUND_ROWS)
	Local $iCols = UBound($adotNetVersions, $UBOUND_COLUMNS)
	For $i = 1 To $iRows - 1
		FileWrite($FileHandler, "===================.Net Version " & $i & " Details ==========================" & @CRLF)
		For $j = 0 To $iCols - 1
			If $j == 0 Then
				FileWrite($FileHandler, ".Net Main Version # " & $adotNetVersions[$i][$j] & @CRLF)
			ElseIf $j == 1 Then
				FileWrite($FileHandler, ".Net Full Version String - " & $adotNetVersions[$i][$j] & @CRLF)
			ElseIf $j == 2 Then
				FileWrite($FileHandler, ".Net Full Version Details - " & $adotNetVersions[$i][$j] & @CRLF)
			EndIf
		Next
		FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	Next
EndFunc   ;==>REPORT_DOT_NET_VERSION_DETAILS

;SMARTCARD_DRIVER_DETAILS function to get details of smartcard drivers installed in client machine.

Func SMARTCARD_DRIVER_DETAILS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Smartcard Reader Driver Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	Local $GEMMWVersionNo[2]
	Local $sKey1
	Local $Driver[200]
	If @OSArch = "X64" Then
		$sKey = "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Class"
	Else
		$sKey = "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Class"
	EndIf
	For $j = 1 To 200
		$AppKey = RegEnumKey($sKey, $j)
		If @error <> 0 Then ExitLoop
		If StringInStr(RegRead($sKey & "\" & $AppKey, "Class"), "SmartCardReader") Then
			$sKey1 = $sKey & "\" & $AppKey
			ExitLoop
		EndIf
	Next
	For $i = 1 To 200
		$AppKey2 = RegEnumKey($sKey1, $i)
		If @error <> 0 Then ExitLoop
		$Driver[$i] = RegRead($sKey1 & "\" & $AppKey2, "DriverDesc")
		If @error <> 0 Then ExitLoop
	Next
	$Driver = _ArrayUnique($Driver)
	For $l = UBound($Driver) - 1 To 0 Step -1
		If $Driver[$l] = "" Then
			_ArrayDelete($Driver, $l)
		EndIf
	Next
	For $k = 1 To UBound($Driver) - 1
		FileWrite($FileHandler, "List of Smartcard Drivers:" & "[" & $k & "] : " & $Driver[$k] & @CRLF)
	Next
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>SMARTCARD_DRIVER_DETAILS

;CARD_READER_DETAILS function to get details of smartcard readers in client machine.

Func CARD_READER_DETAILS($sCommand)
	Local $nResult = Run('"' & @ComSpec & '" /c ' & $sCommand, @SystemDir, @SW_HIDE, 6)
	ProcessWaitClose($nResult)
	Return StdoutRead($nResult)
EndFunc   ;==>CARD_READER_DETAILS

;SMARTCARD_READER_LIST function to get details of smartcard readers in client machine and report in log.

Func SMARTCARD_READER_LIST()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> Smartcard Reader Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	$listofReaders = CARD_READER_DETAILS('certutil -scinfo')
	$substring1 = StringInStr($listofReaders, "---")
	$strleft = StringLeft($listofReaders, $substring1 - 1)
	$substring2 = StringInStr($strleft, "us:")
	$strright = StringTrimLeft($strleft, $substring2)
	$strright = StringTrimLeft($strright, 3)
	$strright = StringTrimRight($strright, 2)
	If $strright <> "" Then
		FileWrite($FileHandler, "Total Smartcard " & $strright & @CRLF)
	Else
		FileWrite($FileHandler, "No Smartcard Reader Connected." & @CRLF)
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)

EndFunc   ;==>SMARTCARD_READER_LIST

;VM_HORIZON_CLIENT_DETAILS function to get details about VM Horizon Client installed in client machine.

Func VM_HORIZON_CLIENT_DETAILS()
	FileWrite($FileHandler, ">>>>>>>>>>>>>>>>>>> VMware Horizon Client Details <<<<<<<<<<<<<<<<<<<" & @CRLF)
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
	Local $VmHorizonClientVersionNo[2]
	If @OSArch == "X86" Then
		$sKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
	Else
		$sKey = "HKEY_LOCAL_MACHINE64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
	EndIf
	$VmHorizonClientVersionNo = REG_READ_SOFTWARE_DETAILS($sKey, "VMware Horizon", Null)
	If $VmHorizonClientVersionNo[0] == Null Then
		FileWrite($FileHandler, "VMware Horizon Client is not installed." & @CRLF)
	Else
		FileWrite($FileHandler, "VMware Horizon Client version is : " & $VmHorizonClientVersionNo[0] & @CRLF)
		FileWrite($FileHandler, "VMware Horizon Client details : " & $VmHorizonClientVersionNo[1] & @CRLF)
	EndIf
	FileWrite($FileHandler, "==========================================================================================" & @CRLF)
EndFunc   ;==>VM_HORIZON_CLIENT_DETAILS



;======================Calling all Fucntion based on categories=============================================

; Operating System Details
OPERATING_SYSTEM_WITH_ARCH()

;Printer Details
LOCAL_LIST_OF_ALL_PRINTERS()

;Java & JRE details
JAVA32BIT_AND_JRE_DETAILS()
JAVA64BIT_AND_JRE_DETAILS()

;Browser Versions details
BROWSER_CHROME_DETAILS()
BROWSER_FIREFOX_DETAILS()
BROWSER_INTERNET_EXPLORER_DETAILS()
INTERNET_EXPLORER_TRUSTED_SITES_DETAILS()

;.net version details
REPORT_DOT_NET_VERSION_DETAILS()

;Java details

JAVA_SECURITY_DETAILS()
JAVA_SECURITY_TRUSTED_SITES_DETAILS()

;Smartcard Details
SMARTCARD_DRIVER_DETAILS()
SMARTCARD_READER_LIST()

;VM Horinzon Client details
VM_HORIZON_CLIENT_DETAILS()
;Broken Spring Check for HSCIC Identity Agents

;Other Information (environment Path)
USERS_MACHINE_ENVIRONMENT_PATH_DETAILS()

FileWrite($FileHandler, "==========================================================================================" & @CRLF)
FileWrite($FileHandler, "Report Generation Completed Date and Time is: " & _NowCalcDate() & " " & _NowTime() & @CRLF)
FileWrite($FileHandler, "==========================================================================================")



;==============================================================================================================================
; Write report Details to HTML File
;==============================================================================================================================
Global $sHTML = ""
Global $arr1 = []
Global $FileHand
Global $FileHandler1

$FileHandler1 = StringLeft($FileHandler, StringLen($FileHandler) - 4) & ".html"

WriteHtml($FileHandler, $FileHandler1)

Func WriteHtml(ByRef $file, ByRef $fileWrite)

	FileOpen($fileWrite, 2)
	$sHTML &= "<HTML>" & @CRLF
	$sHTML &= "<HEAD>" & @CRLF
	$sHTML &= "<TITLE>Diagnostic Log Report</TITLE>" & @CRLF
	$sHTML &= "<H1 align=center >" & "Detailed Summary Report" & "</H1>" & @CRLF
	$sHTML &= "</HEAD>" & @CRLF
	$sHTML &= "<BR>" & @CRLF

	Global $iConst1 = 0
	Global $iTable = 0
	Global $iRows = 0
	Global $iRowE = 0
	Global $sHeader = ""
	Global $sPos = 21
	Global $iCounter = 0

	FileOpen($file, 0)
	For $i = 1 To 2
		$line = FileReadLine($file, $i)
		If (StringInStr($line, ">>>>>>>>>>>>>>>>>>>") > 0) Then
			Local $ePos = StringInStr($line, "<")
			$sHeader = StringMid($line, $sPos, $ePos - $sPos - 1)
		Else
			$sContent = $line
		EndIf
	Next
	$sHTML &= "<TABLE style=width:100% , border=0>" & @CRLF
	$sHTML &= "<TR bgcolor=#F0EEEE>" & @CRLF
	$sHTML &= "<TH>" & $sHeader & "</TH>" & @CRLF
	$sHTML &= "</TR>" & @CRLF
	$sHTML &= "</TR>" & @CRLF
	$sHTML &= "<TH bgcolor=#F4F719>" & $sContent & "</TH>" & @CRLF
	$sHTML &= "</TR>" & @CRLF
	$sHTML &= "</TABLE>" & @CRLF
	$sHTML &= "<BR>" & @CRLF

	For $i = 4 To _FileCountLines($file) - 3
		$iCounter = $iCounter + 1
		$line = FileReadLine($file, $i)
		If (StringInStr($line, ">>>>>>>>>>>>>>>>>>>") > 0) Then
			$iConst1 = 1
			$iTable = 1
			$iRows = 1
			$i = $i + 1
			Local $ePos = StringInStr($line, "<")
			$sHeader = StringMid($line, $sPos, $ePos - $sPos - 1)
		ElseIf ((StringInStr($line, "======") > 0) And (StringInStr($line, "Details") > 0)) Then
			$line = StringMid($line, 20, StringLen($line) - 35) & " :-"
			_ArrayAdd($arr1, $line)
		ElseIf ((StringInStr($line, "======") > 0) And (StringInStr($line, "bit") > 0)) Then
			$line = StringMid($line, 20, StringLen($line) - 35)
			_ArrayAdd($arr1, $line)
		ElseIf (StringInStr($line, "======") > 0) Then
			$iRowE = $iCounter
			CreateTable($iRowE, $sHeader, $arr1)
			$iCounter = 0
			For $j = UBound($arr1) To 0 Step -1
				_ArrayDelete($arr1, $j)
			Next
		ElseIf (StringInStr($line, "Report Generation Completed") > 0) Then
			$sHTML &= "<TABLE style=width:100% , border=0>" & @CRLF
			$sHTML &= "<TR>" & @CRLF
			$sHTML &= "<TH bgcolor=#F0EFEE>" & $line & "</TH>" & @CRLF
			$sHTML &= "</TR>" & @CRLF
			$sHTML &= "</TABLE>" & @CRLF
			$sHTML &= "<BR>" & @CRLF

		Else
			_ArrayAdd($arr1, $line)
		EndIf
	Next

	For $l = _FileCountLines($file) - 3 To _FileCountLines($file) - 1

		$line = FileReadLine($file, $l)
		If (StringInStr($line, "Report Generation Completed") > 0) Then
			;  msgbox(0,'', $line)
			$sHTML &= "<TABLE style=width:100% , border=0>" & @CRLF
			$sHTML &= "<TR>" & @CRLF
			$sHTML &= "<TH bgcolor=#F4F719>" & $line & "</TH>" & @CRLF
			$sHTML &= "</TR>" & @CRLF
			$sHTML &= "</TABLE>" & @CRLF
			$sHTML &= "<BR>" & @CRLF

		EndIf

	Next

	$sHTML &= "</HTML>"

	FileWrite($fileWrite, $sHTML)
	FileClose($file)
EndFunc   ;==>WriteHtml

Func CreateTable(ByRef $ierow, ByRef $sheadercontent, ByRef $textArray)
	$sHTML &= "<TABLE style=width:100% , border=0>" & @CRLF
	$sHTML &= "<TR bgcolor=#F0EEEE>" & @CRLF
	$sHTML &= "<TH>" & "<font size=4 face=Trebuchet MS>" & $sheadercontent & "</font>" & "</TH>" & @CRLF
	$sHTML &= "</TR>" & @CRLF
	For $k = 0 To UBound($textArray) - 1
		$sHTML &= "<TR bgcolor=#CAFCF1>" & @CRLF
		$sHTML &= "<TH align=left style=font-weight:normal>" & "<font size=3 face=Trebuchet MS>" & $textArray[$k] & "</font>" & "</TH>" & @CRLF
		$sHTML &= "</TR>" & @CRLF
	Next
	$sHTML &= "</TABLE>" & @CRLF
EndFunc   ;==>CreateTable

;===============================================================================================================================
;END
;===============================================================================================================================


;===============================================================================================================================
;If User don't want UI for Diagnostic Tool they can comment following code section
;===============================================================================================================================



;===============================================================================================================================
;GUI setup and updation of key data on UI
;===============================================================================================================================


#CS Opt("GUIOnEventMode", 1)
	; Local $Mainpage = GUICreate('Client Machine Details', 800, 450, -1, -1, BitXOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX, $WS_MINIMIZEBOX), -1, WinGetHandle(AutoItWinGetTitle()))
	; GUISetFont(10, 300, "Comic Sans MS")
	; GUISetBkColor(0x00E0FFFF)
	; ;GUICtrlCreateLabel("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", -1, 165)
	; GUISetFont(11, 700, "Comic Sans MS")
	; GUICtrlCreateLabel("Operating System Details", 300, 10, 250)
	; GUISetFont(10, 300, "Comic Sans MS")
	; GUICtrlCreateLabel("", 0, 1, 450)
	; GUICtrlCreateLabel("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", -1, 25)
	; GUICtrlCreateLabel("OS Version is: " & @OSVersion, 15, 45, 150)
	; GUICtrlCreateLabel("OS Architect is: " & @OSArch, 525, 45, 150)
	; GUICtrlCreateLabel("", 0, 1, 450)
	; GUICtrlCreateLabel("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", -1, 75)
	; GUISetFont(11, 700, "Comic Sans MS")
	; GUICtrlCreateLabel("Java JRE/JDK Details", 525, 100, 250)
	; GUISetFont(10, 300, "Comic Sans MS")
	; GUICtrlCreateLabel("", 0, 1, 450)
	; GUICtrlCreateLabel("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", -1, 125)
	; GUISetFont(10, 700, "Comic Sans MS")
	; GUICtrlCreateLabel("32 bit Java Details", 15, 150, 150)
	; GUISetFont(10, 300, "Comic Sans MS")
	; If $C_version32 == "32 bit JRE/JDK is not Installed." Then
	; 	GUICtrlCreateLabel($C_version32, 15, 175, 250)
	; Else
	; 	GUICtrlCreateLabel("32 bit JRE/JDK version is :" & $C_version32, 15, 175, 450)
	; EndIf
	; GUISetFont(10, 700, "Comic Sans MS")
	; GUICtrlCreateLabel("64 bit Java Details", 525, 150, 250)
	; GUISetFont(10, 300, "Comic Sans MS")
	; If $C_version64 == "64 bit JRE/JDK is not Installed." Then
	; 	GUICtrlCreateLabel($C_version64, 525, 175, 250)
	; Else
	; 	GUICtrlCreateLabel("64 bit JRE/JDK version is :" & $C_version64, 525, 175, 250)
	; EndIf
	; GUICtrlCreateLabel("", 0, 1, 450)
	; GUICtrlCreateLabel("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", -1, 200)
	; GUISetFont(11, 700, "Comic Sans MS")
	; GUICtrlCreateLabel("Browser Details", 300, 225, 250)
	; GUISetFont(10, 300, "Comic Sans MS")
	; GUICtrlCreateLabel("", 0, 1, 450)
	; GUICtrlCreateLabel("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", -1, 250)
	; GUISetFont(10, 700, "Comic Sans MS")
	; GUICtrlCreateLabel("Chrome Browser Details", 15, 275, 150)
	; GUISetFont(10, 300, "Comic Sans MS")
	; If $Chrome == "Chrome is not installed." Then
	; 	GUICtrlCreateLabel($Chrome, 15, 275, 250)
	; Else
	; 	GUICtrlCreateLabel("Chrome version is :" & $Chrome, 15, 275, 450)
	; EndIf
	; GUISetFont(10, 700, "Comic Sans MS")
	; GUICtrlCreateLabel("FireFox Browser Details", 525, 275, 150)
	; GUISetFont(10, 300, "Comic Sans MS")
	; If $FireFox == "FireFox is not installed." Then
	; 	GUICtrlCreateLabel($FireFox, 525, 275, 250)
	; Else
	; 	GUICtrlCreateLabel("FireFox version is :" & $FireFox, 525, 275, 450)
	; EndIf
	; GUISetFont(10, 700, "Comic Sans MS")
	; GUICtrlCreateLabel("Internet Explorer Browser Details", 15, 300, 250)
	; GUISetFont(10, 300, "Comic Sans MS")
	; GUICtrlCreateLabel("Internet Explorer version is :" & $svcVersion, 15, 300, 450)
	; GUICtrlCreateLabel("", 0, 1, 450)
	; GUICtrlCreateLabel("------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", -1, 325)
	; GUISetFont(10, 700, "Comic Sans MS")
	; GUICtrlCreateLabel("Detailed Text File : " & @DesktopDir & "\" & $FileName, 15, 350)
	; GUICtrlCreateLabel("Detailed HTML file : " & $FileHandler1, 15, 390)
	; GUISetFont(9, 700, "Comic Sans MS")
	; $viewReportTxt = GUICtrlCreateButton("ViewReportTxt", 650, 340, 110, 30)
	; $viewReportHtm = GUICtrlCreateButton("ViewReportHTML", 650, 380, 110, 30)
	; GUICtrlSetBkColor($viewReportTxt, $COLOR_YELLOW)
	; GUICtrlSetBkColor($viewReportHtm, $COLOR_YELLOW)
	; GUISetState(@SW_SHOW, $Mainpage)
	; GUISetOnEvent($GUI_EVENT_CLOSE, "OnExit")
	; GUICtrlSetOnEvent($viewReportTxt, "TextButton")
	; GUICtrlSetOnEvent($viewReportHtm, "HTMLButton")
	;
	; Func OpenReportTxt()
	; 	Local Const $sFilePath = @DesktopDir & "\" & $FileName
	; 	Run("notepad.exe " & $sFilePath, @WindowsDir, @SW_MAXIMIZE)
	; EndFunc   ;==>OpenReportTxt
	; Func OpenReportHTML()
	; 	_IECreate($FileHandler1)
	; EndFunc   ;==>OpenReportHTML
#CE


#CS Func TextButton()
	; 	OpenReportTxt()
	; EndFunc   ;==>TextButton
	;
	; Func HTMLButton()
	; 	OpenReportHTML()
	; EndFunc   ;==>HTMLButton
#CE

#CS While 1
	; 	Sleep(1000)
	; WEnd
	; Func OnExit()
	; 	Exit
	; EndFunc   ;==>OnExit
#CE


