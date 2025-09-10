<#
.SYNOPSIS
	Checks for legacy versions of Adobe Acrobat Reader and any optional 
	components.

.DESCRIPTION
	The script searches multiple locations for various legacy versions 
	of the software and initiates either an uninstall action or a removal
	action, depending on the age. The most recent versiohn will then be installed.

.NOTES
	- The remote registry must be enabled.
	- The script must be run with "Unrestricted" mode for group policy to 
	  process it from SYSVOL. DO NOT USE BYPASS.
	- The script has been tested against Windows 2000 Professional, 
	  Windows XP SP3, Windows Vista SP2, and Windows 7 SP1. Windows NT 4 has not 
	  been tested.
	- Updated: 07/26/2010
#>

#Tells the script to continue, even on error.
$ErrorActionPreference = 'SilentlyContinue'

#Script Functions
################################################################################

#This function is designed for software that has a GUID in its uninstall string.
#It will obtain the version string of the requested software installation.
Function Get-COMSoftwareVersion
	{
	Param
		(
		[Parameter(mandatory=$true, position=0)]
		[string]$DisplayName
		)
	
	#Checks whether the system is 32-bit or 64-bit:
	$WinArch = [System.Environment]::GetEnvironmentVariable("PROCESSOR_" + `
	"ARCHITECTURE")	
	
	If ($WinArch -eq 'x86')
		{
		$RegPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
		}
	ElseIf ($WinArch -eq 'AMD64')
		{
		$RegPath = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion' + `
		'\Uninstall'
		}
	
	#Opens the HKLM\ root and adds the uninstall path:
	$Key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($RegPath) 
	
	#Creates an array of all subkeys in $Key:
	$arrSubKeys = $Key.GetSubKeyNames()
	
	#Checks each subkey to see if it matches the software package:
	ForEach ($subKey in $arrSubKeys)
	 	{
		[string]$Key = $Key
		$Key = $Key.Replace("HKEY_LOCAL_MACHINE","HKLM:")
		$Path = $Key + '\' + $subKey
		$PropertyValue = (Get-ItemProperty -Path $Path -Name DisplayName)
		If ($PropertyValue -eq $DisplayName)
		 	{
		 	$PropertyValue = (Get-ItemProperty -Path $Path -Name DisplayVersion)
			return $PropertyValue
			}
		}
	}
	
#This function is designed for software that has a GUID in its uninstall string.
#It will search the registry for the software and initiate a silent uninstall.
Function Remove-COMSoftware
	{
	Param
		(
		[Parameter(mandatory=$true, position=0)]
		[string]$RegProperty
		,
		[Parameter(mandatory=$true, position=1)]
		[string]$RegValue
		)
	
	#Checks to see whether the system is 32-bit or 64-bit:
	$WinArch = [System.Environment]::GetEnvironmentVariable("PROCESSOR_" + `
	"ARCHITECTURE")	
	
	If ($WinArch -eq 'x86')
		{
		$RegPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
		}
	ElseIf ($WinArch -eq 'AMD64')
		{
		$RegPath = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion' + `
		'\Uninstall'
		}
	
	#Opens the HKLM\ root and adds the uninstall path:	
	$Key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($RegPath) 

	#Creates an array of all subkeys in $Key:
	$arrSubKeys = $Key.GetSubKeyNames()
	
	ForEach ($subKey in $arrSubKeys)
	 	{
		#Concatenates the strings to form the full registry path of the software 
		[string]$subKeyPath = "$Key" + '\' + "$subKey"
		$subKeyPath = $subKeyPath.Replace( 'HKEY_LOCAL_MACHINE\','HKLM:\')
		
		#Checks the display name of each subkey:
		[string]$searchValue = Get-ItemProperty -Path $subKeyPath `
		-Name $RegProperty
		
		#Matches the search result and the Value provided:
		If ($searchValue.Contains($RegValue))
		 	{
			#Checks the uninstal string from the registry:
			[string]$unInstString = Get-ItemProperty -Path $subKeyPath `
			-Name UninstallString
			
			#Obtains the substring that contains the GUID:
			[int]$GuidRoot = $unInstString.IndexOf('{', 3)
			[string]$GUID = $unInstString.SubString($GuidRoot, 38)
			$GUID = "/x " + $GUID + " /qn"
			[string]$MSIEXEC = 'msiexec.exe'
			
			#Runs the uninstall using the program name and arguments:
			$Uninstall = [Diagnostics.Process]::Start("$MSIEXEC","$GUID")
			$Uninstall.WaitForExit()
			}
		}
	}

#This function is designed for software that does no have a GUID and is not 
#managed by InstallShield. It searches the registry for the requested software
#and removes it by obtaining the uninstall string.
Function Remove-NonCOMSoftware
	{
	Param
		(
		[Parameter(mandatory=$true, position=0)]
		[string]$Application
		,
		[Parameter(mandatory=$true, position=1)]
		[string]$Property
		,
		[Parameter(mandatory=$true, position=2)]
		[string]$PropertyValue
		)
	
	#Checks to see whether the system is 32-bit or 64-bit:
	$WinArch = [System.Environment]::GetEnvironmentVariable("PROCESSOR_" + `
	"ARCHITECTURE")	
	
	If ($WinArch -eq 'x86')
		{
		$Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
		}
	ElseIf ($WinArch -eq 'AMD64')
		{
		$Path = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
		}
	#Cancatenates the SOFTWARE path and the requested application:
	$Path = $Path + '\' + $Application
	
	#Opens the HKLM\ root and adds the uninstall path:
	[string]$Key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($Path)
	
	#Alters the $Key path to match a mapped registry drive:
	$Key = $Key.Replace("HKEY_LOCAL_MACHINE","HKLM:")
	
	#Sets the subkey property value and formats it as a usable string:
	[string]$KeyProperty = (Get-ItemProperty -Path $Key -Name $Property)
	$KeyProperty = $KeyProperty.Replace("@{DisplayName=","")
	$KeyProperty = $KeyProperty.Replace("}","")
	
	#Matches the obtained value with the provided value:
	If ($KeyProperty -eq $PropertyValue)
		{
		#Obtains the uninstall string from the registry and formats it: 
		[string]$unInstExe = Get-ItemProperty -Path $Key -Name "UninstallString"
		$unInstExe = $unInstExe.Replace("@{UninstallString=","")
		$unInstExe = $unInstExe.Replace("}","")
		
		#Sets executable options (-a generally means silent uninstall)
		[string]$UnInstOpt = '-a'
		
		#Performs uninstall of requested software:
		$Uninstall = [Diagnostics.Process]::Start("$unInstExe","$UnInstOpt")
		$Uninstall.WaitForExit()
		}
	}

#This function is designed for software that does not have a GUID, but is
#managed by InstallShield (isuninst.exe). It will locate the software in the
#registry and use the uninstall string to perform a silent uninstall.
Function Remove-IsUninstallSoftware
	{
	Param
		(
		[Parameter(mandatory=$true, position=0)]
		[string]$Application 
		,
		[Parameter(mandatory=$true, position=1)]
		[string]$Property 
		,
		[Parameter(mandatory=$true, position=2)]
		[string]$PropertyValue 
		)
	
	#Checks if the system is 32-bit or 64-bit:
	$WinArch = [System.Environment]::GetEnvironmentVariable("PROCESSOR_" + `
	"ARCHITECTURE")
	
	If ($WinArch -eq 'x86')
		{
		$Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
		}
	ElseIf ($WinArch -eq 'AMD64')
		{
		$Path = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion' + `
		'\Uninstall'
		}
	
	#Sets the registry path to HKLM\ and the SOFTWARE path and formats it for a 
	#mapped registry drive:
	[string]$Key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($Path)
	$Key = $Key + '\' + $Application
	$Key = $Key.Replace("HKEY_LOCAL_MACHINE","HKLM:")
	
	#Obtains the subkey property value and formats it:
	[string]$regKey = (Get-ItemProperty -Path $Key -Name $Property)
	$RegKey = $RegKey.Replace("@{$Property=","")
	$RegKey = $RegKey.Replace("}","")
	
	#Matches the obtained value with the provided value:
	If ($regKey -eq $PropertyValue)	
		{
		#Sets the InstallShield executable:
		[string]$ISUNINST = 'C:\Windows\IsUninst.exe'
		
		#Obtains the uninstall string from the registry and formats it:
		[string]$UninstString = (Get-ItemProperty -Path $Key `
		-Name "UninstallString")
		$UninstString = $UninstString.Replace("}","")
		
		#Indexes and obtains the arguments from the uninstall string:
		$IsUninstOptStart = $UninstString.IndexOf("-f",0)
		[string]$IsUninstOpt = $UninstString.SubString($IsUninstOptStart)
		
		#Adds the silent option to the arguments:
		$IsUninstOpt = '-a ' + "$IsUninstOpt"
		
		#Performs the uninstall of the requested software:
		$Uninstall = [Diagnostics.Process]::Start("$ISUNINST","$IsUninstOpt")
		$Uninstall.WaitForExit()
		}
	}

#This function tests whether or not a registry path exists by providing a path 
#and value to test for. It will return a boolean value.
Function Test-PathReg
	{
	Param 
		(
		[Parameter(mandatory=$true, position=0)]
		[string]$RegKey
		,
		[Parameter(mandatory=$true, position=1)]
		[string]$Property
		)
	
	#Checks to see if the system is 32-bit or 64-bit: 
	$WinArch = [System.Environment]::GetEnvironmentVariable("PROCESSOR_" + `
	"ARCHITECTURE")
	
	If ($WinArch -eq 'x86')
		{
		$Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
		}
	ElseIf ($WinArch -eq 'AMD64')
		{
		$Path = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion' + `
		'\Uninstall'
		}
	
	#Opens the HKLM root and adds the uninstall path:
	[string]$Key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($Path)
	
	#Concatenates the $Key and provided software subkey, then formats it with a
	#mapped registry drive:
	$Path = $Key + '\' + $RegKey
	$Path = $Path.Replace('HKEY_LOCAL_MACHINE','HKLM:')
	
	#Obtains the name of the provided software key and compares it to the
	#provided software name:
	$compare = (Get-ItemProperty -LiteralPath $Path).psbase.members `
	| %{$_.name} `
	| compare $Property -IncludeEqual -ExcludeDifferent

	If ($compare.SideIndicator -like "==")
		{
		return $true
		}
	Else
		{
		return $false
		}
	}

#Main Script
################################################################################

#This section checks the registry to see if the current version of 
#Acrobat Reader (9.3.3) is installed.

#Checks WMI to see if the software is installed:
[string]$AdobeReader = Get-WmiObject -Class Win32_Product -ComputerName . `
| Where-Object -FilterScript {$_.Name -eq "Adobe Reader 9.3.3"} `
| Select Version 
$AdobeReader = $AdobeReader.Replace("@{Version=","")
$AdobeReader = $AdobeReader.Replace("}","")

If ($AdobeReader -eq "") 
	{
	#The current version of Adobe Reader is not installed. The script will now
	#check for older software packages related to Adobe Reader.
	
	#This section checks to see if the 32-bit version of Acrobat Reader 
	#3.x is installed and removes it if found.
	
	$program = "Adobe Acrobat Reader 3.01"
	$programProperty = 'DisplayName'
	$progPropertyValue = 'Adobe Acrobat Reader 3.01'
	
	$programExists = (Test-PathReg $program $programProperty)
	
	If ($programExists -eq $true)
		{
		Remove-IsUninstallSoftware $program $programProperty $progPropertyValue
		}
	
	#This section checks to see if Acrobat Reader 4.x is installed
	#and removes it if found.
	
	$program = "Adobe Acrobat 4.0"
	$programProperty = 'DisplayVersion'
	$progPropertyValue = '4.0'
	
	$regPathExists = (Test-PathReg $program $programProperty)
	
	If ($regPathExists -eq $true)
		{
		Remove-IsUninstallSoftware $program $programProperty $progPropertyValue
		
		#checks to see if shared files were removed:
		$filePath = Test-Path "C:\Program Files\Common Files\Adobe\Web\"
		
		If ($filePath = $true)
	        {
			$WinArch = [System.Environment]::GetEnvironmentVariable("PROC" + `
			"ESSOR_ARCHITECTURE")
			If ($WinArch -eq 'x86')
				{
				Remove-Item -Force -Recurse "C:\Program Files\Common Files\" + `
				"Adobe\Web"
				}
			ElseIf ($WinArch -eq 'AMD64')
				{
				Remove-Item -Force -Recurse "C:\Program Files (x86)" + `
				"\Common Files\Adobe\Web"
				}
	        }
		}
	
	#This section checks to see if Acrobat Reader 5.x is installed
	#and removes it if found.
	
	$program = "Adobe Acrobat 5.0"
	$programProperty = 'DisplayVersion'
	$progPropertyValue = '5.0'
	
	$regPathExists = (Test-PathReg $program $programProperty)
	
	If ($regPathExists -eq $true)
		{
		Remove-IsUninstallSoftware $program $programProperty $progPropertyValue
		
		#checks to see if shared files were removed:
		$filePath = Test-Path 'C:\Program Files\Common Files\Adobe\Web\'
		If ($filePath -eq $true)
	        {
	    	$WinArch = [System.Environment]::GetEnvironmentVariable("PROC" + `
			"ESSOR_ARCHITECTURE")
			If ($WinArch -eq 'x86')
				{
				Remove-Item -Force -Recurse 'C:\Program Files' + `
				'\Common Files\Adobe\Web'
				}
			ElseIf ($WinArch -eq 'AMD64')
				{
				Remove-Item -Force -Recurse 'C:\Program Files (x86)' + `
				'\Common Files\Adobe\Web'
				}
	        }
		}
	
	#This section checks to see if Acrobat Reader 6-9.x are installed
	#and removes them if found.
	
	$programProperty = 'DisplayName'
	$progPropertyValue = 'Adobe Reader'
	
	Remove-COMSoftware $programProperty $progPropertyValue
	
	#This section checks to see if Acrobat Reader 6 updates are installed
	#and removes them if found. Updates for this version each remained
	#installed, even after installing the most recent update.

	$programProperty = 'Comments'
	$progPropertyValue = 'Adobe Acrobat - Reader 6'
	
	Remove-COMSoftware $programProperty $progPropertyValue
	
	#This section checks to see if Adobe Atmosphere Player is installed
	#and removes it if found. This program was for Adobe Reader 6.0.1, and
	#cannot be silently uninstalled.
	
	$program = 'Adobe Atmosphere Player'
	$programProperty = 'DisplayName'
	
	$regPathExists = (Test-PathReg $program $programProperty)
	
	If ($regPathExists -eq $true)
		{
		$WinArch = [System.Environment]::GetEnvironmentVariable("PROC" + `
		"ESSOR_ARCHITECTURE")
		
		If ($WinArch -eq 'x86')
			{
			Remove-Item -LiteralPath 'C:\Windows\atmoUn.exe' -Force
			Remove-Item -LiteralPath 'C:\Program Files\Adobe\Acrobat 7.0' + `
			'\Setup Files\Atmosphere3D' -Force -Recurse
			Remove-Item -LiteralPath 'C:\Program Files\Viewpoint' + `
			-Force -Recurse
			Remove-Item -LiteralPath 'C:\Program Files\Adobe\Acrobat 6.0' + `
			'\Reader\plug_ins\Multimedia\MPP\AtmosphereMPP.mpp' -Force
			Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows" + `
			"\CurrentVersion\Uninstall\Adobe Atmosphere Player" -Force -Recurse
			Remove-Item -Path "HKLM:\SOFTWARE\Adobe\Atmosphere Player" + `
			-Force -Recurse
			Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows" + `
			"\CurrentVersion\App Management\ARPCache" + `
			"\Adobe Atmosphere Player" -Force -Recurse
			Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows" + `
			"\CurrentVersion\App Paths\atmosphere.dll" -Force -Recurse
			}
		ElseIf ($WinArch -eq 'AMD64')
			{
			Remove-Item -LiteralPath 'C:\Windows\atmoUn.exe' -Force
			Remove-Item -LiteralPath 'C:\Program Files (x86)\Adobe" + `
			"\Acrobat 7.0\Setup Files\Atmosphere3D' -Force -Recurse
			Remove-Item -LiteralPath 'C:\Program Files (x86)\Viewpoint' + `
			-Force -Recurse
			Remove-Item -LiteralPath 'C:\Program Files (x86)\Adobe' + `
			'\Acrobat 6.0\Reader\plug_ins\Multimedia\MPP' + `
			'\AtmosphereMPP.mpp' -Force
			Remove-Item -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft" + `
			"\Windows\CurrentVersion\Uninstall" + `
			"\Adobe Atmosphere Player" -Force -Recurse
			Remove-Item -Path "HKLM:\SOFTWARE\Wow6432Node\Adobe" + `
			"\Atmosphere Player" -Force -Recurse
			Remove-Item -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft" + `
			"\Windows\CurrentVersion\App Management\ARPCache" + `
			"\Adobe Atmosphere Player" -Force -Recurse
			Remove-Item -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft" + `
			"\Windows\CurrentVersion\App Paths\atmosphere.dll" -Force -Recurse
			}
		}
	
	#This section removes the Adobe Download Manager plug-in for Internet 
	#Explorer. This has been known to have security vulnerabilities.
	
	$regKey = '{E2883E8F-472F-4fb0-9522-AC9BF37916A7}'
	$RegProperty = 'DisplayName'
	
	$regPathExists = Test-PathReg $regKey $RegProperty
	
	If ($regPathExists -eq $true)
		{
		$EXE = 'C:\WINDOWS\system32\rundll32.exe'
		$Options = '"C:\Program Files\NOS\bin\getPlus_Helper.dll"' + `
		',Uninstall /IE2883E8F-472F-4fb0-9522-AC9BF37916A7 /Get1'
		
		$Uninstall = [Diagnostics.Process]::Start($EXE, $Options)
		}
	
	#This section installs the Acrobat Reader MSI package.
	
	$MSIEXEC = 'msiexec.exe'
	$MSIEOpt = '/i \\files\deploy$\Adobe_Acrobat_Reader\AcroRead.msi' + `
	' TRANSFORMS=\\files\deploy$\Adobe_Acrobat_Reader\AcroRead.mst /qn'
	
	$Install = [Diagnostics.Process]::Start("$MSIEXEC","$MSIEOpt")
	$Install.WaitForExit()
	
	#This section prompts the user to reboot after the installation.
	
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") `
	| Out-Null
	$rebootMsg = New-Object -ComObject wscript.shell
	$result = [System.Windows.Forms.MessageBox]::Show("The system must be " + `
	"restarted to apply critical security updates to" + `
	" Adobe Acrobat Reader. Please save and close all work, then click " + `
	"the OK button.","System Updates", `
	[System.Windows.Forms.MessageBoxButtons]::YesNo, `
	[System.Windows.Forms.MessageBoxIcon]::Question, `
	[System.Windows.Forms.MessageBoxDefaultButton]::Button1)
	
	If ($result -eq [System.Windows.Forms.DialogResult]::Yes) 
		{
		Restart-Computer -Force
		}
	Else
		{
		$result = [System.Windows.Forms.MessageBox]::Show("Acrobat " + `
		"Reader will not function until a reboot is completed. " + `
		"Please restart your computer as soon as possible.", `
		[System.Windows.Forms.MessageBoxButtons]::VbOK,0)
		}							
	}
