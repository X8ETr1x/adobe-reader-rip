############################Install_Flash.ps1##################################

#Last updated: 2-9-2011
#Required files: Flash.msi (name is dependant on version), mms.cfg
#Purpose: To install Adobe Flash Player
#Supported Operating Systems: Windows XP Professional (32-bit),
#Windows Vista Business (32 and 64-bit), Windows 7 Professional (32 and 64-bit)

############################Installation Notes#################################

#NOTE: Make sure that the execution policy is set to unrestricted before 
#running this script. Best practice dictates setting the policy back to
#restricted after this scipt is finished.

#NOTE: If this script is going to be used on an NT6 operating system, then 
#powershell.exe must be run as administrator first for elevated priveleges. 
#Any file copy procedures will fail if this is not run as administrator.

############################Script#############################################

$ErrorActionPreference = "SilentlyContinue"

#Checks the system architecture:
$WA = [System.Environment]::GetEnvironmentVariable("PROCESSOR_" + `
"ARCHITECTURE")	

#Checks the current working directory:
$dir = [Environment]::CurrentDirectory=
(Get-Location -PSProvider FileSystem).ProviderPath

#Launches the application installation:
$Exec = 'msiexec.exe'
$Opt = "/i $dir\10.1.102.64.msi /qn"
$I = [Diagnostics.Process]::Start($Exec, $Opt)
$I.WaitForExit()

#Adds the Flash configuration file, based on architecture:
If ($WA -eq 'x86')
	{
	Copy-Item "$dir\mms.cfg" -Destination `
	c:\Windows\System32\Macromed\Flash -Force
	}
ElseIf ($WA -eq 'AMD64')
	{
	Copy-Item "$dir\mms.cfg" -Destination `
	c:\Windows\SysWOW64\Macromed\Flash -Force
	}
############################End of Script######################################
