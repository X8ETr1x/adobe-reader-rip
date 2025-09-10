############################Install_QT.ps1#####################################

#Last updated: 2-10-2011
#Required files: AppleApplicationSupport.msi, QuickTime.msi, 
#QuickTimeInstallerAdmin.exe
#Purpose: To the Apple Quicktime Player.
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

#Retrieves the architecture of the operating system:
$WA = [System.Environment]::GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")

#Converts the working directory to the Windows file system:
$dir = [Environment]::CurrentDirectory=
(Get-Location -PSProvider FileSystem).ProviderPath

#Installs the Apple Application Support:
$Exec = 'msiexec.exe'
$Option = " /i $dir\AppleApplicationSupport.MSI /qb"
$Install = [Diagnostics.Process]::Start($Exec, $Option)
$Install.WaitForExit()

#Installs Quicktime:
$Exec = 'msiexec.exe'
$Option = "/i $dir\QuickTime.msi ASUWISINSTALLED=0 " + `
"APPLEAPPLICATIONSUPPORTISINSTALLED={0C34B801-6AEC-4667-B053-03A67E2D0415}" + `
" DESKTOP_SHORTCUTs=NO QT_TRAY_ICON=NO SCHEDULE_ASUW=NO /qb"
$Install = [Diagnostics.Process]::Start($Exec, $Option)
$Install.WaitForExit()

$subKey = "Microsoft\Windows\CurrentVersion\Run"

#Disables the autorun registry entry:
If ($WinArch -eq 'x86')
	{
	Remove-ItemProperty -Path "HKLM:SOFTWARE\$subKey" -Name `
	"QuickTime Task" -Force
	}
ElseIf ($WinArch -eq 'AMD64')
	{
	Remove-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\$subKey" -Name `
	"QuickTime Task" -Force
	}

############################End of Script######################################
