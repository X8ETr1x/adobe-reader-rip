############################Install_AcroRead.ps1###############################

#Last updated: 2-10-2011
#Required files: Acrobat Pro administrative installation, MSP patch files
#Purpose: To install Adobe Acrobat Reader
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

#Converts the working directory to the Windows file system:
$dir = [Environment]::CurrentDirectory=
(Get-Location -PSProvider FileSystem).ProviderPath

#Installs the base application:
$Exec = 'msiexec.exe'
$Option = " /i $dir\AcroRead.msi TRANSFORMS=$dir\AcroRead.mst /qb"
$Install = [Diagnostics.Process]::Start($Exec, $Option)
$Install.WaitForExit()

#If out-of-cycle security patches are available, then enable the below code.
#It is important to differentiate between quarterly and OOC patches. Quarterly
#patches can be applied directly to upgrade an administrative installation.
#OOC patches must be applied to the destination system seperately, as they
#cannot be applied to administrative installations.

<#Installs the most recent patch:
$Exec = 'msiexec.exe'
$Option = " /p patch.msp /qn"
$Install = [Diagnostics.Process]::Start($Exec, $Option)
$Install.WaitForExit()#>

############################End of Script######################################
