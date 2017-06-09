<#
.AUTHOR
David Pearson, www.dptechjournal.net

.DESCRIPTION
Installs Read & Write 12

.NOTES
Firewall port command New-NetFireWallRule requires Win8.1 or better.
Need to get RWSettings.xml from users profile after install, and set:
StartUpWizardHasRun to true and AutoCheckForUpdates to false 
Then copy file to C:\Users\profile\AppData\Roaming\Texthelp\ReadAndWrite\12 during install
Program is licensed by adding key to registry.
#>

$currentDirectory = split-path -parent $MyInvocation.MyCommand.Definition

# Uninstall Existing Software
$SoftwareInstalls = get-wmiobject -namespace root\cimv2\sms -query "select * from SMS_InstalledSoftware where ProductName = 'Read And Write 11'"
foreach ($SoftwareInstall in $SoftwareInstalls)
	{
		$SoftwareInstall.productname
		$software = $Softwareinstall.softwarecode
		$arguments = "/x $software /qn /norestart"
		start-process msiexec.exe -ArgumentList $arguments -wait
	}

if (Test-Path "$ENV:SystemDrive\RW Admin")
{
	remove-item -Path "$ENV:SystemDrive\RW Admin" -Force -Recurse
}

# Start Install
Start-Process "$currentDirectory\Read&Write.exe" -ArgumentList "/v/qn" -Wait

reg add "HKLM\SOFTWARE\WOW6432Node\Texthelp\Read&Write" /v "ProductCode" /t REG_SZ /d "MY_LICENSE_CODE" /f

# Copy file to Default Profile that removes first run autoupdate and disables autoupdate for all New Users
if (! (Test-Path "$ENV:SystemDrive\Users\Default\AppData\Roaming\Texthelp\ReadAndWrite\12"))
{
	mkdir "$ENV:SystemDrive\Users\Default\AppData\Roaming\Texthelp\ReadAndWrite\12"
}

Copy-Item -Path "$currentDirectory\RWSettings.xml" -Destination "$ENV:SystemDrive\Users\Default\AppData\Roaming\Texthelp\ReadAndWrite\12\" -Force


# Copy file to all the user profiles that removes first run autoupdate and disables autoupdate for existing users.
$Users = Get-ChildItem -Path $ENV:SystemDrive\Users\ -Exclude "Public","Default.migrated"
foreach ($User in $Users) {
		$profile = $User.Name
	if (! (Test-Path "$ENV:SystemDrive\Users\$profile\AppData\Roaming\Texthelp\ReadAndWrite\12"))
		{
			mkdir "$ENV:SystemDrive\Users\$profile\AppData\Roaming\Texthelp\ReadAndWrite\12"
		}
	Copy-Item -Path "$currentDirectory\RWSettings.xml" -Destination "$ENV:SystemDrive\Users\$profile\AppData\Roaming\Texthelp\ReadAndWrite\12\" -Force

}

# Remove Shortcut from Public Desktop
if (Test-Path "$ENV:PUBLIC\Desktop\Read&Write.lnk")
{
	Remove-Item -Path "$ENV:PUBLIC\Desktop\Read&Write.lnk" -Force
}

# Open Firewall Ports for .exe
New-NetFirewallRule -DisplayName "Read&Write 12" -Direction Inbound -Program "${ENV:ProgramFiles(x86)}\texthelp\read and write 12\readandwrite.exe" -Protocol tcp -Action Allow
New-NetFirewallRule -DisplayName "Read&Write 12" -Direction Inbound -Program "${ENV:ProgramFiles(x86)}\texthelp\read and write 12\readandwrite.exe" -Protocol udp -Action Allow

