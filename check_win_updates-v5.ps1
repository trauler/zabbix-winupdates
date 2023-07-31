# Change $ZabbixInstallPath to wherever your Zabbix 2 Agent is installed

$ZabbixInstallPath = "$Env:Programfiles\Zabbix Agent 2"
$ZabbixConfFile = "$Env:Programfiles\Zabbix Agent 2"

$returnStateOK = 0
$returnStateWarning = 1
$returnStateUnknown = 3
$returnStateOptionalUpdates = $returnStateWarning
$Senderarg0 = "$ZabbixInstallPath\zabbix_sender.exe"
$Senderarg1 = '-vv'
$Senderarg2 = '-c'
$Senderarg3 = "$ZabbixConfFile\zabbix_agent2.conf"
$Senderarg4 = '-i'
$SenderargUpdateReboot = '\updatereboot.txt'
$Senderarglastupdated = '\lastupdated.txt'
$Senderargcountcritical = '\countcritical.txt'
$SenderargcountOptional = '\countOptional.txt'
$SenderargcountHidden = '\countHidden.txt'
$Countcriticalnum = '\countcriticalnum.txt'
$Senderarg5 = '-k'
$Senderargupdating = 'Winupdates.Updating'
$Senderarg6 = '-o'
$Senderarg7 = '0'

# Last update and write to tmp file

$windowsUpdateObject = New-Object -ComObject Microsoft.Update.AutoUpdate
Write-Output "- Winupdates.LastUpdated $($windowsUpdateObject.Results.LastInstallationSuccessDate)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderarglastupdated

# Get rebbot status and write to tmp file

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"){ 
	Write-Output "- Winupdates.Reboot 1" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargUpdateReboot
    Write-Host "`t There is a reboot pending" -ForeGroundColor "Red"
}else {
	Write-Output "- Winupdates.Reboot 0" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargUpdateReboot
    Write-Host "`t No reboot pending" -ForeGroundColor "Green"
		}

# Checks available windows updates

$updateSession = new-object -com "Microsoft.Update.Session"
$updates=$updateSession.CreateupdateSearcher().Search(("IsInstalled=0 and Type='Software'")).Updates

$criticalTitles = "";
$countCritical = 0;
$countOptional = 0;
$countHidden = 0;

# Count available updates

foreach ($update in $updates) {
	if ($update.IsHidden) {
		$countHidden++
	}
	elseif ($update.AutoSelectOnWebSites) {
		$criticalTitles += $update.Title + " `n"
		$countCritical++
	} else {
		$countOptional++
	}
}

# If no updates required, write it to tmp file and send to zabbix

if ($updates.Count -eq 0) {

	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$Countcriticalnum
	Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden
    Write-Host "`t There are no pending updates" -ForeGroundColor "Green"
	
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7 -s "$env:computername"
	
	exit $returnStateOK
}

# Count each type of updates and write it tio tmp file

if (($countCritical + $countOptional) -gt 0) {

	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$Countcriticalnum
	Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden
    Write-Host "`t There are $($countCritical) critical updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countOptional) optional updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countHidden) hidden updates available" -ForeGroundColor "Yellow"
	
    & $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7 -s "$env:computername"
}

if ($countOptional -gt 0) {
	exit $returnStateOptionalUpdates
}

if ($countHidden -gt 0) {
	
	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$Countcriticalnum
	Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden
    Write-Host "`t There are $($countCritical) critical updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countOptional) optional updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countHidden) hidden updates available" -ForeGroundColor "Yellow"
	
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden -s "$env:computername"
	& $Senderarg0 $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7 -s "$env:computername"
	
	exit $returnStateOK
}

exit $returnStateUnknown
