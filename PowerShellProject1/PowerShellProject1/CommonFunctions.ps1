#
# CommonFunctions.ps1
#
#Region Function: Load-SharePointPowerShell
<#
.SYNOPSIS
	Adds the SharePoint PowerShell Snapin
.DESCRIPTION
	Adds the SharePoint PowerShell snapin and sets the thread options.
	Dependencies: "Write-Line" and "Confirm-LocalSession" functions.
.EXAMPLE
	Load-SharePointPowerShell
#>
Function global:Load-SharePointPowerShell
{
    If ((Get-PsSnapin |?{$_.Name -eq "Microsoft.SharePoint.PowerShell"})-eq $null)
    {
        Write-Line
        Write-Host -ForegroundColor White " - Loading SharePoint PowerShell Snapin..." -NoNewline
        # Added the line below to match what the SharePoint.ps1 file implements (normally called via the SharePoint Management Shell Start Menu shortcut)
        If (Confirm-LocalSession) {$Host.Runspace.ThreadOptions = "ReuseThread"}
        Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop | Out-Null
		Write-Host "Done" -ForegroundColor Green
        Write-Line
    }
	else
	{
		Write-Host -ForegroundColor White " - SharePoint PowerShell Snapin is already loaded."
	}
}
#EndRegion Function: Load-SharePointPowerShell

#Region Function: Confirm-LocalSession
<#
.SYNOPSIS
	Confirms PowerShell is being run on a local server.
.DESCRIPTION
	Confirms PowerShell is being run on a local server.
	Returns $false if we are running over a PS remote session, otherwise it returns $true.
.EXAMPLE
	if (Confirm-LocalSession) {Do stuff}
.EXAMPLE
	if (Confirm-LocalSession) {Do stuff} else {Do different stuff}
#>
Function global:Confirm-LocalSession
{
    # Another way
    # If ((Get-Process -Id $PID).ProcessName -eq "wsmprovhost") {Return $false}
    If ($Host.Name -eq "ServerRemoteHost") {Return $false}
    Else {Return $true}
}
#EndRegion Function: Confirm-LocalSession

#Region Function: Write-Line
<#
.SYNOPSIS
	Writes a nice line of dashes across the screen
.DESCRIPTION
	Writes a nice line of dashes across the screen
.EXAMPLE
	Write-Line
#>
Function global:Write-Line
{
    Write-Host -ForegroundColor White "--------------------------------------------------------------"
}
#EndRegion Function: Write-Line

#Region Function: Restart-RemoteServer
<#
.SYNOPSIS
	Restarts the specified remote server.
.DESCRIPTION
	Restarts the specified remote server using the -Force parameter. This will forcibly close all running
	applications on the target machine. All logon sessions will also be terminated without notice. Use the
	WhatIf switch to test the remote restart function without affecting the target machine.
.EXAMPLE
	Restart-RemoteServer -Server "server.domain.com" -Credential $cred -LogFolder "D:\temp"
.EXAMPLE
	Restart-RemoteServer -Server "server.domain.com" -Credential $cred -LogFolder "D:\temp" -WhatIf
#>
Function Restart-RemoteServer
{
	param
	(
		[Parameter(Mandatory=$true)]
		[string]$Server,
		[Parameter(Mandatory=$true)]
		[System.Management.Automation.CredentialAttribute()]
		$Credential,
		[Parameter(Mandatory=$true)]
		[string]$LogFolder,
		[switch]$WhatIf
	)

	$ErrorActionPreference = "SilentlyContinue"

	try
	{
		if ($WhatIf)
		{
			Restart-Computer -ComputerName $server -Credential $Credential -Force -WhatIf
			exit
		}
		else
		{
			$LastReboot = Get-EventLog -ComputerName $server -LogName system | ?{$_.EventID -eq '6005'} | Select -ExpandProperty TimeGenerated | select -first 1
			(Invoke-WmiMethod -ComputerName $server -Path "Win32_Service.Name='HealthService'" -Name PauseService).ReturnValue | Out-Null
			Restart-Computer -ComputerName $server -Force

			#New loop with counter, exit script if server did not reboot.
			$max = 20;$i = 0
			do
			{
		 		if($i -gt $max)
				{
					$hash = @{
						 "Server" =  $server
						 "Status" = "FailedToReboot!"
						 "LastRebootTime" = "$LastReboot"
						 "CurrentRebootTime" = "FailedToReboot!"
					}
					$newRow = New-Object PsObject -Property $hash
			 		$rnd = Get-Random -Minimum 5 -Maximum 40
			 		Start-Sleep -Seconds $rnd
			 		Export-Csv $logFolder\RebootResults.csv -InputObject $newrow -Append -Force
					Write-Host "Failed to reboot $server" -ForegroundColor Yellow
					exit
				}
			
				$i++
				Write-Host "Wait for server to reboot"
				Start-Sleep -Seconds 15
			}
			while(Test-path "\\$server\c$")
			$max = 20;$i = 0
			do
			{
				if($i -gt $max)
				{
					$hash = @{
						 "Server" =  $server
						 "Status" = "FailedToComeOnline!"
						 "LastRebootTime" = "$LastReboot"
						 "CurrentRebootTime" = "FailedToReboot!"
					}
					$newRow = New-Object PsObject -Property $hash
					$rnd = Get-Random -Minimum 5 -Maximum 40
					Start-Sleep -Seconds $rnd
					Export-Csv $logFolder\RebootResults.csv -InputObject $newrow -Append -Force
	    			"$server did not come online"
	    			exit
				}
	    		$i++
	    		"Wait for [$server] to come online"
	    		Start-Sleep -Seconds 15
			}
			while(-not(Test-path "\\$server\c$"))
		
			$CurrentReboot = Get-EventLog -ComputerName $server -LogName system | Where-Object {$_.EventID -eq '6005'} | Select -ExpandProperty TimeGenerated | select -first 1
			$hash = @{
					 "Server" =  $server
					 "Status" = "RebootSuccessful"
					 "LastRebootTime" = $LastReboot
					 "CurrentRebootTime" = "$CurrentReboot"
			}
			$newRow = New-Object PsObject -Property $hash
			$rnd = Get-Random -Minimum 5 -Maximum 40
			Start-Sleep -Seconds $rnd
			Export-Csv $logFolder\RebootResults.csv -InputObject $newrow -Append -Force
		}

	}
	Catch
	{
		$errMsg = $_.Exception
		Write-Error "Failed with $errMsg"
	}
}
#EndRegion Function: Restart-RemoteServer

#Region Function: Enable-ServerRemotePowerShell
function Enable-ServerRemotePowershell
{
	try
	{
		# Enable PSRemoting
		Write-Host "Enabling PSRemoting..." -ForegroundColor Yellow -NoNewline
		Enable-PSRemoting -Force -Confirm:$false | Out-Null
		Write-Host Done -ForegroundColor Green

		# Enable Cred SSP
		Write-Host "Enable CredSSP..." -ForegroundColor Yellow -NoNewline
		Enable-WSManCredSSP -Role Server -Force | Out-Null
		Write-Host "Done" -ForegroundColor Green

		# Configure optimal Shell properties
		Write-Host "Setting Shell properties..." -ForegroundColor Yellow -NoNewline
		# Increase MaxMemoryPerShellMB setting (default is 150)
		Set-Item WSMan:\localhost\shell\MaxMemoryPerShellMB 1024
		# Decrease MaxShellsPerUser and MaxConcurrentUsers (default is 5)
		Set-Item WSMan:\localhost\shell\MaxShellPerUser 2
		Set-Item WSMan:\localhost\Shell\MacConcurrentUsers 2
		Register-PSSessionConfiguration -Name "SharePoint" -StartupScript "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\CONFIG\PowerShell\Registration\SharePoint.ps1" -Force -ThreadOptions ReuseThread
		Set-PSSessionConfiguration -Name SharePoint -ShowSecurityDescriptorUI -Force
		Write-Host "Done" -ForegroundColor Green
	}
	catch [system.exception]
	{
		Write-Host "Failed to enable Remote PowerShell on $server. See the following error for details." -ForegroundColor Red
		Write-Host ""
		Write-Error $_
	}
}
#EndRegion Function: Enable-ServerRemotePowerShell

#Region Function: Start-RemotePowerShell
<#
.Synopsis
Starts a remote PowerShell Session.
.DESCRIPTION
Starts a remote PowerShell session using the supplied credentials and server name.
Requires a credential object for the credential variable.
Requires PowerShell Remoting enabled on the destination server.
.EXAMPLE
Start-RemotePowerShell -Server "server.domain.com" -Credential $cred
#>
function Start-RemotePowerShell
{
	param
	(
		[Parameter(Mandatory=$true)]
		[String]$Server,
		[Parameter(Mandatory=$true)]
		[System.Management.Automation.CredentialAttribute()]
		$Credential
	)
	
	$session = New-PSSession -ComputerName $server -Authentication Credssp -Credential $credential

	# Prompt to select remote session type
	$title = "Remote PowerShell Session Type" 
	$message = "What type of remote session would you like to start?"

	$EnterPSSession = New-Object System.Management.Automation.Host.ChoiceDescription "&Enter-PSSession", `
    "Uses the Enter-PSSession cmdlet to start the remote PowerShell session."

	$InvokeCommand = New-Object System.Management.Automation.Host.ChoiceDescription "&Invoke-Command", `
    "Uses the Invoke-Command cmdlet to start the remote PowerShell session."

	$Cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel", `
	"Cancels the remote session connection attempt and exits the script."

	$options = [System.Management.Automation.Host.ChoiceDescription[]]($EnterPSSession, $InvokeCommand, $Cancel)

	$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

	switch ($result)
		{
			0 
			{
				Write-Host "Using Enter-PSSession..." -ForegroundColor Yellow
				Enter-PSSession -Session $session
			}
			1 
			{
				Write-Host "Using Invoke-Command..." -ForegroundColor Yellow
				Invoke-Command -Session $session -ScriptBlock {Add-PSSnapin Microsoft.SharePoint.PowerShell}
			}
			2 
			{
				Write-Host "Remote session cancelled. Exiting the script." -ForegroundColor Yellow
				Exit
			}

		}
}
#Region Function: Start-RemotePowerShell

#Region Function: Get-SPSiteDBInfo
<#.Synopsis
Creates a SiteInfo.csv file with DBServer and DBAlas info for the SharePoint site collection specified.
.DESCRIPTION
Creates a SiteInfo.csv file with DBServer and DBAlas info for the SharePoint site specified. If no site
is specified, the CSV file will contain info for all site collections.
(travis.hinman@gmail.com)
.PARAMETER OutputFolder
Mandataory parameter defining the folder for the CSV file output.
.PARAMETER SiteUrl
Optional paramter specifying for which site collection the info will be collected. If no site collection
URL is provided, the function will run against all site collections.
.EXAMPLE
Get-SPSIteDBInfo -OutputFolder "D:\temp"
.EXAMPLE
Get-SPSIteDBInfo -SiteUrl https://portal.domain.com/sites/teams -OutputFolder "D:\temp"
#>
Function Get-SPSiteDBInfo
{
	param
	(
		$siteUrl
	)
	# Create a dictionary hash to translate DB server aliases to the actual instances
	$dict = @{}

	# Create an array to store the site information
	$siteInfos = @()

	# Set variables
	Write-Host "Getting list of DB server aliases and building dictionary..." -ForegroundColor Yellow -NoNewline
	$regPath = "HKLM:\SOFTWARe\Microsoft\MSSQLServer\Client\ConnectTo"
	$aliasList = Get-Item $regPath

	# Get the list of DB server aliases and add them to the dictionary hash
	try
	{
		foreach ($alias in $aliasList.Property)
		{
			$dbInstance = $aliasList.GetValue($alias)
			$dbInstance = $dbInstance.Substring($dbInstance.IndexOf(",")+1)
			#Write-Host $alias $dbInstance
			$dict.Add($alias,$dbInstance)
		}
		Write-Host "Done" -ForegroundColor Green
	}
	catch [system.exception]
	{
		Write-Host "Unable to create dictionary for DB aliases. See error message for more information." -ForegroundColor Red
		Write-Error $_
		exit
	}

	# Get the properties of each site collection in each content database
	try
	{
		Write-Host "Getting site collection information..." -ForegroundColor Yellow -NoNewline
		if ($siteUrl)
		{
			$site = Get-SPSite
			# Create a custom PSObject and add the site properties to it
			$siteInfo = New-Object PSObject
			$siteInfo | Add-Member -type NoteProperty -Name SiteUrl -Value $site.Url
			$siteInfo | Add-Member -type NoteProperty -Name ContentDB -Value $site.ContentDatabase.Name
			$siteInfo | Add-Member -type NoteProperty -Name DBServer -Value $dict.Item($site.ContentDatabase.Server)
			$siteInfo | Add-Member -type NoteProperty -Name DBAlias -Value $site.ContentDatabase.Server

			# Add the site info to the site information array
			$siteInfos += $siteInfo
		}
		else
		{
			foreach ($db in Get-SPContentDatabase)
			{
				foreach ($site in Get-SPSite -ContentDatabase $db)
				{
					# Create a custom PSObject and add the site properties to it
					$siteInfo = New-Object PSObject
					$siteInfo | Add-Member -type NoteProperty -Name SiteUrl -Value $site.Url
					$siteInfo | Add-Member -type NoteProperty -Name ContentDB -Value $site.ContentDatabase.Name
					$siteInfo | Add-Member -type NoteProperty -Name DBServer -Value $dict.Item($site.ContentDatabase.Server)
					$siteInfo | Add-Member -type NoteProperty -Name DBAlias -Value $site.ContentDatabase.Server
					#$siteinfo | Add-Member -Type NoteProperty -Name Quota -Value $site.Quota

					# Add the current site info to the site information array
					$siteInfos += $siteInfo
				}
			}
		}
	}
	catch [system.exception]
	{
		Write-Host "Encountered an error while building the site information list. See error message for more information." -ForegroundColor Red
		Write-Host ""
		Write-Error $_
		exit
	}
	Write-Host "Done" -ForegroundColor Green

	# Export site information to CSV file
	$siteInfos | Export-CSV -NoTypeInformation "$([Environment]::GetFolderPath("Desktop"))\SiteInfo.csv"
	Write-Host "Finished exporting site collection database info. You can view the CSV file here:" -ForegroundColor Yellow
	Write-Host "$([Environment]::GetFolderPath("Desktop"))\SiteInfo.csv"
}
#EndRegion Function: Get-SPSiteDBInfo

#Region Function: Get-SEPExclusions
<#
.Synopsis
	Creates two files: one with list of excluded file extensions; another with excluded directories.
.DESCRIPTION
	Creates two files: one with list of excluded file extensions; another with excluded directories.
	Script will automatically write the files to the desktop of the current logged on user, then open the files.
	Dependencies: "ConvertTo-Boolean" function
	Create By: Travis Hinman (travis.hinman@ey.com)
.EXAMPLE
	Get-SEPExclusions
#>
Function Get-SEPExclusions
{
	$dirExclusions = @()
	$dirExclusions += "The following directory exclusions were found for Symantec Endpoint Protection."
	
	# Get directory exclusions
	$regKeyDir = "HKLM:\SOFTWARE\Wow6432Node\Symantec\Symantec Endpoint Protection\AV\Exclusions\ScanningEngines\Directory\Admin"
	$dirExclusionsParent = Get-Item $regKeyDir
	$dirExclusionIDs = $dirExclusionsParent.GetSubKeyNames()

	# process each subkey
	foreach ($id in $dirExclusionIDs)
	{
		$exclusionKey = Get-Item "$regKeyDir\$id"
		$exclusionDir = $exclusionKey.GetValue("DirectoryName")
		$excludeSubDirs = $exclusionKey.GetValue("ExcludeSubDirs") | ConvertTo-Boolean
		
		$dir = New-Object PSObject
		$dir | Add-Member -Type NoteProperty -Name "Directory" -Value $exclusionDir
		$dir | Add-Member -Type NoteProperty -Name "ExcludeSubDirs" -Value $excludeSubDirs.ToString()
		$dirExclusions += $dir
	}

	# Get extenstion exclusions
	$regKeyExts = "HKLM:\SOFTWARE\Wow6432Node\Symantec\Symantec Endpoint Protection\AV\Exclusions\ScanningEngines\Extensions\Admin\Extensions"
	$extExclusions = Get-Item $regKeyExts

	# Get list of extensions from key
	$exts = ($extExclusions.GetValue("Exts")).Split(",")

	# Write exclusions to desktop
	$dirExclusions | Sort Directory | Export-CSV -NoTypeInformation "$([Environment]::GetFolderPath("Desktop"))\DirExclusions.csv"
	$exts | Sort | Out-File "$([Environment]::GetFolderPath("Desktop"))\ExtExclusions.txt" 

	# Open files to present findings
	Notepad.exe "$([Environment]::GetFolderPath("Desktop"))\DirExclusions.csv"
	Notepad.exe "$([Environment]::GetFolderPath("Desktop"))\ExtExclusions.txt"
}
#EndRegion Function: Get-SEPExclusions

#Region Function: ConvertTo-Boolean
<#
.SYNOPSIS
	Converts common string inputs to boolean values
.DESCRIPTION
	Converts common string inputs to boolean values.
	Dependencies: None
    Created By Travis Hinman (travis.hinman@ey.com)
.EXAMPLE
	$value = $variable.Property | ConvertTo-Boolean
#>
Function ConvertTo-Boolean 
{ 
	param 
	( 
		[Parameter(Mandatory=$false,ValueFromPipeline=$true)]
		[string]$value 
	) 
	switch ($value) 
	{ 
		"y" { return $true; } 
		"yes" { return $true; } 
		"true" { return $true; } 
		"t" { return $true; } 
		1 { return $true; } 
		"n" { return $false; } 
		"no" { return $false; } 
		"false" { return $false; } 
		"f" { return $false; }  
		0 { return $false; } 
	} 
} 
#EndRegion Function: ConvertTo-Boolean

#Region Function: Get-SPServerDiskSizes
<#
.SYNOPSIS
	Returns an array of disk sizes and free space for all drives on each sharepoint server in the farm.
.DESCRIPTION
	Returns an array of disk sizes and free space for all drives on each sharepoint server in the farm.
    Excludes drives with drive size of $null.
	Dependencies: "Load-SharePointPowerShell" function.
    Created By Travis Hinman (travis.hinman@ey.com)
.EXAMPLE
	Get-SPServerDiskSizes | Export-CSV "D:\temp\disksizes.csv"
#>
Function Get-SPServerDiskSizes
{
	# Ensure SharePoint PowerShell Module is loaded
	Load-SharePointPowerShell

    # Create array to hold disk information
    $driveInfo = @()

    # Get all SharePoint servers in farm
    $servers = (Get-SPServer | ?{$_.Role -ne "Invalid"}).Name
    
    # Process disk information for each disk on each server
    foreach ($serv in $servers)
    {
        $disks = Get-WmiObject Win32_LogicalDisk -ComputerName $serv | ?{$_.Size -ne $null}

        # Add disk information to array
        foreach ($disk in $disks)
        {
            $servInfo = New-Object PSObject
            $servInfo | Add-Member -Type NoteProperty -Name "ServerName" -Value $serv
            $servInfo | Add-Member -Type NoteProperty -Name "DiskSize" -Value $([Math]::Round($disk.Size/1GB,2))
            $servInfo | Add-Member -Type NoteProperty -Name "FreeSpace" -Value $([Math]::Round($disk.FreeSpace/1GB,2))
            $servInfo | Add-Member -Type NoteProperty -Name "DeviceID" -Value $disk.DeviceID
            $driveInfo += $servInfo
        }
    }
    # Return array information
    return $driveInfo
}
#EndRegion Function: Get-SPServerDiskSizes


<#.Synopsis
<!<SnippetShortDescription>!>
.DESCRIPTION
<!<SnippetLongDescription>!>
.EXAMPLE
<!<SnippetExample>!>
.EXAMPLE
<!<SnippetAnotherExample>!>
#>