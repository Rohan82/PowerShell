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
        WriteLine
        Write-Host -ForegroundColor White " - Loading SharePoint PowerShell Snapin..."
        # Added the line below to match what the SharePoint.ps1 file implements (normally called via the SharePoint Management Shell Start Menu shortcut)
        If (Confirm-LocalSession) {$Host.Runspace.ThreadOptions = "ReuseThread"}
        Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop | Out-Null
        WriteLine
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

#Region Function: Restart-Server
Function Restart-Server
{
	param
	(
		[Parameter(Mandatory=$true)][string]$server,
		[Parameter(Mandatory=$true)][string]$logFolder
	)

	$ErrorActionPreference = "SilentlyContinue"

	try
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
			    "Failed to reboot $server"
			    exit
			}#exit script and log failed to reboot.
			
		    $i++
			"Wait for server to reboot"
		    Start-Sleep -Seconds 15
		}#end do
		
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
				Export-Csv D:\RebootResults.csv -InputObject $newrow -Append -Force
	    		"$server did not come online"
	    		exit
			}#exit script and log failed to come online.
	    	$i++
	    	"Wait for [$server] to come online"
	    	Start-Sleep -Seconds 15
		}#end do
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
		Export-Csv D:\RebootResults.csv -InputObject $newrow -Append -Force

	}#End try.

	Catch
	{
		$errMsg = $_.Exception
		"Failed with $errMsg"
	}
}
#EndRegion Function: Restart-Server

#Region Function: Enable-ServerRemotePowerShell
function Enable-ServerRemotePowershell
{
	try
	{
		# Enable PSRemoting
		Write-Host "Enabling PSRemoting..." -ForegroundColor Yellow -NoNewline
		Enable-PSRemoting -Confirm:$false | Out-Null
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
		Write-Host "Failed to enable Remote PowerShell on $env:COMPUTERNAME. See the following error for details." -ForegroundColor Red
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
		$siteUrl,
		[Parameter(Mandatory=$true)]$OutputFolder
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

					# Add the current site info to the site information array
					$siteInfos += $siteInfo
				}
			}
		}
	}
	catch [system.exception]
	{
		Write-Host "Encountered an error while building the site information list. See error message for more information." -ForegroundColor Red
		Write-Error $_
		exit
	}
	Write-Host "Done" -ForegroundColor Green

	# Export site information to CSV file
	if (!(Test-Path $OutputFolder))
	{
		# Create the specified directory for the SiteInfo.csv file
		New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
	}
	$siteInfos | Export-CSV -NoTypeInformation $OutputFolder\SiteInfo.csv
	Write-Host "Finished exporting site collection database info. You can view the CSV file here:" -ForegroundColor Yellow
	Write-Host "$OutputFolder\SiteInfo.csv"
}
#EndRegion Function: Get-SPSiteDBInfo

<#.Synopsis
<!<SnippetShortDescription>!>
.DESCRIPTION
<!<SnippetLongDescription>!>
.EXAMPLE
<!<SnippetExample>!>
.EXAMPLE
<!<SnippetAnotherExample>!>
#>
