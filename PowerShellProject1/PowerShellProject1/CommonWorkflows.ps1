#
# CommonWorkflows.ps1
#
#Region Workflow: ResumeWorkflow
workflow Resume_Workflow
{
    # Rename-Computer -NewName some_name -Force -Passthru
    Restart-Computer -Wait
    # Do some stuff
	InlineScript
	{
		Write-Host "Completed computer restart."
	}
}
# Create the scheduled job properties
$options = New-ScheduledJobOption -RunElevated -ContinueIfGoingOnBattery -StartIfOnBattery
$secpasswd = ConvertTo-SecureString "" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential ("US\hinmatr", $secpasswd)
$AtStartup = New-JobTrigger -AtStartup

# Register the scheduled job
Register-ScheduledJob -Name Resume_Workflow_Job -Trigger $AtStartup -ScriptBlock ({[System.Management.Automation.Remoting.PSSessionConfigurationData]::IsServerManager = $true; Import-Module PSWorkflow; Resume-Job -Name new_resume_workflow_job -Wait}) -ScheduledJobOption $options
# Execute the workflow as a new job
Resume_Workflow -AsJob -JobName new_resume_workflow_job
#EndRegion Workflow: ResumeWorkflow