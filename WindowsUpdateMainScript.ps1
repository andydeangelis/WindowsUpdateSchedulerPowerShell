<#
	.SYNOPSIS
		A brief description of the  file.
	
	.DESCRIPTION
		The main launcher file to remotely check for and install Windows Updates on multiple PCs based on AD OUs.
	
	.PARAMETER SearchBase
		Used by the Get-ADComputer function to return a list of servers.
	
	.PARAMETER Credential
		Takes a PS Credential object to for remote authorization.
	
	.PARAMETER InstallUpdates
		Switch parameter to confirm installing any updates found. Omitting this
		simply installs the PSWindowsUpdate PS module.
	
	.PARAMETER ListAvailableUpdates
		Switch parameter to list all available updates and dump to a excel report.
	
	.PARAMETER NumJobs
		A description of the NumJobs parameter.
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
		Created on:   	3/28/2019 12:45 PM
		Created by:   	andy-user
		Organization:
		Filename:
		===========================================================================
#>
param
(
	[Parameter(Mandatory = $true)]
	[string]$SearchBase,
	[Parameter(Mandatory = $true)]
	[System.Management.Automation.Credential()]
	[ValidateNotNull()]
	[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,
	[Parameter(Mandatory = $false)][switch]$GetUpdateHistory,
	[Parameter(Mandatory = $false)][switch]$InstallUpdates,
	[Parameter(Mandatory = $false)][switch]$ListAvailableUpdates,
	[Parameter(Mandatory = $false)]
	[ValidateRange(0, 256)]
	[int]$NumJobs
)

# Source the function files.

. ".\Test-TCPport.ps1"

# First, let's make sure there are no stale jobs.

Get-Job | Stop-Job
Get-Job | Remove-Job

# Ensure the PSWindowsUpdate, PendingReboot and ImportExcel modules are installed on the local machine.

if (-not (Get-InstalledModule -Name "PSWindowsUpdate" -ErrorAction SilentlyContinue) -or -not (Get-InstalledModule -Name "ImportExcel" -ErrorAction SilentlyContinue) -or -not (Get-InstalledModule -Name "PendingReboot" -ErrorAction SilentlyContinue))
{
	try
	{
		Set-ExecutionPolicy Unrestricted -Force
		Write-Host "Installing PSWindowsUpdate,ImportExcel and PendingReboot modules on" $ENV:COMPUTERNAME -ForegroundColor Yellow
		Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
		Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
		Install-Module -Name "PSWindowsUpdate" -Force -Scope AllUsers
		Install-Module -Name "ImportExcel" -Force -Scope AllUsers
		Install-Module -Name "PendingReboot" -Force -Scope AllUsers
		Install-Module -Name "PSParallel" -Force -Scope AllUsers
		Write-Host "All required PoSH modules successfully installed on" $ENV:COMPUTERNAME -ForegroundColor Green
		Import-Module "PSWindowsUpdate"
		Import-Module "ImportExcel"
		Import-Module "PendingReboot"
		Import-Module "PSParallel"
	}
	catch
	{
		
		Write-Host "Unable to install one or more modules on" $ENV:COMPUTERNAME ". Please install manually or resolve connectivity issues." -ForegroundColor Red
	}
}
else
{
	Write-Host "All required modules are already installed. Checking for module updates on" $ENV:COMPUTERNAME -ForegroundColor Green
	try
	{
		Set-ExecutionPolicy Unrestricted -Force
		Import-Module "PSWindowsUpdate"
		Import-Module "ImportExcel"
		Import-Module "PendingReboot"
		Import-Module "PSParallel"
		Update-WUModule -Online -Confirm:$false
		Update-Module -Name "ImportExcel" -Force -Confirm:$false
		Update-Module -Name "PendingReboot" -Force -Confirm:$false
		Update-Module "PSParallel" -Force -Confirm:$false
		Write-Host "All required modules are up to date on" $ENV:COMPUTERNAME -ForegroundColor Green
	}
	catch
	{
		Write-Host "Unable to update one or more modules on" $ENV:COMPUTERNAME ". Please install manually or resolve connectivity issues." -ForegroundColor Red
	}
}

# Get a list of computers from the correct OU in AD.
Import-Module -Name ActiveDirectory

$remoteComputers = Get-ADComputer -Credential $Credential -Filter { servicePrincipalName -notlike "*MSClusterVirtualServer*" } `
					-SearchBase "$SearchBase" | ?{ $_.Enabled -eq "True" }

# Get the date and time for the log file.

$datetime = get-date -f MM-dd-yyyy_hh.mm.ss

# Define the log file.

#$psModuleInstallLog = "$PSScriptRoot\WSUS_Reports\$datetime\PSModuleInstall_$datetime.log"
$listUpdatesXLSX = "$PSScriptRoot\WSUS_Reports\AvailableUpdateReports\$datetime\AvailableUpdates_$datetime.xlsx"
$installedUpdatesXLSX = "$PSScriptRoot\WSUS_Reports\InstalledUpdateReports\$datetime\InstalledUpdates_$datetime.xlsx"
$updateHistorytXLSX = "$PSScriptRoot\WSUS_Reports\UpdateHistoryReports\$datetime\UpdateHistory_$datetime.xlsx"

<#if (-not (Get-ChildItem -Path "$PSScriptRoot\WSUS_Reports\$datetime" -ErrorAction SilentlyContinue))
{
	try	{ New-Item -ItemType Directory -Name "WSUS_Reports" -Path $PSScriptRoot -ErrorAction Stop }
	catch{ "Directory $PSScriptRoot\WSUS_Reports already exists or access is denied." }
	
	try { New-Item -ItemType Directory -Name "$datetime" -Path "$PSScriptRoot\WSUS_Reports" -ErrorAction Stop }
	catch { "Directory $PSScriptRoot\WSUS_Reports\$datetime already exists or access is denied." }
	
	try
	{
		New-Item -ItemType File -Name "PSModuleInstall_$datetime.log" -Path "$PSScriptRoot\WSUS_Reports\$datetime" -ErrorAction Stop
	}
	catch { "Unable to create the file $psModuleInstallLog. File already exists or access is denied." }
}/#>

if (-not (Get-ChildItem -Path "$PSScriptRoot\WSUS_Reports\$datetime" -ErrorAction SilentlyContinue))
{
	try { New-Item -ItemType Directory -Name "WSUS_Reports" -Path $PSScriptRoot -ErrorAction Stop }
	catch { "Directory $PSScriptRoot\WSUS_Reports already exists or access is denied." }
	
	if ($ListAvailableUpdates)
	{
		try { New-Item -ItemType Directory -Name "AvailableUpdateReports" -Path "$PSScriptRoot\WSUS_Reports" -ErrorAction Stop }
		catch { "Directory $PSScriptRoot\WSUS_Reports\AvailableUpdateReports already exists or access is denied." }
		
		try { New-Item -ItemType Directory -Name "$datetime" -Path "$PSScriptRoot\WSUS_Reports\AvailableUpdateReports" -ErrorAction Stop }
		catch { "Directory $PSScriptRoot\WSUS_Reports\AvailableUpdateReports\$datetime already exists or access is denied." }
	}
	
	if ($InstallUpdates)
	{
		try { New-Item -ItemType Directory -Name "InstalledUpdateReports" -Path "$PSScriptRoot\WSUS_Reports" -ErrorAction Stop }
		catch { "Directory $PSScriptRoot\WSUS_Reports\InstalledUpdateReports already exists or access is denied." }
		
		try { New-Item -ItemType Directory -Name "$datetime" -Path "$PSScriptRoot\WSUS_Reports\InstalledUpdateReports" -ErrorAction Stop }
		catch { "Directory $PSScriptRoot\WSUS_Reports\InstalledUpdateReports\$datetime already exists or access is denied." }
	}
	
	if ($GetUpdateHistory)
	{
		try { New-Item -ItemType Directory -Name "UpdateHistoryReports" -Path "$PSScriptRoot\WSUS_Reports" -ErrorAction Stop }
		catch { "Directory $PSScriptRoot\WSUS_Reports\UpdateHistoryReports already exists or access is denied." }
		
		try { New-Item -ItemType Directory -Name "$datetime" -Path "$PSScriptRoot\WSUS_Reports\UpdateHistoryReports" -ErrorAction Stop }
		catch { "Directory $PSScriptRoot\WSUS_Reports\UpdateHistoryReports\$datetime already exists or access is denied." }
	}
}

# if the -NumThreads parameter is not set, we are going to determine the number of simultaneous jobs 
# to run based on the number of threads available.

if (-not $NumJobs)
{
	$processors = get-wmiobject -computername localhost Win32_ComputerSystem
	$NumJobs = 0
	try
	{
		$NumJobs = @($processors).NumberOfLogicalProcessors
	}
	catch
	{
		$NumJobs = @($processors).NumberOfProcessors
	}
}


Write-Host "Installing prerequisites for $($remoteComputers.Count) devices..." -ForegroundColor Green

# Let's install pre-reqs here.

$WUModuleInstallScript = Get-Content ".\InstallPreReqs.ps1" -Raw
$preReqScriptBlock = [scriptblock]::Create($WUModuleInstallScript)

foreach ($computer in $remoteComputers)
{
	$tcpConnect = Test-TCPport -ComputerName $computer.Name -TCPport "5985" -ErrorAction SilentlyContinue
	$tcpConnectSec = Test-TCPport -ComputerName $computer.Name -TCPport "5986" -ErrorAction SilentlyContinue
	
	if ($computer.Enabled -and ($tcpConnect -or $tcpConnectSec))
	{
		while (@(Get-Job | ?{ $_.State -eq "Running" }).Count -ge $NumJobs)
		{
			Write-Host "Waiting for open thread...($NumJobs Maximum)"
			Start-Sleep -Seconds 3
		}
		
		try
		{
			$session = New-PSSession -ComputerName $computer.Name -Credential $Credential -ErrorAction Stop
			Invoke-Command -Session $session -ScriptBlock $preReqScriptBlock -ArgumentList ($computer.Name) -AsJob -JobName $computer.Name -ErrorAction Stop
		}
		catch
		{
			"Unable to connect to $($computer.Name)."
		}
	}
	else
	{
		Write-Host "Unable to connect to $($computer.Name)." -ForegroundColor Red
	}
	
}

# Now we wait until all jobs have completed.

Write-Host "Installing prerequisites using a maximum of $NumJobs simultaneous threads..." -NoNewline -ForegroundColor Green

do
{
	Write-Host "." -NoNewline -ForegroundColor Green
	Start-Sleep -Milliseconds 500
}
while ((Get-Job -State Running).Count -gt 0)

Write-Host "." -ForegroundColor DarkGreen
Write-Host "All jobs completed!" -ForegroundColor Magenta

$jobs = (Get-Job)

$preReqScriptResult = @()

ForEach ($job in $jobs)
{
	$preReqScriptResult += Get-Job -Name $job.Name | Receive-Job
	Get-Job -Name $job.Name | Remove-Job
}

Get-PSSession | ft

Get-PSSession | Remove-PSSession

$preReqScriptResult | Out-File -PSPath "Prereqinstall.log"

Write-Host "All available updates have been retrieved!" -ForegroundColor Green


if ($ListAvailableUpdates)
{
	# Let's set up the Runspaces.
	
	<#$Throttle = $NumJobs
	$ListSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
	$ListRunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Throttle, $ListSessionState, $Host)
	$ListRunspacePool.Open()
	$ListJobs = @()/#>
	
	$getWUContent = Get-Content ".\Get-MicrosoftUpdate.ps1" -Raw
	$newScriptContent = $getWUContent + "`nGet-MicrosoftUpdate"
	$ListScriptBlock = [scriptblock]::Create($newScriptContent)
	
	Write-Host "Processing jobs for $($remoteComputers.Count) devices..." -ForegroundColor Green
	foreach ($computer in $remoteComputers)
	{
		$tcpConnect = Test-TCPport -ComputerName $computer.Name -TCPport "5985" -ErrorAction SilentlyContinue
		$tcpConnectSec = Test-TCPport -ComputerName $computer.Name -TCPport "5986" -ErrorAction SilentlyContinue
		
		if ($computer.Enabled -and ($tcpConnect -or $tcpConnectSec))
		{	
			while (@(Get-Job | ?{ $_.State -eq "Running" }).Count -ge $NumJobs)
			{
				Write-Host "Waiting for open thread...($NumJobs Maximum)"
				Start-Sleep -Seconds 3
				foreach ($job in (Get-Job | ?{ $_.State -eq "Running" }))
				{
					$beginTime = (Get-Job $job.Name | select *).PSBeginTime
					if ($beginTime.AddMinutes(5) -lt (Get-Date))
					{
						Get-Job $job.Name | Stop-Job
					}
				}
			}
			
			try
			{
				$session = New-PSSession -ComputerName $computer.Name -Credential $Credential -ErrorAction Stop
				Invoke-Command -Session $session -ScriptBlock $ListScriptBlock -AsJob -JobName $computer.Name -ErrorAction Stop
			}
			catch
			{
				"Unable to connect to $($computer.Name)."
			}
		}
		else
		{
			Write-Host "Unable to connect to $($computer.Name)." -ForegroundColor Red
		}
		
	}
	
	# Now we wait until all jobs have completed.
	
	Write-Host "Gathering available updates using a maximum of $NumJobs simultaneous threads..." -NoNewline -ForegroundColor Green
	
	do
	{
		Write-Host "." -NoNewline -ForegroundColor Green
		Start-Sleep -Milliseconds 500
	}
	while ((Get-Job -State Running).Count -gt 0)
	
	Write-Host "." -ForegroundColor DarkGreen
	Write-Host "All jobs completed!" -ForegroundColor Magenta
	
	$jobs = (Get-Job)
	
	$ListScriptResult = @()
	
	ForEach ($job in $jobs)
	{
		$data = Get-Job -Name $job.Name | Receive-Job
		Get-Job -Name $job.Name | Remove-Job		
		
		#$data = $data | ?{ $_.KB -ne $null } # | select ComputerName, KB, Date, Size, Description
		$data = $data | select PSComputerName, MsrcSeverity, Title, SupportUrl
		$ListScriptResult += $data
		
		Clear-Variable data
	}
	
	Get-PSSession | Remove-PSSession
	
	Write-Host "All available updates have been retrieved!" -ForegroundColor Green
	
	try
	{
		$listWUWorksheet = "WUWorksheet"
		$listWUTable = "WUTable"
		
		$excel = $ListScriptResult | Export-Excel -Path $listUpdatesXLSX -AutoSize -WorksheetName $listWUWorksheet -FreezeTopRow -TableName $listWUTable -PassThru
		$excel.Save(); $excel.Dispose()
	}
	catch
	{
		"Unable to create spreadsheet."
	}
}

if ($InstallUpdates)
{
	# Let's get the content of the Install-WindowsUpdates.ps1 file as raw data.
	$runWUContent = Get-Content ".\Install-WindowsUpdates.ps1" -Raw
	# Now we're going to add the actual function call to the previous imported content.
	$newScriptContent = $runWUContent + "`nInstall-WindowsUpdates"
	# Now, we'll convert it to a script block object to pass to the Invoke-Command cmdlet.
	$runWUScriptBlock = [scriptblock]::Create($newScriptContent)
	
	# We're aldo going to get the content of the XML file in the root directory to create the scheduled task.
	$scheduledTaskXML = Get-Content ".\PSWindowsUpdate.xml" -Raw
	
	Write-Host "Creating remote scheduled tasks..." -ForegroundColor Yellow
	
	$connectedComputers = @()
	
	# We'll start by creating the remote script and XML files and scheduled tasks on each computer.
	
	foreach ($computer in $remoteComputers)
	{
		# Use the custom Test-TCPport function to verify we can connect to the computer.
		
		$tcpConnect = Test-TCPport -ComputerName $computer.Name -TCPport "5985"
		$tcpConnectSec = Test-TCPport -ComputerName $computer.Name -TCPport "5986"
		
		if ($computer.Enabled -and ($tcpConnect -or $tcpConnectSec))
		{
			while (@(Get-Job | ?{ $_.State -eq "Running" }).Count -ge $NumJobs)
			{
				Write-Host "Waiting for open thread...($NumJobs Maximum)"
				Start-Sleep -Seconds 3
			}
			
			try
			{
				# For each computer, create a session to the remote computer, create the .ps1 files needed,
				# then create the scheduled task.
				
				$session = New-PSSession -ComputerName $computer.Name -Credential $Credential -ErrorAction Stop
				Invoke-Command -Session $session -AsJob -JobName $computer.Name -ErrorAction Stop -ScriptBlock {
					
					$scriptCommand = $using:runWUScriptBlock
					$XMLContent = $using:scheduledTaskXML
					
					if (-not (Get-ChildItem -Path "C:\Scripts" -ErrorAction SilentlyContinue))
					{
						try { New-Item -ItemType Directory -Name "Scripts" -Path "C:\" -ErrorAction Stop }
						catch { "Directory C:\Scripts already exists or access is denied." }
					}
					
					$scriptCommand | Out-File "C:\Scripts\Install-WindowsUpdates.ps1" -Force
					$XMLContent | Out-File "C:\Scripts\PSWindowsUpdate.xml" -Force
					Set-Location -Path "C:\Scripts"
					Register-ScheduledTask -Xml (Get-Content ".\PSWindowsUpdate.xml" | Out-String) -TaskName "PSWindowsUpdate" -Force
					Get-ScheduledTask -TaskName "PSWindowsUpdate"
				}
				
				$connectedComputers += $computer.Name
				
			}
			catch
			{
				"Unable to create scheduled task on $($computer.Name)."
			}
		}
		else
		{
			Write-Host "Unable to connect to $($computer.Name)." -ForegroundColor Red
		}
	}
	
	$jobs = (Get-Job)
	
	Write-Host "Waiting for outstanding jobs..." -NoNewline -ForegroundColor DarkGreen
	do
	{
		Write-Host "." -NoNewline -ForegroundColor DarkGreen
		Start-Sleep -Milliseconds 500
	}
	while ((Get-Job -State Running).Count -gt 0)
	
	Write-Host "." -ForegroundColor DarkGreen
	Write-Host "All jobs completed!" -ForegroundColor Magenta
	
	Write-Host "Starting for Windows Update tasks..." -ForegroundColor DarkYellow
	
	# Now that the files and jobs have been created, let's execute the update task.
	
	foreach ($computer in $connectedComputers)
	{
		$taskCimSession = New-CimSession -ComputerName $computer -Credential $Credential
		$task = Get-ScheduledTask -TaskName PSWindowsUpdate -Session $taskCimsession -ErrorAction SilentlyContinue
		
		if ($task)
		{
			try
			{
				$task | Start-ScheduledTask -ErrorAction Stop
				Write-Host "Task (PSWindowsUpdate) on device $computer has been started."
			}
			catch
			{
				"Unable to start scheduled task (PSWindowsUpdate) on device $($computer.Name)."
			}
			$taskCimSession | Remove-CimSession
			Clear-Variable taskCimSession
		}
		else
		{
			Write-Host "Task (PSWindowsUpdate) does not exist on device $($computer.Name)." -ForegroundColor Red
		}
		
	}
	
	# Now, we wait for the update tasks to complete.
	
	Write-Host "Waiting for Windows Update tasks..." -NoNewline -ForegroundColor Green
	
	do
	{
		foreach ($job in $jobs)
		{
			$runningCimSession = New-CimSession -ComputerName $job.Name -Credential $Credential
			$taskStatus = Get-ScheduledTask -TaskName PSWindowsUpdate -Session $runningCimsession -ErrorAction SilentlyContinue
			
			if ($taskStatus.State -eq "Ready")
			{
				try
				{
					Get-Job -Name $job.Name | Remove-Job -ErrorAction Stop
				}
				catch
				{
					"Unable to remove job $($job.Name)."
				}
				
				try
				{
					Unregister-ScheduledTask -Session $runningCimSession -TaskName PSWindowsUpdate -Confirm:$false -ErrorAction Stop
				}
				catch
				{
					"Unable to remove scheduled task (PSWindowsUpdate)."
				}
			}
			
			Get-CimSession | Remove-CimSession
			Clear-Variable runningCimSession
			Clear-Variable taskStatus
		}
		Write-Host "." -NoNewline -ForegroundColor Green
		Start-Sleep -Seconds 5
	}
	while ((Get-Job).Count -gt 0)
	
	Write-Host "." -ForegroundColor Green
	Write-Host "All updates installed! See below for pending reboot information!" -ForegroundColor Magenta
	
	Get-PSSession | Remove-PSSession
	
	# All update jobs have been completed, but let's determine which machines require a reboot.
	
	$compsNeedingReboot = @()
	
	foreach ($computer in $connectedComputers)
	{
		$pendingReboot = Test-PendingReboot -ComputerName $computer -Credential $Credential -SkipConfigurationManagerClientCheck
		
		if ($pendingReboot.IsRebootPending)
		{
			$compsNeedingReboot += $computer
			
			#Restart-Computer -ComputerName $computer -Wait -For PowerShell -Timeout 600 -Delay 2 -Force -ErrorAction Stop -AsJob
		}
	}
	
	#Send the reboot command to all machines in batches. The number of simultaneous reboot commands is equal to the
	# value of the $NumJobs parameter.
	
	if ($compsNeedingReboot.Count -gt 0)
	{
		foreach ($comp in $compsNeedingReboot) { Write-Host "Rebooting machine $comp." -ForegroundColor Green}
		
		$compsNeedingReboot | Invoke-Parallel { Restart-Computer -ComputerName $_ -Wait -For PowerShell -Timeout 600 -Delay 2 -Force -ErrorAction Stop } -ThrottleLimit $NumJobs
	}
	
}

if ($GetUpdateHistory)
{
	# $WUHistoryScriptBlock = { Import-Module PSWindowsUpdate; Get-WUHistory }
	$WUHistoryScriptBlock = { Get-Hotfix }
	
	#$newScriptContent
	
	Write-Host "Gathering Windows Update history for $($remoteComputers.Count) devices..." -ForegroundColor Green
	
	foreach ($computer in $remoteComputers)
	{
		$tcpConnect = Test-TCPport -ComputerName $computer.Name -TCPport "5985"
		$tcpConnectSec = Test-TCPport -ComputerName $computer.Name -TCPport "5986"
		
		if ($computer.Enabled -and ($tcpConnect -or $tcpConnectSec))
		{
			while (@(Get-Job | ?{ $_.State -eq "Running" }).Count -ge $NumJobs)
			{
				Write-Host "Waiting for open thread...($NumJobs Maximum)"
				Start-Sleep -Seconds 3
			}
			
			try
			{
				$session = New-PSSession -ComputerName $computer.Name -Credential $Credential -ErrorAction Stop
				Invoke-Command -Session $session -ScriptBlock { Set-ExecutionPolicy Unrestricted }
				Invoke-Command -Session $session -ScriptBlock $WUHistoryScriptBlock -AsJob -JobName $computer.Name
			}
			catch
			{
				"Unable to connect to $($computer.Name)."
			}
		}
		else
		{
			Write-Host "Unable to connect to $($computer.Name)." -ForegroundColor Red
		}
		
	}
	
	$jobs = (Get-Job)
	
	Write-Host "Waiting for outstanding jobs..." -NoNewline -ForegroundColor DarkGreen
	do
	{
		Write-Host "." -NoNewline -ForegroundColor DarkGreen
		Start-Sleep -Milliseconds 500
	}
	while ((Get-Job -State Running).Count -gt 0)
	
	Write-Host "." -ForegroundColor DarkGreen
	Write-Host "All jobs completed!" -ForegroundColor Magenta
	
	$WUHistory = @()
	
	foreach ($job in $jobs)
	{
		$data = Get-Job -Name $job.Name | Receive-Job
		Remove-Job $job
		
		
		$data = $data | ?{ $_.HotFixID -ne $null } | select PSComputerName, HotFixID, InstalledBy, InstalledOn, Description, Caption
		
		$WUHistory += $data
		
		Clear-Variable data
	}
	
	Get-PSSession | Remove-PSSession
	
	try
	{
		$WUHistoryWorksheet = "WUHistoryWorksheet"
		$WUHistoryTable = "WUHistoryTable"
		
		$excel = $WUHistory | Export-Excel -Path $updateHistorytXLSX -AutoSize -WorksheetName $WUHistoryWorksheet -FreezeTopRow -TableName $WUHistoryTable -PassThru
		$excel.Save(); $excel.Dispose()
	}
	catch
	{
		"Unable to create spreadsheet."
	}
	
}