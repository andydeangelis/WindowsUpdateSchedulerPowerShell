<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
	 Created on:   	4/15/2019 12:21 PM
	 Created by:   	andy-user
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

function Install-WindowsUpdates
{
	# Create a new COM object for an update session.
	$IUpdateSession = New-Object -ComObject Microsoft.Update.Session
	# Instantiate the update searcher method in the COM object.
	$IUpdateSearcher = $IUpdateSession.CreateUpdateSearcher()
	
	# Perform the initial search for available updates.
	$searchResult = $IUpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
	$updateCount = $searchResult.Updates.Count
	$availUpdates = $searchResult.Updates
	
	if ($updateCount -gt 0)
	{
		$updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
		$updatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
		
		foreach ($update in $availUpdates)
		{
			# Let's check to see if the updates have already been downloaded.
						
			if (-not ($update.IsDownloaded))			
			{
				# If updates have not been downloaded, we'll add them to a different collection and download them.
				$updatesToDownload.Add($update)
			}
		}
		
		if ($updatesToDownload.Count -gt 0)
		{
			try
			{
				$updateDownloader = $IUpdateSession.CreateUpdateDownloader()
				$updateDownloader.Updates = $updatesToDownload
				$updateDownloader.Download()
			}
			catch
			{
				"Unable to download updates."
			}
		}
		
		# Now that all the updates have been downloaded, let's get to installing.
		
		# We'll start by refreshing the searchResult, updateCount and availUpdates variables.
		
		Clear-Variable searchResult
		Clear-Variable updateCount
		Clear-Variable availUpdates
		
		$searchResult = $IUpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
		$updateCount = $searchResult.Updates.Count
		$availUpdates = $searchResult.Updates
		
		foreach ($update in $availUpdates)
		{
			if ($update.IsDownloaded)
			{
				$update.AcceptEula()
				$updatesToInstall.Add($update)				
			}
		}
		
		# Now that we have all our updates in a collection, let's install them.
		
		If ($updatesToInstall.Count -gt 0)
		{
			$updateInstaller = $IUpdateSession.CreateUpdateInstaller()
			$updateInstaller.Updates = $updatesToInstall
			#$updateInstaller | get-member
			$installResult = $updateInstaller.Install()
		}
	}
	
	# Determine if a reboot is required.
	
	# $rebootRequired = $updateInstallationResult.RebootRequired
	
	# return $rebootRequired
}