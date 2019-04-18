<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
	 Created on:   	4/15/2019 10:56 AM
	 Created by:   	andy-user
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

#region Get-MicrosoftUpdates
function Get-MicrosoftUpdate
{
	$IUpdateSession = New-Object -ComObject Microsoft.Update.Session
	$IUpdateSearcher = $IUpdateSession.CreateUpdateSearcher()
	
	$searchResult = $IUpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
	$updateCount = $searchResult.Updates.Count
	
	$availUpdates = $searchResult.Updates | select Title, MsrcSeverity, SupportURL
		
	return $availUpdates
}
#endregion