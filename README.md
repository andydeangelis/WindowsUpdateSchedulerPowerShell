# WindowsUpdateSchedulerPowerShell

This script suite is composed of three parts that allow you to fine tune scheduling of Windows update.

This script can be run from a single node (for example, a WSUS server, a domain controller, a Windows 10 workstation, etc.) to remotely control and update every other node.

Parameters:

	- The main wrapper file is called WindowsUpdateMainScript.ps1. It is the main file that multi-threads the update process and generates the Excel reports. It takes the following parameters:
		-SearchBase: Mandatory string parameter. This is the distinguished name of the OU in Active Directory housing the target machines.
		-Credential: Mandatory PSCredential object. The PSCredential object used to run the jobs. Must be an admin on both the controller server and target machines.
		-NumJobs: Optional positive integer value. Optional parameter to set number of simultaneous jobs. If not set, the script will determine the number of simultaneous jobs based on the number of logical cores. This option is ignored if -ListAvailableUpdates or -InstallUpdates switch is missing.
		-InstallUpdates: Optional switch parameter. Tells the script to install the updates it finds available. Generates an Excel report upon completion. 
		-ListAvailableUpdates: Optional switch parameter. Lists the available updates to install. Generates an Excel report upon completion. 
		-GetUpdateHistory: Optional switch parameter. Tells the script to install the updates it finds available. Generates an Excel report upon completion. 
	
Pre-reqs:

	- This has only been tested with PoSH 4 and higher. PoSH v3.0 "should" work, but I haven't tested it. With that said, don't expect it to run on Server 2003/2008/XP/etc.
	- Target servers (including the controller server running this script) should have internet access, specifically to the Microsoft PSGallery. If internet access from these nodes exists, the appropriate modules will be installed automatically.
		- If internet access is not possible from the controller node, you will need to manually install the ImportExcel, PSWindowsUpdate, PendingReboot and PSParallel modules from the PSGallery.
		- The PSWindowsUpdate and Pending reboot modules need to be installed manually on all WU client nodes if they do not have internet access.
	- PS Remoting ports need to be open from the controller server to all target servers in order to pass the Invoke-Command cmdlet.
	- If running from the main launcher script, the account specified in the PSCredential object must have the ability to read from AD (no local accounts).
	- The PSCredential object passed to the main script must have admin rights on the target servers.
	- Ideally, your target machines should be configured to "Download and Notify" or "Notify before Download." Automatic installation renders this script useless, and disabling of automatic updates will break the script. These settings can be accomplished via GPO. If using a WSUS server, only approved updates will be listed.
	- The script ignores Disabled computer accounts, so it doesn't try to connect to them.
	- The script does test connectivity to computer accounts, so if a computer account is enabled but not responding (i.e. powered off), it will also be ignored.
	
Files:

	- The file structure must remain as is in order to run the script. 
	- There are three function definition files. Note that the functions do not create reports; they only output to the screen.
		- Get-MicrosoftUpdate.ps1 creates the Get-MicrosoftUpdate function. This can be sourced and used external to the main script. Once sourced, it can be run to list the available updates of the local machine only.
			Example 1:
			
				PS C:\Scripts\Projects\WindowsUpdate> Get-MicrosoftUpdate | fl

				Title        : Definition Update for Windows Defender Antivirus - KB2267602 (Definition 1.291.2208.0)
				MsrcSeverity :
				SupportUrl   : https://go.microsoft.com/fwlink/?LinkId=52661
				
		- Install-WindowsUpdates.ps1 creates the Install-WindowsUpdates function. This function will download and install any available Windows Updates, but it does not perform a reboot (even if the machine requires it). This can be sourced and used external to the main script. Once sourced, it can be run to list the available updates of the local machine only.
			Example 1:
			
				PS> Install-WindowsUpdate
		
		- Test-TCPPort.ps1 tests a TCP socket connect to a specified TCP port. While not as robust as the built-in Test-NetConnection function, it returns much faster if the connection fails (important for scripts). It uses the System.Net.Sockets.TcpClient class to create the connection.
			Example 1:
			
				PS C:\Scripts\Projects\WindowsUpdate> Test-TCPport -ComputerName andy-2k16-vmm2 -TCPport 5985

				hostname         port open
				--------         ---- ----
				{andy-2k16-vmm2} 5985 True
				
	- InstallPreReqs.ps1 is called to install the prerequisite components on any remote machines. It is a straight ps1 script, not a sourced function.
	- PSWindowsUpdate.xml is the configuration file for the Install-WindowsUpdate scheduled task. Windows Update COM components cannot be called remotelty, so we will use this to create a scheduled task local to each target machine, and then we will initiate the task to install the Windows Updates.
		
Main Script Usage (Manual Run)

	- Usage of the script is pretty easy. Simply pass the DN, the credential object and the task to run it.
		Example 1 (running manually, get a list of available updates):
		
			PS> .\WindowsUpdateLauncher.ps1 -SearchBase "ou=servers,dc=testdomain,dc=local" -Credential (Get-Credential) -ListAvailableUpdates -NumJobs 4
				- The above command will list all available updates to all enabled computer accounts in the Servers OU, and it will export the list of available updates into an Excel spreadsheet.
				
		Example 2 (running manually, install all available updates):
		
			PS> .\WindowsUpdateLauncher.ps1 -SearchBase "ou=Servers,dc=testdomain,dc=local" -Credential (Get-Credential) -InstallUpdates -NumJobs 4
				- The above command will install all available updates to all enabled computer accounts in the Servers OU, and it will export the list of updates into an Excel spreadsheet.
				
		Example 3 (running manually, generate currently installed updates report):
		
			PS> .\WindowsUpdateLauncher.ps1 -SearchBase "ou=Servers,dc=testdomain,dc=local" -Credential (Get-Credential) -GetUpdateHistory -NumJobs 4
				- The above command will install all available updates to all enabled computer accounts in the Servers OU, and it will export the list of updates into an Excel spreadsheet.
				
		Example 4 (running manually, install only pre-reqs):
		
			PS> .\WindowsUpdateLauncher.ps1 -SearchBase "ou=Servers,dc=testdomain,dc=local" -Credential (Get-Credential) -GetUpdateHistory -NumJobs 4
				- The above command will install all available updates to all enabled computer accounts in the Servers OU, and it will export the list of updates into an Excel spreadsheet.
				
Main Script Usage (Scheduled Task)

	- You can also configure the script to run on a schedule (say, the third Wednesday of every month) by exporting credentials to a secured XML file for later use. There are some caveats regarding this:
		- You generate the credential file by using the Export-Clixml cmdlet. Be sure to note where you place the resulting XML file.
		
			PS> (Get-Credential) | Export-Clixml WinUpdateDomainCreds.xml
			
		- Once created, you can use the XML file instead of haviung to enter credentials manually each time.
		
			PS> .\WindowsUpdateLauncher.ps1 -SearchBase "ou=Domain Controllers,dc=testdomain,dc=local" -Credential (Import-Clixml WinUpdateDomainCreds.xml) -InstallUpdates -NumJobs 8
				- Note: The user account running the command from the command prompt must be the same user account that generated the XML file. XML files created using the Export-Clixml use the Windows Data Protection API, and as such, they are tied to the user account that creates them. No other user accounts (including Domain/Enterprise Admins) can decrypt these files.
				
		- To create the scheduled task, perform the following steps:
		
			1. Log in to the controller server with the same account that will run the scheduled task. I recommend creating a service account in AD. For our example, we'll create a user called DOMAIN\UpdateAccount.
			
			2. While logged in to the controller as DOMAIN\UpdateAccount, open a PowerShell window and run:
			
				PS> (Get-Credential) | ExportClixml WinUpdateDomainCreds.xml
				
			3. Now, create a scheduled task that will run one of the options above. Set the sceduled task parameters (triggers, logon account, etc.). I won't go into all the details of creating a Scheduled Task here, but it's pretty simple. The contect/command should be similar to the following:
			
				Scheduled command to generate a report listing available updates, with a maximum of 4 simultaneous jobs:
			
					PS> .\WindowsUpdateLauncher.ps1 -SearchBase "ou=Domain Controllers,dc=testdomain,dc=local" -Credential (Import-Clixml WinUpdateDomainCreds.xml) -ListAvailableUpdates -NumJobs 4
					
				Scheduled command to install all available updates on target machines, with a maximum of 4 simultaneous jobs:
				
					PS> .\WindowsUpdateLauncher.ps1 -SearchBase "ou=Domain Controllers,dc=testdomain,dc=local" -Credential (Import-Clixml WinUpdateDomainCreds.xml) -InstallUpdates -NumJobs 4
					
				Scheduled command to install pre-reqs only on all target machines:
				
					PS> .\WindowsUpdateLauncher.ps1 -SearchBase "ou=Domain Controllers,dc=testdomain,dc=local" -Credential (Import-Clixml WinUpdateDomainCreds.xml)
		