# Get-FTPData.ps1 -Account <accountname>
# PowerShell Script to receive data via FTP or FTPS with external configuration file.

# Scriptversion			: 1 (2017-02-04) (alpha)
# Requirements			: PowerShell | WinSCP (WinSCPnet.dll,winscp.exe) | Windows Server 2008 R2 and newer
# Author			: Chris Kragt <ck@kragt.pro>
# Documentation			: currently not available

Function Global:Get-FTPData
{
	Param
    (
		[Parameter(Position=0 , Mandatory = $true, ValueFromPipeline = $true)]
		$Account,
 
		[Parameter(Position=1 , Mandatory = $false, ValueFromPipeline = $true)]
		$PreRun,

		[Parameter(Position=2 , Mandatory = $false, ValueFromPipeline = $true)]
		$PostRun
	)

	Process
	{
		If ($PreRun) 
		{
			$PreRun
		}

		[xml]$AccountData        = Get-Content "c:\powershell\run\Get-FTPData\ftp_accounts.xml"
		[string]$FTPServer       = $AccountData.accounts.$Account.target.server;
		[string]$FTPSecurity     = $AccountData.accounts.$Account.target.security;
		[string]$FTPPosttransfer = $AccountData.accounts.$Account.target.posttransfer;
		[string]$FTPUsername     = $AccountData.accounts.$Account.target.username;
		[string]$FTPPassword     = $AccountData.accounts.$Account.target.password; 
		[string]$RemoteFolder    = $AccountData.accounts.$Account.environment.remotefolder;
		[string]$LocalFolder     = $AccountData.accounts.$Account.environment.localfolder;
		[string]$PrerunAbort     = $AccountData.accounts.$Account.settings.prerunabort;
		[string]$RunUntill       = $AccountData.accounts.$Account.settings.rununtill;

		If ($FTPSecurity -ne "Ftp" -or $FTPSecurity -ne "Ftps") 
		{
			Write-Host "Aborting due to misconfiguration:" -ForegroundColor Red
			Write-Host ("FTP Security must be either Ftp or Ftps. Found {0} in Accounttree {1}." -f $FTPSecurity, $Account)
			Start-Sleep 5; exit 0
		}

		If ($FTPPosttransfer -ne "delete" -or $FTPPosttransfer -ne "keep") 
		{
			Write-Host "Aborting due to misconfiguration:"  -ForegroundColor Red
			Write-Host ("FTP Post Transfer must be either delete or keep. Found {0} in Accounttree {1}." -f $FTPPosttransfer, $Account)
			Start-Sleep 5; exit 0
		}

		If ($RunUntill) 
		{
			$Date = Get-Date; $CurrentDay = $Date.Day;

			If ($CurrentDay -ge $RunUntill) 
			{
				#fix error: not getting xml value rununtill.
				Write-Host "Aborting due to configuration setting:" -ForegroundColor Red
				Write-Host ("This script will only run to day of month {0}. Today is day number {1}." -f $RunUntill, $CurrentDay)
				Start-Sleep 5; exit 0
			}
		}

		If ($PrerunAbort -eq "true") 
		{
			Write-Host "Aborting due to configuration setting:" -ForegroundColor Red
			Write-Host ("Prerun abort set in accounttree {0} to true." -f $Account)
			Start-Sleep 5; exit 0
		}

		[Reflection.Assembly]::LoadFrom("c:\powershell\run\Get-FTPData\WinSCPnet.dll") | Out-Null

		If ($FTPSecurity -eq 'Ftps') 
		{
			$SessionOptions = New-Object WinSCP.SessionOptions -Property @{
				Protocol  = [WinSCP.Protocol]::Ftp
				FtpSecure = [WinSCP.FtpSecure]::explicittls
				HostName  = $FTPServer
				UserName  = $FTPUsername
				Password  = $FTPPassword
			}
		}
		Else 
		{
			$SessionOptions = New-Object WinSCP.SessionOptions -Property @{
				Protocol = [WinSCP.Protocol]::Ftp
				HostName = $FTPServer
				UserName = $FTPUsername
				Password = $FTPPassword
			}
		}

		$Session = New-Object WinSCP.Session

		$Session.Open($SessionOptions)

		# Synchronize files to local directory, collect results
		$synchronizationResult = $Session.SynchronizeDirectories([WinSCP.SynchronizationMode]::Local, $LocalFolder, $RemoteFolder, $False)
 
		# Iterate over every download
		foreach ($download in $synchronizationResult.Downloads)
		{
			# Success or error?
			if ($download.Error -eq $Null)
			{
				if ($FTPPosttransfer -eq "delete") {
					Write-Host ("Download of {0} succeeded, removing from source" -f $download.FileName)
					# Download succeeded, remove file from source
					$removalResult = $Session.RemoveFiles($Session.EscapeFileMask($download.FileName))
 
					if ($removalResult.IsSuccess)
					{
					Write-Host ("Removing of file {0} succeeded" -f $download.FileName)
						$AccountData.accounts.$Account.internals.prerunabort = "true"
					}
					else
					{
						Write-Host ("Removing of file {0} failed" -f $download.FileName)
					}
				}
				else {
					Write-Host ("Download of {0} succeeded, source files are not deleted." -f $download.FileName)
				}

			}

			else
			{
				Write-Host ("Download of {0} failed: {1}" -f $download.FileName, $download.Error.Message)
			}
		}

		$Session.Dispose()

		If ($PostRun) 
		{
			$PostRun
		}
	}
}
