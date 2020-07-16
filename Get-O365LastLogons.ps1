# ***************************************************************************************************
# ***************************************************************************************************
#
#  Author       : Cary GARVIN
#  Contact      : cary(at)garvin.tech
#  LinkedIn     : https://www.linkedin.com/in/cary-garvin
#  GitHub       : https://github.com/carygarvin/
#
#
#  Script Name  : Get-O365LastLogons.ps1
#  Version      : 1.0
#  Release date : 05/05/2020 (CET)
#  History      : The present script has been developped for Organizations to have an audit view on last actions by users and guests in their Office 365 tenant.
#  Purpose      : The present Script generates a list of Office365 last logons (in fact activities) along with basic information such as WorkLoad (O365 product), Client IP Address and so on.
#                 The Script will output 2 CSV files, one with Last Logons for Office365 Users (differenciated on 'UserType' property) and another one for Office 365 Guests.
#
#
#  This script is to be launched within "Exchange Online PowerShell" in order to invoke the cmdlet 'Search-UnifiedAuditLog' around which the present Script is built.
#  Running it from a PS-Session is not advised in case the User Management Admin account used to run it is subject to MFA.
#  Supply your O365 User Management Admin credentials to the MsolService when prompted.
#  Auditing for the O365 tenant should be enabled otherwise Unified Audit Logs could potentially not contain any worthwhile information...
#
#




####################################################################################################
#                    User configurable parameters (default below values are best)                  #
####################################################################################################

$LogLookbackDays = 90                                                          # 90 days prior date (The Max available in Office365).
$PauseDelayBetweenQueryBatches = 90                                            # Pause delay for when an error is encountered due to either Office365 throttling (soft error) or any other more serious reason




####################################################################################################
#                             Global script Constants and variables                                #
####################################################################################################

$script:ScriptPath = split-path -parent $MyInvocation.MyCommand.Definition     # Extraction of current script path
$script:ScriptName = (Get-Item $MyInvocation.MyCommand).Basename               # Extraction of current script base name (i.e. without file extension)
$script:MyDocsFolder = [Environment]::GetFolderPath("MyDocuments")             # 'My Documents' folder of user currently running the script (used for optional History files)
$script:ExecutionTimeStamp = get-date -format "yyyy-MM-dd_HH-mm-ss"            # Script launch timestamp used to stamp generated output files

$startDate = "{0:yyyy-MM-dd}" -f (get-date).AddDays(-$LogLookbackDays)         # Starting from $LogLookbackDays days ago
$endDate = "{0:yyyy-MM-dd}" -f (get-date)                                      # Ending at current date.




####################################################################################################
#                Handling of command line swtich and other script start info                       #
####################################################################################################

If ($args -ne $null)
	{
	write-host "Number of arguments : >$($args.length)<"
	If (($args.length -gt 1) -or ($args[0] -ne "-InputList"))
		{
		write-host "The present script supports either no, or only one command line switch."
		write-host "The supported switch is '-InputList' to provide to the Script a list of users for which the Last Logons information needs to be collected."
		write-host "The syntax is '.\Get-O365LastLogons.ps1 -InputList UserList.txt' where UserList.txt is the file containing the list of Users for which the Office365 Last Logons information needs to be queried."
		write-host "If no command line switch is passed when invoking the script, it will output current situation files in the directory where it resides with the information gathered from Office365 for the last $LogLookbackDays days."
		Break
		}
	$ScriptExecutionMode = $args[0]
	If ($ScriptExecutionMode -eq "-InputList") {$InputFile = $args[1]}
	}

write-host "Please make sure this script is run from the 'Exchange Online PowerShell' and not from a PSSession to 'https://outlook.office365.com/powershell-liveid/ !'" -foregroundcolor "yellow"
write-host "You will be prompted to supply a User Management Admin account in order to connect to MSOLService.`r`n`r`n" -foregroundcolor "yellow"
start-sleep -s 5

Connect-EXOPSSession
Connect-MsolService




####################################################################################################
#                                       Script Functions                                           #
####################################################################################################

Function Get-AllLastLoginInfo
	{
	param
		(
		[Parameter(Mandatory = $true)]
		[string]$MSOLObjectType,
		[Parameter(Mandatory = $true)]
		[array]$RemainingQueries
		)

	$NumberOfQueries = 0
	$QueryPassNr = 1
	$AllMSOLObjectsLastLogin = @()
	$TotalObjectsToQuery = $RemainingQueries.length
	write-host "Processing $MSOLObjectType objects" -foregroundcolor "yellow"
	While ($RemainingQueries.length -gt 0)
		{
		$FailedQueries = @()
		write-host "Processing Query Pass Nr $QueryPassNr : $($RemainingQueries.length) $MSOLObjectType to process..." -foregroundcolor "yellow"
		ForEach ($MSOLObject in $RemainingQueries)
			{
			If (!($Error)) {$NumberOfErrorsBefore = 0}
			Else {$NumberOfErrorsBefore = $Error.count}
			If ($NumberOfErrorsBefore -eq 255)
				{
				$error | out-file -Encoding utf8 "$($script:ScriptPath)\$($script:ScriptName)_$($script:ExecutionTimeStamp)_errors.log" -Append
				$error.clear();	$NumberOfErrorsBefore = 0
				}
			$NumberOfQueries++
			Try
				{
				$lastLoginInfo = Search-UnifiedAuditLog -UserIds $MSOLObject.UserPrincipalName -StartDate $startDate -EndDate $endDate -OutVariable +auditData -ErrorAction Stop | Foreach-Object {$_.CreationDate = [DateTime]$_.CreationDate; $_} | Sort-Object CreationDate | Select-Object -Last 1 | Select CreationDate, UserIds, RecordType, AuditData
				$LastLoginDate = $lastLoginInfo.CreationDate
				write-host "Query #$NumberOfQueries @ $(Get-Date) | $($MSOLObjectType -replace "".$""): $($MSOLObject.UserPrincipalName) | Last Login: $LastLoginDate" -foregroundcolor "white"
				If ($lastLoginInfo) {$MSOLObjectLastLogin = $MSOLObject | Select-Object @{l="LastLoginDate"; e={$LastLoginDate}}, UserPrincipalName, DisplayName, @{l="RecordType"; e={$lastLoginInfo.RecordType}}, @{l="ClientIP";e={(ConvertFrom-Json $lastLoginInfo.AuditData).ClientIp}}, @{l="WorkLoad";e={(ConvertFrom-Json $lastLoginInfo.AuditData).WorkLoad}}, @{l="ObjectId";e={(ConvertFrom-Json $lastLoginInfo.AuditData).ObjectId}}, isLicensed, UserType}
				Else {$MSOLObjectLastLogin = $MSOLObject | Select-Object @{l="LastLoginDate"; e={"No logon in last $LogLookbackDays days"}}, UserPrincipalName, DisplayName, @{l="RecordType"; e={""}}, @{l="ClientIP";e={""}}, @{l="WorkLoad";e={""}}, @{l="ObjectId";e={""}}, isLicensed, UserType}
				$AllMSOLObjectsLastLogin += $MSOLObjectLastLogin
				}
			Catch
				{
				$Error[0]
				$FailedQueries += $MSOLObject
				write-host "Error encountered. Appending to failed $MSOLObjectType for next pass and pausing current pass for $PauseDelayBetweenQueryBatches seconds..." -foregroundcolor "cyan"
				start-sleep -s $PauseDelayBetweenQueryBatches
				}
			}
		If ($FailedQueries.length -eq 0) {write-host "Final pass $QueryPassNr finished with $($AllMSOLObjectsLastLogin.length)/$TotalObjectsToQuery) completed ($([math]::Round(($($AllMSOLObjectsLastLogin.length)*100)/$TotalObjectsToQuery,2))%). No more $MSOLObjectType to query for last logons..." -foregroundcolor "yellow"}
		Else {write-host "Pass $QueryPassNr finished with $($AllMSOLObjectsLastLogin.length)/$TotalObjectsToQuery completed ($([math]::Round(($($AllMSOLObjectsLastLogin.length)*100)/$TotalObjectsToQuery,2))%). Remaining $($FailedQueries.length) $MSOLObjectType to query for last logons..." -foregroundcolor "yellow"}

		$RemainingQueries = $FailedQueries
		$QueryPassNr++
		}
	return $AllMSOLObjectsLastLogin
	}




####################################################################################################
#                                          Script Main                                             #
####################################################################################################

# Script initialisation
$Error.Clear()
# Start Script Timing StopWatch and Transcript
$TotalScriptTimer = [system.diagnostics.stopwatch]::startNew()
Start-Transcript -Path "$($script:ScriptPath)\$($script:ScriptName)_$($script:ExecutionTimeStamp)_transcript.log" -NoClobber | out-null


# Alternative execution modes, i.e. from specified input file or for ALL MSOL Objects present in Azure AD.
$MSOLAccounts  = @()
If ($ScriptExecutionMode -eq "-InputList")
	{
	# Reading list of MSOL accounts listed in input file
	write-host "Reading list of MSOL objects from input file '$InputFile'..." -foregroundcolor "white"
	If (test-path "$($script:MyDocsFolder)\$InputFile") {$MSOLList = get-content "$($script:MyDocsFolder)\$InputFile"}
	Else {write-host "Input file '$($script:MyDocsFolder)\$InputFile' could not be found! Aborting script!" -foregroundcolor "red"; Break}
	write-host "Collecting MSOL info for $($MSOLList.length) objects listed in file '$($script:MyDocsFolder)\$InputFile'." -foregroundcolor "yellow"
	ForEach ($MSOLAccount in $MSOLList)
		{
		Try	{$MSOLAccounts += Get-MsolUser -UserPrincipalName $MSOLAccount -ErrorAction Stop | select UserPrincipalName, DisplayName, isLicensed, UserType}
		Catch
			{
			$NotFoundMSOLEntry = New-Object PSObject
			$NotFoundMSOLEntry | Add-Member NoteProperty -Name "UserPrincipalName" -Value $MSOLAccount
			$NotFoundMSOLEntry | Add-Member NoteProperty -Name "DisplayName" -Value "Not found in Azure"
			$NotFoundMSOLEntry | Add-Member NoteProperty -Name "isLicensed" -Value "Not found in Azure"
			$NotFoundMSOLEntry | Add-Member NoteProperty -Name "UserType" -Value "Not found in Azure"
			$MSOLAccounts += $NotFoundMSOLEntry
			}
		}
	write-host "The `$MSOLAccounts array contains $($MSOLAccounts.length) entries..." -foregroundcolor "red"
	}
Else
	{
	# Collecting ALL MSOL Objects from Azure AD
	write-host "Querying Azure AD for all MSOL objects..." -foregroundcolor "white"
	$MSOLAccounts = Get-MsolUser -All | select UserPrincipalName, DisplayName, isLicensed, UserType
	write-host "A total of $($MSOLAccounts.length) MSOL objects (USERS + GUESTS) have been collected." -foregroundcolor "yellow"
	}



# Processing of Last Logons for MSOL Users
$MSOLUsers = $MSOLAccounts | where {$_.UserType -eq "User"}
write-host "`r`n`r`nTotal MSOL Users : $($MSOLUsers.length)"
If ($MSOLUsers.length -gt 0)
	{
	$MSOLUsers | export-csv "$($script:MyDocsFolder)\$($script:ScriptName)_$($script:ExecutionTimeStamp)_O365_Users_List.csv" -NoTypeInformation
	$AllUsersLastLogin = Get-AllLastLoginInfo "users" $MSOLUsers
	write-host "All passes for USERS completed. Exporting results..." -foregroundcolor "yellow"
	$AllUsersLastLogin | sort -Property LastLoginDate -Descending | export-csv "$($script:MyDocsFolder)\$($script:ScriptName)_$($script:ExecutionTimeStamp)_O365_Users_LastLogins.csv" -NoTypeInformation
	}
Else {write-host "Skipping Last Logon queries for Users since there are none to process!" -foregroundcolor "magenta"}



# Processing of Last Logons for MSOL Guests
$MSOLGuests = $MSOLAccounts | where {$_.UserType -eq "Guest"}
write-host "`r`n`r`nTotal MSOL Guests : $($MSOLGuests.length)"
If ($MSOLGuests.length -gt 0)
	{
	$MSOLGuests | export-csv "$($script:MyDocsFolder)\$($script:ScriptName)_$($script:ExecutionTimeStamp)_O365_Guests_List.csv" -NoTypeInformation
	$AllGuestsLastLogin = Get-AllLastLoginInfo "guests" $MSOLGuests
	write-host "All passes for GUESTS completed. Exporting results..." -foregroundcolor "yellow"
	$AllGuestsLastLogin | sort -Property LastLoginDate -Descending | export-csv "$($script:MyDocsFolder)\$($script:ScriptName)_$($script:ExecutionTimeStamp)_O365_Guests_LastLogins.csv" -NoTypeInformation
	}
Else {write-host "Skipping Last Logon queries for Guests since there are none to process!" -foregroundcolor "magenta"}



# Processing of Last Logons for MSOL Objects from input file no longer or not present in Azure AD
If ($ScriptExecutionMode -eq "-InputList")
	{
	$MSOLNotFoundObjects = $MSOLAccounts | where {$_.UserType -eq "Not found in Azure"}
	write-host "There are >$($MSOLNotFoundObjects.length)< MSOL Objects from the input list not found in Azure AD!"
	If ($MSOLNotFoundObjects)
		{
		$MSOLNotFoundObjects | export-csv "$($script:MyDocsFolder)\$($script:ScriptName)_$($script:ExecutionTimeStamp)_O365_NotFound_List.csv" -NoTypeInformation
		$AllNotFoundLastLogin = Get-AllLastLoginInfo "NotFound" $MSOLNotFoundObjects
		write-host "All passes for 'NotFound' objects completed. Exporting results..." -foregroundcolor "yellow"
		$AllNotFoundLastLogin | sort -Property LastLoginDate -Descending | export-csv "$($script:MyDocsFolder)\$($script:ScriptName)_$($script:ExecutionTimeStamp)_O365_NotFound_LastLogins.csv" -NoTypeInformation
		}
	}



# Script closure and cleanup
$TotalScriptTimer.Stop() 
If ($error)
	{
	$error | out-file  -Encoding utf8 "$($script:MyDocsFolder)\$($script:ScriptName)_$($script:ExecutionTimeStamp)_errors.log" -Append
	write-host "Script ended with $($Error.length) errors!"
	$error.clear()
	}
Else {write-host "Script gracefully ended without errors!"}
Write-Host "`r`nScript completed in $((`"{0:hh\:mm\:ss}`" -f [timespan]::FromSeconds($TotalScriptTimer.Elapsed.TotalSeconds))). Enjoy your new reports!"
Stop-Transcript  | out-null


