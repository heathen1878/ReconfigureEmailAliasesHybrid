<#

.SYNOPSIS
This script take input interactively and reconfigures the primary SMTP address of all mailboxes and distribution groups which have a specific accepted domain.
By default the script runs in test mode, to make changes the argument live should be supplied. 

.DESCRIPTION


.EXAMPLE
ReconfigureEmailAddressesHybrid.ps1 [options]
possible options are live or leave it blank.  

.NOTES
The argument on the script is not case sensitive.

.LINK
https://github.com/heathen1878/ReconfigureEmailAliasesHybrid

#>

# enforces coding rules in expressions, scripts, and script blocks
Set-StrictMode -Version Latest

# Import the necessary functions
. .\Functions.ps1

# Declare variables
[Boolean]$bValidator = $false
[string]$sQuery = ""
[Array]$arrValid = @()
[System.Collections.ArrayList]$arrNewEmailAddresses = @()
[String]$sRemoveFromArray = ""
[PsCredential]$Creds = $null
[String]$sNewPrimaryDomain = ""
[String]$sCurrentPrimaryDomain = ""
[String]$sNewPrimarySMTPAddress = ""
[String]$sUserAlias = ""
[String]$sRunType = ""

# Start of script.
WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Script started...$(Get-Date)"

# Check for arguments on the command line.
switch ($args.count) {

	# If no arguments are specified then run in test mode else run in change mode
	0 {
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Running dry run. Results will be available in the log file"
		$sRunType = "Dry"
	}
	1 {
		If ($args[0] -eq "live"){

			WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Running live. Results will be available in the log file and changes will be made."
			$sRunType = "Live"

		}
		Else{

			Write-Output "Invalid argument specified."
            break;  

		}
	}
}

## Verification of the domain should be tested by connecting to Microsoft Online
While (!$bValidator) {
	
	# Get credentials for Office 365.
	$Creds = Get-Credential -Message "Please enter your Office 365 credentials"
	$bValidator = ConnectToMicrosoftOnline -credentials $creds

}

# Reset the validator
$bValidator = $false

# Get the new domain name from the console
While (!$bValidator) {

	$sNewPrimaryDomain = Read-Host "What domain will be the new primary SMTP domain?"	
	If (Get-MsolDomain -DomainName $sNewPrimaryDomain -ErrorAction SilentlyContinue) {
		$bValidator = $True
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$sNewPrimaryDomain validated"
	}
	else {
		Write-Output "$sNewPrimaryDomain not valid"
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$sNewPrimaryDomain not valid"
	}
}

# Reset the validator
$bValidator = $false

# Get the new domain name from the console
While (!$bValidator) {

	$sCurrentPrimaryDomain = Read-Host "What domain is being replaced?"
	If (Get-MsolDomain -DomainName $sCurrentPrimaryDomain -ErrorAction SilentlyContinue) {
		$bValidator = $True
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$sCurrentPrimaryDomain validated"
	}
	else {
		Write-Output "$sCurrentPrimaryDomain not valid"
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$sCurrentPrimaryDomain not valid"
	}
}

# Reset the validator
$bValidator = $false

# Set validator
$arrValid = ("Place holder")
$arrValid += ($false)

# Call connect to Exchange Online
While (!$arrValid[1]) {
	
	$arrValid = ConnectToExchange -Credentials $creds -Location "R"
	If (!$arrValid[1]){
		$Creds = Get-Credential -Message "Please enter your Office 365 credentials"	
	}	
	Else {
		Import-PSSession $arrValid[0]
	}
}

# Create a query for the Get-Mailbox cmdlet.
$sQuery = -Join ("*@",$sCurrentPrimaryDomain)

# Get a collection of mailboxes which have the current domain has a primary SMTP address. This will work with regular mailboxes, resource mail and shared mailboxes.
Get-Mailbox -ResultSize Unlimited | Where-Object {$_.PrimarySmtpAddress -like $sQuery} | ForEach-Object {
		
	Try 
	{	
		# required for PowerShell 4.0 remoting.
		$Global:ErrorActionPreference="Stop"
		
		$sUserAlias = $_.Alias
		
		# Get the current user in the pipeline into a variable 
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Processing user: $_...$(Get-Date)"

		# convert all instances of SMTP to smtp - believe it or not that is the difference
		# between primary and alias email addresses. 
		$arrNewEmailAddresses = $_.EmailAddresses.Replace("SMTP","smtp")

        WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Email address collection: $arrNewEmailAddresses"
		
		# Get the existing primary SMTP address and use it to create the new primary SMTP address
		# e.g. user@domain.com to user@domain.uk...
		$sNewPrimarySMTPAddress = $_.PrimarySmtpAddress.Replace($sCurrentPrimaryDomain,$sNewPrimaryDomain)
		
        WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "New Email address: $sNewPrimarySMTPAddress"

        # Remove the instance of the domain being added
        
        # Get the primary domain which exists in the array
        ### Need to check if it exists, if not then the add primary domain will suffice. ####
        $arrNewEmailAddresses | ForEach-Object {
           If ($_ -match $sNewPrimaryDomain){
                
                $sRemoveFromArray = $_
        
            }

        }

        # Check whether anything needs removing from the array
        If (-not $sRemoveFromArray -eq ''){

            WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Removing: $sRemoveFromArray"
    
            [void]$arrNewEmailAddresses.Remove($sRemoveFromArray)

        }

		# Make the new SMTP address primary by prefixing it with a uppercase SMTP.
		$sNewPrimarySMTPAddress = -Join ("SMTP:",$sNewPrimarySMTPAddress)
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "New primary SMTP address will be $sNewPrimarySMTPAddress"
		[void]$arrNewEmailAddresses.Add($sNewPrimarySMTPAddress)
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Addresses to be configured: $arrNewEmailAddresses"

		If ($sRunType -eq "live"){

			# Update the email addresses associated with the mailbox
			#Set-Mailbox -Identity $sUserAlias -EmailAddresses $arrNewEmailAddresses
			
			# Check our work.
			Write-Output "Checking the configuration" -foregroundcolor Yellow
			Get-mailbox -Identity $sUserAlias | select-Object PrimarySmtpAddress

		}	
	}		
	Catch 
	{
			WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$Error[0].Exception.Message"	
	}
	Finally
	{
		# Clear variables for next loop 
		clear-variable sUserAlias
		clear-variable sNewPrimarySMTPAddress
		clear-variable arrNewEmailAddresses
        clear-variable sRemoveFromArray
		$Global:ErrorActionPreference="Continue"
		$Error.Clear()
	}
}

# Get a collection of distribution groups which have the current domain has a primary SMTP address.
Get-DistributionGroup -ResultSize Unlimited | Where-Object {$_.PrimarySmtpAddress -like $sQuery} | ForEach-Object {
		
	Try 
	{	
		# required for PowerShell 4.0 remoting.
		$Global:ErrorActionPreference="Stop"
		
		$sUserAlias = $_.Alias
		
		# Get the current user in the pipeline into a variable
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Processing distribution group: $_...$(Get-Date)"
        
        # convert all instances of SMTP to smtp - believe it or not that is the difference
		# between primary and alias email addresses. 
		$arrNewEmailAddresses = $_.EmailAddresses.Replace("SMTP","smtp")

        WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Email address collection: $arrNewEmailAddresses"
		
		# Get the existing primary SMTP address and use it to create the new primary SMTP address
		# e.g. user@domain.com to user@domain.uk...
		$sNewPrimarySMTPAddress = $_.PrimarySmtpAddress.Replace($sCurrentPrimaryDomain,$sNewPrimaryDomain)
		
        WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "New Email address: $sNewPrimarySMTPAddress"

        # Remove the instance of the domain being added
        
        # Get the primary domain which exists in the array
        ### Need to check if it exists, if not then the add primary domain will suffice. ####
        $arrNewEmailAddresses | ForEach-Object {
        
            If ($_ -match $sNewPrimaryDomain){
                
                $sRemoveFromArray = $_
        
            }

        }

                Write-Output $sRemoveFromArray

        # Check whether anything needs removing from the array
        If (-not $sRemoveFromArray -eq ''){

            WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Removing: $sRemoveFromArray"
    
            # Don't return update to screen
            [void]$arrNewEmailAddresses.Remove($sRemoveFromArray)

        }

		# Make the new SMTP address primary by prefixing it with a uppercase SMTP.
		$sNewPrimarySMTPAddress = -Join ("SMTP:",$sNewPrimarySMTPAddress)
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "New primary SMTP address will be $sNewPrimarySMTPAddress"
		
        # Don't return update to screen
        [void]$arrNewEmailAddresses.Add($sNewPrimarySMTPAddress)
		WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "Addresses to be configured: $arrNewEmailAddresses"

		if ($sRunType -eq "live"){

			# Update the email addresses associated with the mailbox
			#Set-DistributionGroup -Identity $sUserAlias -EmailAddresses $arrNewEmailAddresses
		
			# Check our work.
			Write-Output "Checking the configuration" -foregroundcolor Yellow
			Get-DistributionGroup -Identity $sUserAlias | select-Object PrimarySmtpAddress
		}
	}		
	Catch 
	{
		# Any error generated in the Try block are written to a log file in this Catch block.
		If ($Error[0].exception.message -match "is already present in the collection"){
			Write-Output "Error caught. Check the log file for more information." -foregroundcolor Yellow
			WriteToLog -sLogFile "setPrimarySmtpAddress.log" -sLogContent "$Error[0].Exception.Message"
		}
		
	}
	Finally
	{
		# Clear variables for next loop 
		clear-variable sUserAlias
		clear-variable sNewPrimarySMTPAddress
		clear-variable arrNewEmailAddresses
        Clear-Variable sRemoveFromArray
		$Global:ErrorActionPreference="Continue"
		$Error.Clear()
	}
}

# Clean up the remote PowerShell session created earlier.
Clear-Variable arrValid
Get-PSSession | Remove-PSSession