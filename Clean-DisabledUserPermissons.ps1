# ==============================================================================================
# 
# NAME: Clean-DisabledUserPermissons.ps1
# 
# AUTHOR: Tito D Castillote Jr , june.castillote@gmail.com
# DATE  : 9/17/2015
# 
# COMMENT:
# v1.0 - 9/17/2015
#	- Can remove "Send on Behalf" and "Accept Messages From"
#	- Valid target types: Mailbox
#	- Valid input: DistinguishedName,SamAccountName,GUID
# v1.1 - 9/17/2015
#	- Added Distribution Group and Dynamic Distribution Group as valid targets
#	- Valid inputs reduced to just DistinguishedName and GUID
# 
# ==============================================================================================
$scriptVersion = "1.1"
Clear-Host
Write-Host '====================================' -ForegroundColor Yellow
Write-Host "  Clean-DisabledUserPermissons v$($scriptVersion)   " -ForegroundColor Yellow
Write-Host '====================================' -ForegroundColor Yellow
Write-Host ''
Write-Host 'Description: This script will remove the permissions of Disabled Users from any of the following targets - Mailbox, Distribution Group, Dynamic Distribution Group' -ForegroundColor Yellow

#Import ActiveDirectory Module, exit if not found
if (!(Get-Module | where {$_.Name -eq "activedirectory"}))
	{
		try
		{
			Write-Host 'Importing ActiveDirectory Module' -ForegroundColor Yellow
			Import-Module activedirectory -ErrorAction STOP
		}
		catch
		{
			Write-Warning $_.Exception.Message
			EXIT
		}
	}

Write-Host @(
'
Accepted Values:
	DN (eg. CN=admin_castillt,OU=IT Admin Accounts,OU=User Accounts,DC=boral,DC=com,DC=au)
	GUID (eg. 3b3d7ea5-666e-4bf6-bf30-c215c8aa98a5)
	'
) -ForegroundColor Yellow


#Prompt for the AD Identity of the Mailbox
while (($ObjectToCheck = Read-Host "Please enter the Identity of the target object (DN or GUID)") -eq "") {
	Clear-Host
	Write-Host 'The "AD Object" cannot be blank.Please try again...' -ForegroundColor RED
}

#Lookup the AD Identity, exit if not found
try
		{
			Write-Host 'Checking if the AD Object Exists' -ForegroundColor Yellow
			$mObject = Get-AdObject -Identity $ObjectToCheck -Properties *
		}
		catch
		{
			Write-Host $_.Exception.Message -ForegroundColor RED
			Write-Host "Please re-run the script and try again to enter as valid AD Object" -ForegroundColor Green
			EXIT
		}

#Loop thru the permissions and remove the disabled users
ForEach ($object in $mObject) {
	$tempObj = ""
	Write-Host ""
	Write-Host "Processing: $($object.Name)" -ForegroundColor Green
	Write-Host ""
	Write-Host ">>> REMOVING DISABLED Send on Behalf" -ForegroundColor Cyan
	Write-Host ""
	
	if (($object.PublicDelegates) -gt 0) {
		ForEach ($xItem in $object.PublicDelegates) {
			$tempObj = Get-ADUser $xItem
			if ($tempObj.Enabled -eq $false){
				Write-Host "Remove - $($tempObj.DistinguishedName)" -ForegroundColor Yellow
				Set-AdObject -Identity $ObjectToCheck -remove @{PublicDelegates="$($tempObj.DistinguishedName)"}
			}
		}
	}
	else {		
		Write-Host 'No Disabled User Found with Send on Behalf Permission' -ForegroundColor Green	
	}
		
	Write-Host ""
	Write-Host ">>> REMOVING DISABLED Accept Messages From" -ForegroundColor Cyan
	Write-Host ""
	
	if (($object.AuthOrig) -gt 0) {
		ForEach ($zItem in $object.AuthOrig) {
			$tempObj = Get-ADUser $zItem
			if ($tempObj.Enabled -eq $false){
				Write-Host "Remove - $($tempObj.DistinguishedName)" -ForegroundColor Yellow
				Set-AdObject -Identity $ObjectToCheck -remove @{AuthOrig="$($tempObj.DistinguishedName)"}
			}
		}
	}
	else {
		Write-Host 'No Disabled User Found with Allowed Sender Permission' -ForegroundColor Green	
	}
	Write-Host ""
	Write-Host "Clean-up Done." -ForegroundColor Green
}