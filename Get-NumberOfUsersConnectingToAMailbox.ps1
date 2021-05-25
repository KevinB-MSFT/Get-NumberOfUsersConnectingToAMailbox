#########################################################################################
# LEGAL DISCLAIMER
# This Sample Code is provided for the purpose of illustration only and is not
# intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
# RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
# EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
# MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
# nonexclusive, royalty-free right to use and modify the Sample Code and to
# reproduce and distribute the object code form of the Sample Code, provided
# that You agree: (i) to not use Our name, logo, or trademarks to market Your
# software product in which the Sample Code is embedded; (ii) to include a valid
# copyright notice on Your software product in which the Sample Code is embedded;
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
# against any claims or lawsuits, including attorneys’ fees, that arise or result
# from the use or distribution of the Sample Code.
# 
# This posting is provided "AS IS" with no warranties, and confers no rights. Use
# of included script samples are subject to the terms specified at 
# https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx.
#
# Exchange Online Device partnership inventory
# Get-NumberOfUsersConnectingToAMailbox
#  
# Created by: Kevin Bloom 05/25/2021 Kevin.Bloom@Microsoft.com  
#
# The script is intended to ran in an Exchange Management Shell Session
# Requires an email address to be passed through when running the script
# Script will do the following:
# 1. Get the Mailboxes information such as Database and ExchangeGUID
# 2. Retrieve the Store Query information
# 3. Export the Store Query raw data
# 4. Filter and de-duplicat the Store Query data
# 5. Write the Store Query data to the screen
#
# How to Run the script
# .\Get-NumberOfUsersConnectingToAMailbox.ps1 -EmailAddress <Email Address>
#
#########################################################################################
#
#########################################################################################
param (
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string] $EmailAddress
)
#This section will get the current script location, import the ManagedStoreDiagnosticFunctions.ps1 as it is needed to run Get-StoreQuery, and change the lcoation back to original location
$Location = Get-Location
CD $exscripts
. .\ManagedStoreDiagnosticFunctions.ps1
cd $Location

#Initializes the variables
$Date = Get-Date -Format "MMddyyyy_HHmm"
$DB = (Get-Mailbox -Identity $EmailAddress).database.name
$MailboxGUID = (Get-Mailbox -Identity $EmailAddress).exchangeguid.guid

#Retrieves the Store Query information and exports the raw data to CSV and XML for reference
$StoreQuery = Get-StoreQuery -Database $DB -Query "SELECT * FROM Session WHERE LastUsedMailboxGuid = '$MailboxGUID'"
$StoreQuery | Export-Csv .\StoreQueryRaw_$Date.csv
$StoreQuery | Export-Clixml .\StoreQueryRaw_$Date.xml

#Filters, sorts, and deduplicats the raw data
$StoreQueryDedup = $StoreQuery | Where-Object {$_.ApplicationId -notlike "*viaProxy" -and $_.userDN -notlike "*Microsoft System Attendant"} | Sort-Object UserDN  | Group-Object -Property UserDN | select @{n="list";e={ $_.group | select -first 1  }} | select -ExpandProperty list

#Writes the number of users and who are connecting to the mailbox
Write-Host -ForegroundColor Cyan "There are $($StoreQueryDedup.Count) active connections to Shared Mailbox, $EmailAddress.  Complete list of DNs:"
foreach ($user in $StoreQueryDedup.userdn) {Write-Host -ForegroundColor Cyan $user}
