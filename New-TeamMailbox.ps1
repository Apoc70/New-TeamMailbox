<# 
    .SYNOPSIS 
    Creates a new shared mailbox, security groups for full access and send-as permission 
    and adds the security groups to the shared mailbox configuration.

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2016-07-08

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    More information can be found at http://www.granikos.eu/en/scripts

    .DESCRIPTION 
    This scripts creates a new shared mailbox (aka team mailbox) and security groups
    for full access and and send-as delegation. Security groups are created using a
    naming convention.
 
    .NOTES 
    Requirements 
    - Windows Server 2012 R2 
    - Exchange 2013/2016 Management Shell (aka EMS)

    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 
   
    .PARAMETER TeamMailboxName
    Name attribute of the new team mailbox

    .PARAMETER TeamMailboxDisplayName
    Display name attribute of the new team mailbox

    .PARAMETER TeamMailboxAlias
    Alias attribute of the new team mailbox

    .PARAMETER TeamMailboxSmtpAddress
    Primary SMTP address attribute the new team mailbox

    .PARAMETER DepartmentPrefix
    Department prefix for automatically generated security groups (optional)

    .PARAMETER GroupFullAccessMembers
    String array containing full access members
    
    .PARAMETER GroupFullAccessMembers
    String array containing send as members

    .EXAMPLE 
    Create a new team mailbox, empty full access and empty send-as security groups

    .\New-TeamMailbox.ps1 -TeamMailboxName "TM-Exchange Admins" -TeamMailboxDisplayName "Exchange Admins" -TeamMailboxAlias "TM-ExchangeAdmins" -TeamMailboxSmtpAddress "ExchangeAdmins@mcsmemail.de" -DepartmentPrefix "IT"
#> 
param (
    [parameter(Mandatory=$true,HelpMessage='Team Mailbox Name')]
        [string]$TeamMailboxName,
    [parameter(Mandatory=$true,HelpMessage='Team Mailbox Display Name')]
        [string]$TeamMailboxDisplayName,
    [parameter(Mandatory=$true,HelpMessage='Team Mailbox Alias')]
        [string]$TeamMailboxAlias,
    [parameter(Mandatory=$true,HelpMessage='Team Mailbox Email Address')]
        [string]$TeamMailboxSmtpAddress,
    [parameter(Mandatory=$false,HelpMessage='Department prefix for group name')] 
        [string]$DepartmentPrefix = "",
    [parameter(Mandatory=$false,HelpMessage='Full access group members')]
        $GroupFullAccessMembers = @(''),
    [parameter(Mandatory=$false,HelpMessage='Send as group members')]
        $GroupSendAsMember = @('')
)

# Script Path
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path


if(Test-Path "$($scriptPath)\Settings.xml") {
    # Load Script settings
    [xml]$Config = Get-Content -Path "$($scriptPath)\Settings.xml"
    
    Write-Verbose 'Loading script settings'
    
    # Group settings
    $groupPrefix = $Config.Settings.GroupSettings.Prefix
    $groupSendAsSuffix = $Config.Settings.GroupSettings.SendAsSuffix
    $groupFullAccessSuffix = $Config.Settings.GroupSettings.FullAccessSuffix
    $groupTargetOU = $Config.Settings.GroupSettings.TargetOU
    $groupDomain = $Config.Settings.GroupSettings.Domain
    
    # Team mailbox settings
    $teamMailboxTargetOU = $Config.Settings.AccountSettings.TargetOU

    # General settings
    $sleepSeconds = $Config.Settings.GeneralSettings.Sleep

    Write-Verbose 'Script settings loaded'    
}
else {
    Write-Error 'Script settings file settings.xml missing'
    exit 99
}

# add department prefix to group prefix, if configured
if($DepartmentPrefix -ne '') {
    # Change pattern as needed
    $groupPrefix = "$($groupPrefix)$($DepartmentPrefix)_"
}

# Create shared team mailbox
Write-Verbose "New-Mailbox -Shared -Name $($TeamMailboxName) -Alias $($TeamMailboxAlias)"

# New-Mailbox -Shared -Name $TeamMailboxName -Alias $TeamMailboxAlias -OrganizationalUnit $teamMailboxTargetOU -PrimarySmtpAddress $TeamMailboxSmtpAddress -DisplayName $TeamMailboxDisplayName | Out-Null

# Create Full Access group
$groupName = "$($groupPrefix)$($TeamMailboxAlias)$($groupFullAccessSuffix)"
$notes = "FullAccess for mailbox: $($TeamMailboxName)"
$primaryEmail = "$($groupName)@$($groupDomain)"

Write-Host "Creating new FullAccess Group: $($groupName)"

Write-Verbose "New-DistributionGroup -Name $($groupName) -Type Security -OrganizationalUnit $($groupTargetOU) -PrimarySmtpAddress $($primaryEmail)"

if(($GroupFullAccessMembers | Measure-Object).Count -ne 0) {

    Write-Host "Creating FullAccess group and adding members: $($groupName)"

    New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $GroupFullAccessMembers -Notes $notes | Out-Null

    Start-Sleep -Seconds $sleepSeconds

    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}
else {

    Write-Host "Creating empty FullAccess group: $($groupName)"

    New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes | Out-Null

    Start-Sleep -Seconds $sleepSeconds
    
    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}

# add full access group

Write-Verbose "Add-MailboxPermission -Identity $($TeamMailboxSmtpAddress) -User $($primaryEmail)"

Add-MailboxPermission -Identity $TeamMailboxSmtpAddress -User $primaryEmail -AccessRights FullAccess -InheritanceType all | Out-Null

# Create Send As group
$groupName = "$($groupPrefix)$($TeamMailboxAlias)$($groupSendAsSuffix)"
$notes = "SendAs for mailbox: $($TeamMailboxName)"
$primaryEmail = "$($groupName)@$($groupDomain)"

Write-Host "Creating new SendAs Group: $($groupName)"

Write-Verbose "New-DistributionGroup -Name $($groupName) -Type Security -OrganizationalUnit $($groupTargetOU) -PrimarySmtpAddress $($primaryEmail)"

if(($GroupSendAsMember | Measure-Object).Count -ne 0) {

    Write-Host "Creating SendAs group and adding members: $($groupName)"

    New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $GroupSendAsMember -Notes $notes | Out-Null

    Start-Sleep -Seconds $sleepSeconds

    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}
else {

    Write-Host "Not empty SendAs group: $($groupName)"

    New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes | Out-Null

    Start-Sleep -Seconds $sleepSeconds

    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}

# Add SendAs permission
Write-Verbose "Add-ADPermission -Identity $($TeamMailboxName) -User $($groupName)"

Add-ADPermission -Identity $TeamMailboxName -User $groupName -ExtendedRights "Send-As" | Out-Null

Write-Host "Script finished. Team mailbox $($TeamMailboxName) created." 