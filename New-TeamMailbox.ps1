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
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 

    Parameters are not yet supported.

   
    .EXAMPLE 
    .\New-TeamMailbox.ps1
#> 

# Team mailbox parameters
$teamMailboxName = 'MB-COM-Marketing'
$teamMailboxDisplayName = 'Marketing'
$teamMailboxAlias = 'MB-COM-Marketing'
$teamMailboxSmtpAddress = 'marketing@mcsmemail.de'
$departmentPrefix = "COM" # e.g. COM

# Team mailbox target OU
$teamMailboxTargetOU = 'mcsmemail.de/ORG/IT/SharedMailboxes'

# Group members
$groupFullAccessMembers = @('johndoe','janedoe') 
$groupSendAsMember = @()

# Permission group parameters
$groupPrefix = 'sec_'
$groupSendAsSuffix = '_SA'
$groupFullAccessSuffix = '_FA'
$groupTargetOU = 'mcsmemail.de/ORG/IT/Groups'

# add department prefix, if configured
if($departmentPrefix -ne '') {
    # Change pattern as needed
    $groupPrefix = "$($groupPrefix)$($departmentPrefix)_"
}

# additional variables
$sleepSeconds = 10

# Create shared team mailbox
New-Mailbox -Shared -Name $teamMailboxName -Alias $teamMailboxAlias -OrganizationalUnit $teamMailboxTargetOU -PrimarySmtpAddress $teamMailboxSmtpAddress -DisplayName $teamMailboxDisplayName

# Create Full Access group
$groupName = "$($groupPrefix)$($teamMailboxAlias)$($groupFullAccessSuffix)"
$notes = "FullAccess for mailbox: $($teamMailboxName)"
$primaryEmail = "$($groupName)@mcsmemail.de"

Write-Host "Creating new FullAccess Group: $($groupName)"
if(($groupFullAccessMembers | Measure-Object).Count -ne 0) {
    Write-Host "Creating FullAccess group and adding members: $($groupName)"
    New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $groupFullAccessMembers -Notes $notes 
    Start-Sleep -Seconds 5
    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}
else {
    Write-Host "Creating empty FullAccess group: $($groupName)"
    New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes
    Start-Sleep -Seconds 5
    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}

# add full access group
Add-MailboxPermission -Identity $teamMailboxSmtpAddress -User $primaryEmail -AccessRights FullAccess -InheritanceType all

# Create Send As group
$groupName = "$($groupPrefix)$($teamMailboxAlias)$($groupSendAsSuffix)"
$notes = "SendAs for mailbox: $($teamMailboxName)"
$primaryEmail = "$($groupName)@mcsmemail.de"

Write-Host "Creating new SendAs Group: $($groupName)"
if(($groupSendAsMember | Measure-Object).Count -ne 0) {
    Write-Host "Creating SendAs group and adding members: $($groupName)"
    New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Members $groupSendAsMember -Notes $notes
    Start-Sleep -Seconds $sleepSeconds
    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}
else {
    Write-Host "Not empty SendAs group: $($groupName)"
    New-DistributionGroup -Name $groupName -Type Security -OrganizationalUnit $groupTargetOU -PrimarySmtpAddress $primaryEmail -Notes $notes
    Start-Sleep -Seconds $sleepSeconds
    Set-DistributionGroup -Identity $primaryEmail -HiddenFromAddressListsEnabled $true
}

# Add SendAs permission
Add-ADPermission -Identity $teamMailboxName -User $primaryEmail -ExtendedRights "Send-As" 