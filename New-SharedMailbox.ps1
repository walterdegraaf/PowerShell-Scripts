<#
.SYNOPSIS
    Creates a Shared Mailbox in Exchange Online and creates Security Groups for access.
.DESCRIPTION
    This command will create a Shared Mailbox in Exchange Online, will create Security Groups in the right OU and will give the Security Groups FullAccess and SendAs rights.
.EXAMPLE
    PS C:\> New-SharedMailbox -Name "DM | Sekretariat" -EmailAddress sekretariat@drewsmarine.com -OuName DE-HAM-DRE -SecurityGroupShortName Sekretariat
    Explanation of what the example does
.INPUTS
    Name
    EmailAddress
    OuName
    SecurityGroupShortName
.OUTPUTS
    Output (if any)
.NOTES
    SecurityGroupShortName should be a short name that will be added to the security group name. The companyname + SM is automatically added.
#>
function New-SharedMailbox {
    param (
        $Name,
        $EmailAddress,
        $OuName,
        $SecurityGroupShortName
    )

    $ouNameFull = Get-ADOrganizationalUnit -Filter 'Name -like $ouName'
    #name of the group that's need to be created
    $SharedMailboxSecurityGroupName = "$ouName".ToUpper() + "-SM-" + $SecurityGroupShortName

    #Connect to Exchange Online and create the mailbox
    Connect-ExchangeOnline
    Write-Host -ForegroundColor Yellow "Creating the mailbox..."
    New-Mailbox -Shared -Name $Name -Displayname $Name -PrimarySmtpAddress $EmailAddress

    #Create two new groups: one for FullAccess(FA) and one for SendAs(SA)
    $SharedMailboxSecurityGroupNameFA= $SharedMailboxSecurityGroupName + "-FA"
    $SharedMailboxSecurityGroupNameSA = $SharedMailboxSecurityGroupName + "-SA"
    Write-Host -ForegroundColor Yellow "Creating the Security Groups in Active Directory..."
    New-ADGroup -Name $SharedMailboxSecurityGroupNameFA -path $ouNameFull -GroupScope Global
    New-ADGroup -Name $SharedMailboxSecurityGroupNameSA -path $ouNameFull -GroupScope Global

    #Add mail properties to the Security Groups
    $SetEmailToGroupFA = $SharedMailboxSecurityGroupNameFA + "@domain.com"
    $SetEmailToGroupSA = $SharedMailboxSecurityGroupNameSA + "@domain.com"
    Set-ADGroup $SharedMailboxSecurityGroupNameFA -Replace @{mail=$SetEmailToGroupFA}
    Set-ADGroup $SharedMailboxSecurityGroupNameSA -Replace @{mail=$SetEmailToGroupSA}

    #Sync changes to Azure, need this to create the mail-enabled-security-grops in Exchange Online
    Invoke-Command -ComputerName sa-sr-ams-as-01.mc1s.com -ScriptBlock {Start-ADSyncSyncCycle -PolicyType delta} 
    #Wait 5 minutes for the sync to complete
    Write-Host -ForegroundColor Yellow "Sync to Azure AD is in progress and will take 5 minutes..."
    Start-Sleep -Seconds 300
    Write-Host -ForegroundColor Green "Sync completed"
    
    Add-MailboxPermission -Identity $EmailAddress -User $SetEmailToGroupFA -AccessRights FullAccess
    Add-RecipientPermission -Identity $EmailAddress -Trustee $SetEmailToGroupSA -AccessRights SendAs -Confirm:$false
}
