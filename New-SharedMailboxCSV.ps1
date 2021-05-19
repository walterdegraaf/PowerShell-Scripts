<#
.SYNOPSIS
    Asks for a CSV file and will create Shared Mailboxes in Exchange Online.
.DESCRIPTION
    This command will create a Shared Mailbox in Exchange Online, will create Security Groups in the right OU and will give the Security Groups FullAccess and SendAs rights.
.EXAMPLE
    PS C:\> New-SharedMailboxCSV -CSVLocation C:\users\User\Desktop\import.csv
.INPUTS
    CSV Location
.OUTPUTS

.NOTES

#>
function New-SharedMailboxCSV {
    param (
        $CSVLocation
    )

    $Mailboxes = Import-Csv $CSVLocation -Delimiter ";"
    Connect-ExchangeOnline

    foreach ($mailbox in $Mailboxes) {
        $Name = $mailbox.name
        $EMailAddress = $mailbox.emailaddress
        $OuName = $mailbox.OuName
        $SecurityGroupShortName = $mailbox.SecurityGroupShortName
        

        $ouNameFull = Get-ADOrganizationalUnit -Filter 'Name -like $ouName'
        #name of the group that's need to be created
        $SharedMailboxSecurityGroupName = "$ouName".ToUpper() + "-SM-" + $SecurityGroupShortName

        #Connect to Exchange Online and create the mailbox
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
    }

    #Sync changes to Azure, need this to create the mail-enabled-security-grops in Exchange Online
    Invoke-Command -ComputerName sa-sr-ams-as-01.mc1s.com -ScriptBlock {Start-ADSyncSyncCycle -PolicyType delta} 
    Write-Host -ForegroundColor Yellow "Sync to Azure AD is in progress and will take 5 minutes..."
    #Wait 5 minutes for the sync to complete
    Start-Sleep -Seconds 300
    Write-Host -ForegroundColor Green "Sync completed"

    foreach ($mailbox in $Mailboxes) {     
         Add-MailboxPermission -Identity $EmailAddress -User $SetEmailToGroupFA -AccessRights FullAccess
         Add-RecipientPermission -Identity $EmailAddress -Trustee $SetEmailToGroupSA -AccessRights SendAs -Confirm:$false
    }
}
