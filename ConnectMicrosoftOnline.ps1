function ConnectExchange {
    param (
        $UPN
    )
    if ((Get-Module) -notcontains "ExchangeOnlineManagement") {
        install-module ExchangeOnlineManagement
        Import-Module ExchangeOnlineManagement
    }  
    Connect-ExchangeOnline -UserPrincipalName $UPN -ShowProgress $true
}

function ConnectMSOL {
    param (
        $UPN
    )
    if ((Get-Module) -notcontains "MSOnline") {
        install-module MSOnline
        import-module MSOnline    
    }
    Connect-MSOLservice
}

function ConnectAzure {

    if ((Get-Module) -notcontains "AzureAD") {
        install-module AzureAD
        Import-Module AzureAD -ErrorAction SilentlyContinue
    }
    Connect-AzureAD
}
