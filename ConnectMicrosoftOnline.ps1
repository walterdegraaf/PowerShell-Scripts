function ConnectAzure {
$AzureModule = Get-Module -Name "AzureAD" -ErrorAction SilentlyContinue
$ImportAzureModule = Import-Module "AzureAD" -ErrorAction SilentlyContinue -ErrorVariable ImportError
    if (!$AzureModule) {
        $ImportAzureModule                      
        if ($ImportError) {
            Write-Warning "Module is not installed..."
            Write-Host "Installing the module..."
            install-module AzureAD
        }
        Write-Host "Importing the module..."
        $ImportAzureModule
        
    }
    
    Write-Host -ForegroundColor "Green" "Connecting to AzureAD..."
    Connect-AzureAD
}

function ConnectMSOL {
    $MSOLModule = Get-Module -Name "MSOnline" -ErrorAction SilentlyContinue
    $ImportMSOLModule = Import-Module "MSOnline" -ErrorAction SilentlyContinue -ErrorVariable ImportError
        if (!$MSOLModule) {
            $ImportMSOLModule                     
            if ($ImportError) {
                Write-Warning "Module is not installed..."
                Write-Host "Installing the module..."
                install-module MSOnline
            }
            Write-Host "Importing the module..."
            $ImportMSOLModule
            
        }
        
        Write-Host -ForegroundColor "Green" "Connecting to MSOnline..."
        Connect-MSOLservice
    }

function ConnectExchange {
    $ExchangeModule = Get-Module -Name "ExchangeOnlineManagement" -ErrorAction SilentlyContinue
    $ImportExchangeModule = Import-Module "ExchangeOnlineManagement" -ErrorAction SilentlyContinue -ErrorVariable ImportError
        if (!$ExchangeModule) {
            $ImportExchangeModule                     
            if ($ImportError) {
                Write-Warning "Module is not installed..."
                Write-Host "Installing the module..."
                install-module ExchangeOnlineManagement
            }
            Write-Host "Importing the module...."
            $ImportExchangeModule
            
        }
        
        Write-Host -ForegroundColor "Green" "Connecting to Exchange Online..."
        Connect-ExchangeOnline
    }
