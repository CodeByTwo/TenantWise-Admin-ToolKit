<#
.SYNOPSIS
Creates a service principal in Entra ID Tenant to be used with TenantWise SharePoint integration.

.DESCRIPTION
Checks if the Az PowerShell module is installed, handles its installation, connects to Entra ID, and creates a service principal. Captures and displays the auto-generated secret.
#>

# Function to check and install the Az PowerShell module
function Ensure-AzModule {
    Write-Host "Checking if the Az PowerShell module is installed..." -ForegroundColor Yellow
    if (-not (Get-Module -ListAvailable -Name Az)) {
        Write-Host "Az PowerShell module is not installed." -ForegroundColor Yellow
        $installAz = Read-Host -Prompt "Would you like to install the Az PowerShell module? (y/n)"
        if ($installAz -eq 'y') {
            Write-Host "Installing Az PowerShell module..." -ForegroundColor Yellow
            try {
                Install-Module -Name Az -AllowClobber -Scope CurrentUser -Force
                Write-Host "Az PowerShell module installed successfully." -ForegroundColor Green
            } catch {
                Write-Host "Failed to install Az PowerShell module. Please install it manually and try again." -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "You chose not to install the Az PowerShell module. The script cannot proceed." -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "Az PowerShell module is already installed." -ForegroundColor Green
    }
    Import-Module -Name Az.Accounts -Force
}

# Ensure the Az PowerShell module is installed
Ensure-AzModule

# Prompt the user for the Tenant ID
$tenantId = Read-Host -Prompt "Enter your Azure AD Tenant ID"

# Validate the Tenant ID format
if ($tenantId -notmatch '^[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}$') {
    Write-Host "Invalid Tenant ID format. Please enter a valid GUID." -ForegroundColor Red
    exit
}

# Connect to Azure with the provided Tenant ID
Write-Host "Connecting to Azure with Tenant ID: $tenantId..." -ForegroundColor Yellow
try {
    Connect-AzAccount -TenantId $tenantId | Out-Null
    Write-Host "Successfully connected to Azure!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Azure. Please check your Tenant ID and credentials." -ForegroundColor Red
    exit
}

# Prompt for the service principal name
$spName = Read-Host -Prompt "Enter a name for the service principal"

# Create the service principal
Write-Host "Creating the service principal: $spName..." -ForegroundColor Yellow
try {
    $sp = New-AzADServicePrincipal -DisplayName $spName
    Write-Host "Service Principal created successfully!" -ForegroundColor Green
    Write-Host "Service Principal Application (Client) ID: $($sp.AppId)"
    Write-Host "Client Secret Value (save this securely): $($sp.PasswordCredentials.SecretText)"
    Write-Host "Client Secret expires on: $($sp.PasswordCredentials.EndDateTime)"
    Write-Host "Ensure you copy the secret value now. It cannot be retrieved again!" -ForegroundColor Red
} catch {
    Write-Host "Failed to create the service principal. Error: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Disconnect from Azure
Disconnect-AzAccount | Out-Null

Write-Host "Script execution complete!" -ForegroundColor Cyan