<#
.SYNOPSIS
Grants the `Sites.Selected` Microsoft Graph API permission to a service principal.

.DESCRIPTION
Prompts for an Application ID, assigns the `Sites.Selected` permission, and optionally requires admin consent.
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

# Prompt for Application ID
$appId = Read-Host -Prompt "Enter the Application ID of the service principal"

# Validate the Application ID format
if ($appId -notmatch '^[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}$') {
    Write-Host "Invalid Application ID format. Please enter a valid GUID." -ForegroundColor Red
    exit
}

# Connect to Azure
Write-Host "Connecting to Azure..." -ForegroundColor Yellow
try {
    Connect-AzAccount -Tenant $tenantId | Out-Null
    Write-Host "Successfully connected to Azure!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Azure. Please check your credentials." -ForegroundColor Red
    exit
}

# Fetch the service principal using the Application ID
Write-Host "Fetching the service principal for Application ID: $appId..." -ForegroundColor Yellow
try {
    $sp = Get-AzADApplication -ApplicationId $appId
    Write-Host "Service Principal found: $($sp.DisplayName)" -ForegroundColor Green
} catch {
    Write-Host "Failed to find the service principal. Ensure the Application ID is correct." -ForegroundColor Red
    exit
}

# Define the Microsoft Graph API `Sites.Selected` permission
Write-Host "Assigning the `Sites.Selected` Microsoft Graph API permission..." -ForegroundColor Yellow

# Get the Microsoft Graph API application
$graphApp = Get-AzADServicePrincipal -DisplayName "Microsoft Graph"

# Grant the `Sites.Selected` permission
Write-Host "Assigning the `Sites.Selected` permission to the Application ID: $appId..." -ForegroundColor Yellow
try {
    Add-AzADAppPermission -ObjectId $sp.Id `
        -ApiId $graphApp.AppId `
        -PermissionId "883ea226-0bf2-4a8f-9f9d-92c9162a727d" `
        -Type Role
    Write-Host "`Sites.Selected` permission granted successfully!" -ForegroundColor Green
} catch {
    Write-Host "Failed to grant the `Sites.Selected` permission. Error: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Prompt for admin consent
$consent = Read-Host -Prompt "You must grant admin consent. Do you want to grant admin consent now? (y/n)"
if ($consent -eq 'y') {
    try {
        Write-Host "Granting admin consent for the `Sites.Selected` permission..." -ForegroundColor Yellow
        Start-Process -FilePath "https://login.microsoftonline.com/$($tenantId)/adminconsent?client_id=$($sp.AppId)"
        Write-Host "Admin consent process started. Please complete it in the browser." -ForegroundColor Green
    } catch {
        Write-Host "Failed to initiate admin consent. Error: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "Skipping admin consent." -ForegroundColor Yellow
}

# Disconnect from Azure
Disconnect-AzAccount | Out-Null

Write-Host "Script execution complete!" -ForegroundColor Cyan