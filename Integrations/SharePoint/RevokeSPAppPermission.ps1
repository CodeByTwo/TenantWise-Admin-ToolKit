<#
.SYNOPSIS
Connects to SharePoint Online using the PnP PowerShell module and revokes application permission to site.

.DESCRIPTION
Checks if the PnP.PowerShell module is installed, prompts for credentials and connects to SharePoint Online and revokes application permission to site.
#>

# Function to check and install the PnP PowerShell module
function Ensure-PnPModule {
    Write-Host "Checking if the PnP.PowerShell module is installed..." -ForegroundColor Yellow
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
        Write-Host "PnP.PowerShell module is not installed." -ForegroundColor Yellow
        $installPnP = Read-Host -Prompt "Would you like to install the PnP.PowerShell module? (y/n)"
        if ($installPnP -eq 'y') {
            Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
            try {
                Install-Module -Name "PnP.PowerShell" -AllowClobber -Scope CurrentUser -Force
                Write-Host "PnP.PowerShell module installed successfully." -ForegroundColor Green
            } catch {
                Write-Host "Failed to install PnP.PowerShell module. Please install it manually and try again." -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "You chose not to install the PnP.PowerShell module. The script cannot proceed." -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "PnP.PowerShell module is already installed." -ForegroundColor Green
    }
    Import-Module -Name "PnP.PowerShell" -Force
}

# Ensure the PnP.PowerShell module is installed
Ensure-PnPModule

# Prompt for the SharePoint site URL
$siteUrl = Read-Host -Prompt "Enter the SharePoint site URL (e.g., https://yourtenant.sharepoint.com/sites/yoursite)"

# Validate the URL format
if ($siteUrl -notmatch '^https:\/\/[a-zA-Z0-9\-]+\.sharepoint\.com\/sites\/[a-zA-Z0-9\-]+$') {
    Write-Host "Invalid SharePoint site URL format. Please ensure it follows the pattern: https://<tenant>.sharepoint.com/sites/<site>." -ForegroundColor Red
    exit
}

# Prompt for the Client ID of the registered app
$clientId = Read-Host -Prompt "Enter the Client ID of the registered PnP app. If you have not registered see https://pnp.github.io/powershell/articles/registerapplication.html"

# Validate the Client ID format
if ($clientId -notmatch '^[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}$') {
    Write-Host "Invalid Client ID format. Please ensure it is a valid GUID." -ForegroundColor Red
    exit
}

# Prompt for Application ID
$appId = Read-Host -Prompt "Enter the Application ID of the service principal requesting access to the site"

try {
    #Get the Permission ID for the app using App Id
    $PermissionId = Get-PnPAzureADAppSitePermission -AppIdentity $appId
}
catch {
    Write-Host "Failed to get the Permission ID for the app. Please check the Application ID and try again." -ForegroundColor Red
}

try {
    #Revoke the permission
    Revoke-PnPAzureADAppSitePermission -PermissionId $PermissionId.Id
}
catch {
    Write-Host "Failed to revoke the permission for the app. Please check the Permission ID and try again." -ForegroundColor Red
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline | Out-Null

Write-Host "Script execution complete!" -ForegroundColor Cyan