# TenantWise Admin ToolKit 🚀

Welcome to the TenantWise Admin ToolKit repository! This toolkit helps you set up TenantWise integrations with ease. Future integration setup scripts may be added.

## Microsoft SharePoint 📑

### Dependencies

The PowerShell scripts for Microsoft SharePoint integration require the following dependencies:

- **Az PowerShell**: Used for Entra ID management. [Official Docs](https://learn.microsoft.com/en-us/powershell/azure/new-azureps-module-az)
- **PnP PowerShell**: Used for SharePoint Online management. [Official Docs](https://pnp.github.io/powershell/index.html)

Make sure to install these modules before running the scripts.

### Order of events

Run the included scripts in the following order:

**Mandatory**

1. **CreateServicePrincipal.ps1**: Creates a Service Principal/App Registration with a name captured from prompt. Generates secret for authentication.
2. **GrantSPGraphPermission.ps1**: Grants `Sites.Selected` Graph API permissions to Service Principal via App ID provided. Also handles admin consent for API permissions.
3. **GrantSPAppPermission.ps1**: Grants provided App ID access to SharePoint site URL provided.

**Optional**

1. **RevokeSPAppPermission.ps1**: Revokes provided App ID access to SharePoint site URL provided.