#Requires -Version 5.1
#Requires -Module Microsoft.Graph.Authentication
#Requires -Module Microsoft.Graph.Applications
#Requires -Module Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Updates Microsoft Graph permissions for Device Management endpoints from old to new permissions.

.DESCRIPTION
    This script identifies and updates applications, service principals, and managed identities that have
    the old DeviceManagementConfiguration permissions and adds the new DeviceManagementScripts
    permissions alongside them as required by Microsoft Intune changes effective July 31, 2025.
    
    The old permissions will continue to be used for other endpoints (configuration profiles, etc.)
    and will NOT be removed. The new permissions are required for specific device management script endpoints.

    Existing permissions (will continue to work for configuration profiles and other endpoints):
    - DeviceManagementConfiguration.Read.All
    - DeviceManagementConfiguration.ReadWrite.All
    
    New permissions being added (required for device management script endpoints after July 31, 2025):
    - DeviceManagementScripts.Read.All
    - DeviceManagementScripts.ReadWrite.All

.PARAMETER Whatif
    When specified, shows what changes would be made without actually making them.

.PARAMETER Force
    When specified, applies changes without prompting for confirmation.

.PARAMETER LogPath
    Path to write detailed logs. Defaults to current directory.

.EXAMPLE
    .\Update-DeviceManagement-Permissions.ps1 -Whatif
    Shows what changes would be made without making them.

.EXAMPLE
    .\Update-DeviceManagement-Permissions.ps1 -Force
    Updates all identified applications and service principals without prompting.

.NOTES
    Author: Generated Script
    Version: 1.0
    Date: 2025-01-17
    
    This script requires the following Microsoft Graph permissions:
    - Application.ReadWrite.All
    - Directory.ReadWrite.All
    - AppRoleAssignment.ReadWrite.All
    
    References:
    - https://learn.microsoft.com/en-us/intune/intune-service/fundamentals/in-development#device-management
    - https://learn.microsoft.com/en-us/intune/intune-service/developer/graph-apis-used-by-intune-device-configuration-windows
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [switch]$Whatif,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\DeviceManagement-Permissions-Update.log"
)

# Set up error handling
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# Define permission mappings
$PermissionMappings = @{
    "DeviceManagementConfiguration.Read.All" = "DeviceManagementScripts.Read.All"
    "DeviceManagementConfiguration.ReadWrite.All" = "DeviceManagementScripts.ReadWrite.All"
}

# Define the Microsoft Graph service principal ID
$MicrosoftGraphAppId = "00000003-0000-0000-c000-000000000000"

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    Write-Host $logEntry -ForegroundColor $(
        switch ($Level) {
            "ERROR" { "Red" }
            "WARNING" { "Yellow" }
            "SUCCESS" { "Green" }
            default { "White" }
        }
    )
    
    Add-Content -Path $LogPath -Value $logEntry
}

# Function to get Microsoft Graph Service Principal
function Get-MicrosoftGraphServicePrincipal {
    try {
        Write-Log "Getting Microsoft Graph service principal..."
        $graphSP = Get-MgServicePrincipal -Filter "appId eq '$MicrosoftGraphAppId'"
        if (-not $graphSP) {
            throw "Could not find Microsoft Graph service principal"
        }
        return $graphSP
    }
    catch {
        Write-Log "Error getting Microsoft Graph service principal: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# Function to get permission ID from service principal
function Get-PermissionId {
    param(
        [object]$ServicePrincipal,
        [string]$Permission
    )
    
    $appRole = $ServicePrincipal.AppRoles | Where-Object { $_.Value -eq $Permission }
    if ($appRole) {
        return $appRole.Id
    }
    else {
        Write-Log "Permission '$Permission' not found in service principal app roles" -Level "WARNING"
        return $null
    }
}

# Function to check if application has old permissions
function Test-ApplicationHasOldPermissions {
    param(
        [object]$Application
    )
    
    $hasOldPermissions = $false
    
    if ($Application.RequiredResourceAccess) {
        foreach ($resourceAccess in $Application.RequiredResourceAccess) {
            if ($resourceAccess.ResourceAppId -eq $MicrosoftGraphAppId) {
                foreach ($permission in $resourceAccess.ResourceAccess) {
                    $permissionValue = $graphSP.AppRoles | Where-Object { $_.Id -eq $permission.Id } | Select-Object -ExpandProperty Value
                    if ($permissionValue -in $PermissionMappings.Keys) {
                        $hasOldPermissions = $true
                        break
                    }
                }
            }
            if ($hasOldPermissions) { break }
        }
    }
    
    return $hasOldPermissions
}

# Function to check if service principal has old permission grants
function Test-ServicePrincipalHasOldPermissions {
    param(
        [object]$ServicePrincipal,
        [object]$GraphServicePrincipal
    )
    
    try {
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id -Filter "resourceId eq '$($GraphServicePrincipal.Id)'"
        
        foreach ($assignment in $appRoleAssignments) {
            $permissionValue = $GraphServicePrincipal.AppRoles | Where-Object { $_.Id -eq $assignment.AppRoleId } | Select-Object -ExpandProperty Value
            if ($permissionValue -in $PermissionMappings.Keys) {
                return $true
            }
        }
    }
    catch {
        Write-Log "Error checking service principal permissions for $($ServicePrincipal.DisplayName): $($_.Exception.Message)" -Level "WARNING"
    }
    
    return $false
}

# Function to update application permissions
function Update-ApplicationPermissions {
    param(
        [object]$Application,
        [object]$GraphServicePrincipal
    )
    
    try {
        Write-Log "Updating application: $($Application.DisplayName) (ID: $($Application.Id))"
        
        $needsUpdate = $false
        $updatedResourceAccess = @()
        
        # Process each resource access
        foreach ($resourceAccess in $Application.RequiredResourceAccess) {
            if ($resourceAccess.ResourceAppId -eq $MicrosoftGraphAppId) {
                $newResourceAccess = @()
                
                foreach ($permission in $resourceAccess.ResourceAccess) {
                    $permissionValue = $GraphServicePrincipal.AppRoles | Where-Object { $_.Id -eq $permission.Id } | Select-Object -ExpandProperty Value
                    
                    # Always keep existing permission
                    $newResourceAccess += $permission
                    
                    if ($permissionValue -in $PermissionMappings.Keys) {
                        # Add new permission alongside the old one
                        $newPermissionValue = $PermissionMappings[$permissionValue]
                        $newPermissionId = Get-PermissionId -ServicePrincipal $GraphServicePrincipal -Permission $newPermissionValue
                        
                        if ($newPermissionId) {
                            # Check if new permission is not already present
                            $alreadyExists = $resourceAccess.ResourceAccess | Where-Object { $_.Id -eq $newPermissionId }
                            if (-not $alreadyExists) {
                                Write-Log "  Adding $newPermissionValue alongside existing $permissionValue"
                                $newResourceAccess += @{
                                    Id = $newPermissionId
                                    Type = $permission.Type
                                }
                                $needsUpdate = $true
                            }
                            else {
                                Write-Log "  $newPermissionValue already exists, skipping"
                            }
                        }
                        else {
                            Write-Log "  Could not find new permission ID for $newPermissionValue" -Level "WARNING"
                        }
                    }
                }
                
                $updatedResourceAccess += @{
                    ResourceAppId = $resourceAccess.ResourceAppId
                    ResourceAccess = $newResourceAccess
                }
            }
            else {
                # Keep other resource access unchanged
                $updatedResourceAccess += $resourceAccess
            }
        }
        
        if ($needsUpdate) {
            if ($PSCmdlet.ShouldProcess($Application.DisplayName, "Update application permissions")) {
                if (-not $Whatif) {
                    Update-MgApplication -ApplicationId $Application.Id -RequiredResourceAccess $updatedResourceAccess
                    Write-Log "  Successfully updated application permissions" -Level "SUCCESS"
                }
                else {
                    Write-Log "  [WHATIF] Would update application permissions" -Level "SUCCESS"
                }
            }
        }
        else {
            Write-Log "  No permission updates needed for this application"
        }
    }
    catch {
        Write-Log "Error updating application $($Application.DisplayName): $($_.Exception.Message)" -Level "ERROR"
    }
}

# Function to update service principal permissions
function Update-ServicePrincipalPermissions {
    param(
        [object]$ServicePrincipal,
        [object]$GraphServicePrincipal
    )
    
    try {
        Write-Log "Updating service principal: $($ServicePrincipal.DisplayName) (ID: $($ServicePrincipal.Id))"
        
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id -Filter "resourceId eq '$($GraphServicePrincipal.Id)'"
        $needsUpdate = $false
        
        foreach ($assignment in $appRoleAssignments) {
            $permissionValue = $GraphServicePrincipal.AppRoles | Where-Object { $_.Id -eq $assignment.AppRoleId } | Select-Object -ExpandProperty Value
            
            if ($permissionValue -in $PermissionMappings.Keys) {
                $newPermissionValue = $PermissionMappings[$permissionValue]
                $newPermissionId = Get-PermissionId -ServicePrincipal $GraphServicePrincipal -Permission $newPermissionValue
                
                if ($newPermissionId) {
                    # Check if new permission is not already assigned
                    $existingNewPermission = $appRoleAssignments | Where-Object { $_.AppRoleId -eq $newPermissionId }
                    if (-not $existingNewPermission) {
                        Write-Log "  Adding $newPermissionValue alongside existing $permissionValue"
                        
                        if ($PSCmdlet.ShouldProcess($ServicePrincipal.DisplayName, "Add new service principal permission")) {
                            if (-not $Whatif) {
                                # Add new permission (keep old one)
                                $newAssignment = @{
                                    PrincipalId = $ServicePrincipal.Id
                                    ResourceId = $GraphServicePrincipal.Id
                                    AppRoleId = $newPermissionId
                                }
                                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id -BodyParameter $newAssignment
                                
                                Write-Log "  Successfully added service principal permission" -Level "SUCCESS"
                            }
                            else {
                                Write-Log "  [WHATIF] Would add service principal permission" -Level "SUCCESS"
                            }
                        }
                        $needsUpdate = $true
                    }
                    else {
                        Write-Log "  $newPermissionValue already exists, skipping"
                    }
                }
                else {
                    Write-Log "  Could not find new permission ID for $newPermissionValue" -Level "WARNING"
                }
            }
        }
        
        if (-not $needsUpdate) {
            Write-Log "  No permission updates needed for this service principal"
        }
    }
    catch {
        Write-Log "Error updating service principal $($ServicePrincipal.DisplayName): $($_.Exception.Message)" -Level "ERROR"
    }
}

# Main execution
function Main {
    try {
        Write-Log "Starting Device Management Permissions Update Script"
        Write-Log "Log file: $LogPath"
        
        # Check if running in WhatIf mode
        if ($Whatif) {
            Write-Log "Running in WHATIF mode - no changes will be made" -Level "WARNING"
        }
        
        # Connect to Microsoft Graph
        Write-Log "Connecting to Microsoft Graph..."
        try {
            $context = Get-MgContext
            if (-not $context) {
                Write-Log "Not connected to Microsoft Graph. Attempting to connect..."
                Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"
            }
            else {
                Write-Log "Already connected to Microsoft Graph as $($context.Account)"
            }
        }
        catch {
            Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level "ERROR"
            throw
        }
        
        # Get Microsoft Graph service principal
        $graphSP = Get-MicrosoftGraphServicePrincipal
        Write-Log "Found Microsoft Graph service principal: $($graphSP.DisplayName)"
        
        # Find applications with old permissions
        Write-Log "Scanning applications for old permissions..."
        $applications = Get-MgApplication -All
        $appsToUpdate = @()
        
        foreach ($app in $applications) {
            if (Test-ApplicationHasOldPermissions -Application $app) {
                $appsToUpdate += $app
                Write-Log "Found application with old permissions: $($app.DisplayName)"
            }
        }
        
        # Find service principals with old permissions
        Write-Log "Scanning service principals for old permissions..."
        $servicePrincipals = Get-MgServicePrincipal -All
        $spsToUpdate = @()
        
        foreach ($sp in $servicePrincipals) {
            if (Test-ServicePrincipalHasOldPermissions -ServicePrincipal $sp -GraphServicePrincipal $graphSP) {
                $spsToUpdate += $sp
                Write-Log "Found service principal with old permissions: $($sp.DisplayName)"
            }
        }
        
        # Summary
        Write-Log "Summary of findings:"
        Write-Log "  Applications to update: $($appsToUpdate.Count)"
        Write-Log "  Service principals to update: $($spsToUpdate.Count)"
        
        if ($appsToUpdate.Count -eq 0 -and $spsToUpdate.Count -eq 0) {
            Write-Log "No applications or service principals found with old permissions. No updates needed." -Level "SUCCESS"
            return
        }
        
        # Confirmation prompt
        if (-not $Force -and -not $Whatif) {
            $confirmation = Read-Host "Do you want to proceed with updating permissions? (y/N)"
            if ($confirmation -notmatch "^[Yy]$") {
                Write-Log "Operation cancelled by user"
                return
            }
        }
        
        # Update applications
        if ($appsToUpdate.Count -gt 0) {
            Write-Log "Updating applications..."
            foreach ($app in $appsToUpdate) {
                Update-ApplicationPermissions -Application $app -GraphServicePrincipal $graphSP
            }
        }
        
        # Update service principals
        if ($spsToUpdate.Count -gt 0) {
            Write-Log "Updating service principals..."
            foreach ($sp in $spsToUpdate) {
                Update-ServicePrincipalPermissions -ServicePrincipal $sp -GraphServicePrincipal $graphSP
            }
        }
        
        Write-Log "Device Management Permissions Update completed successfully!" -Level "SUCCESS"
        Write-Log "Important: Please test your applications to ensure they continue to work with the new permissions."
        Write-Log "Note: The old permissions will continue to work for configuration profiles and other endpoints."
        Write-Log "The new permissions are specifically required for device management script endpoints after July 31, 2025."
        
    }
    catch {
        Write-Log "Script execution failed: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Stack trace: $($_.Exception.StackTrace)" -Level "ERROR"
        throw
    }
}

# Execute main function
if ($MyInvocation.InvocationName -ne '.') {
    Main
}
