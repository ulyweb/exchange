#Requires -RunAsAdministrator

<#
.SYNOPSIS
    This script automates the prerequisites for connecting to Exchange Online and provides a menu for daily tasks.

.DESCRIPTION
    The script performs the following actions in order:
    1.  Checks if it is running with administrator privileges.
        - If not, it will stop and instruct the user to run it as an admin.
    2.  Verifies the Execution Policy for a Windows environment.
        - If a scope is not set to 'RemoteSigned', it will be set.
    3.  Checks if the 'PowerShellGet' and 'ExchangeOnlineManagement' modules are installed.
        - If 'ExchangeOnlineManagement' is missing, it will install it for the current user.
    4.  Updates the 'ExchangeOnlineManagement' module to its latest version.
    5.  Finally, it attempts to connect to Exchange Online using 'Connect-ExchangeOnline'.
    6.  After a successful connection, it presents an interactive menu for tasks like viewing mailbox calendars.
#>

# =========================================================================
# Step 1: Check for Administrator Privileges
# =========================================================================

# This #Requires statement at the top of the script is the most reliable way
# to ensure the script is run with administrator rights. If the user doesn't
# run it as an admin, PowerShell will automatically stop the script
# and display an error with instructions. This is more robust than a manual check.

Write-Host "✅ Checking for administrator privileges..." -ForegroundColor Green
# If the script gets past the #Requires statement, it means it's running as an admin.
Write-Host "✅ Running as Administrator." -ForegroundColor Green
Write-Host "----------------------------------------------------"

# =========================================================================
# Step 2: Set Execution Policy to RemoteSigned (Windows Only)
# =========================================================================

# The Execution Policy is a Windows-specific security feature.
# It is not supported on non-Windows platforms like Linux, used in Codespaces.
if ($PSVersionTable.PSVersion.Major -lt 6 -or $PSVersionTable.PSVersion.Platform -eq "Windows") {
    Write-Host "➡️ Checking and setting Execution Policy to 'RemoteSigned'..." -ForegroundColor Yellow

    # This function checks the policy for a given scope and sets it if needed.
    function Set-RequiredExecutionPolicy {
        param(
            [string]$Scope
        )

        $currentPolicy = Get-ExecutionPolicy -Scope $Scope
        if ($currentPolicy -ne 'RemoteSigned') {
            Write-Host "  - The '$Scope' scope is currently set to '$currentPolicy'. Setting to 'RemoteSigned'..."
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope $Scope -Force
            Write-Host "  - The '$Scope' scope is now set to 'RemoteSigned'." -ForegroundColor Green
        } else {
            Write-Host "  - The '$Scope' scope is already 'RemoteSigned'." -ForegroundColor Green
        }
    }

    # Set the policy for both scopes as requested.
    Set-RequiredExecutionPolicy -Scope CurrentUser
    Set-RequiredExecutionPolicy -Scope LocalMachine
} else {
    Write-Host "ℹ️ Skipping Execution Policy check. This is not supported on this platform." -ForegroundColor Yellow
}

Write-Host "----------------------------------------------------"

# =========================================================================
# Step 3: Check, Install, and Update the 'ExchangeOnlineManagement' Module
# =========================================================================

# Check if the PowerShellGet module is installed, as it's required to install other modules.
Write-Host "➡️ Checking for the 'PowerShellGet' module..." -ForegroundColor Yellow
if (-not (Get-Module -ListAvailable -Name PowerShellGet)) {
    Write-Host "  - 'PowerShellGet' module not found. Installing now..." -ForegroundColor Red
    # Installing with CurrentUser scope for Codespaces/non-admin environments.
    Install-Module -Name PowerShellGet -Force -AllowClobber -Scope CurrentUser -Confirm:$false
    Write-Host "✅ 'PowerShellGet' installed successfully." -ForegroundColor Green
} else {
    Write-Host "✅ 'PowerShellGet' module is installed." -ForegroundColor Green
}
Write-Host "----------------------------------------------------"

# Now check for the ExchangeOnlineManagement module.
Write-Host "➡️ Checking for the 'ExchangeOnlineManagement' module..." -ForegroundColor Yellow
$moduleName = "ExchangeOnlineManagement"
if (-not (Get-Module -Name $moduleName -ListAvailable)) {
    Write-Host "  - '$moduleName' module is not installed. Installing it now..." -ForegroundColor Red
    # Installing with CurrentUser scope to avoid permission issues in Codespaces/non-admin environments.
    Install-Module -Name $moduleName -Force -Scope CurrentUser -Confirm:$false
    Write-Host "✅ '$moduleName' module installed successfully." -ForegroundColor Green
} else {
    Write-Host "✅ '$moduleName' module is already installed." -ForegroundColor Green

    Write-Host "➡️ Checking for updates for '$moduleName' module..." -ForegroundColor Yellow
    try {
        # Check for the latest version from the gallery and compare it to the installed version.
        $installedVersion = (Get-InstalledModule -Name $moduleName).Version
        $latestVersion = (Find-Module -Name $moduleName -ErrorAction SilentlyContinue).Version

        if ($latestVersion -and ($latestVersion -gt $installedVersion)) {
            Write-Host "  - A newer version ($latestVersion) is available. Updating now..." -ForegroundColor Cyan
            # Update with CurrentUser scope for compatibility.
            Update-Module -Name $moduleName -Force -Scope CurrentUser -Confirm:$false -ErrorAction Stop
            Write-Host "✅ '$moduleName' module is updated to the latest version." -ForegroundColor Green
        } else {
            Write-Host "✅ '$moduleName' module is already up-to-date." -ForegroundColor Green
        }
    } catch {
        Write-Host "❌ An error occurred during the update process: $_" -ForegroundColor Red
        Write-Host "  - Continuing with the existing module version." -ForegroundColor Yellow
    }
}

Write-Host "----------------------------------------------------"

# =========================================================================
# Step 4: Connect to Exchange Online (only if a session is not active)
# =========================================================================

Write-Host "➡️ Attempting to connect to Exchange Online..." -ForegroundColor Yellow
try {
    # Ensure the module is loaded first to allow the Get-Mailbox command to work.
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
    
    # Test for an active, working connection by trying a simple command.
    Get-Mailbox -ErrorAction SilentlyContinue | Out-Null
    
    # If the above command succeeds, we know we have an active connection.
    Write-Host "✅ An active Exchange Online session already exists. Skipping connection." -ForegroundColor Green

} catch {
    # If the test command fails, we need to create a new session.
    Write-Host "  - No active Exchange Online session found. Connecting now..." -ForegroundColor Yellow
    
    # Determine which authentication method to use based on the platform.
    if ($PSVersionTable.PSVersion.Major -lt 6 -or $PSVersionTable.PSVersion.Platform -eq "Windows") {
        # Use standard authentication for Windows environments.
        Write-Host "  - Using standard web-based authentication..." -ForegroundColor Yellow
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -ShowProgress:$true
    } else {
        # Use device authentication for non-Windows (headless) environments like Codespaces.
        Write-Host "  - Using device authentication (no browser required)..." -ForegroundColor Yellow
        Write-Host "  - Follow the instructions below to sign in from your local browser." -ForegroundColor Cyan
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -ShowProgress:$true -UseDeviceAuthentication
    }

    Write-Host "✅ Successfully connected to Exchange Online." -ForegroundColor Green
}

# =========================================================================
# Step 5: Daily Task Menu
# =========================================================================

# New function to display a mailbox's calendar permissions
function Display-MailboxCalendarPermissions {
    Write-Host "--- Display Mailbox Calendar Permissions ---" -ForegroundColor Cyan
    $mailbox = Read-Host "Enter the mailbox UPN or email address"
    
    # We need to specify the calendar folder to get its permissions
    $calendarFolder = "$mailbox" + ":\Calendar"

    Write-Host "Checking calendar permissions for '$mailbox'..." -ForegroundColor Cyan
    
    try {
        $calendarPerms = Get-MailboxFolderPermission -Identity $calendarFolder -ErrorAction Stop
        Write-Host "✅ Calendar Permissions for '$mailbox' found:" -ForegroundColor Green
        $calendarPerms | Format-List
    } catch {
        Write-Host "❌ An error occurred. Mailbox '$mailbox' not found, or you do not have permissions to view its calendar." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the main menu." | Out-Null
}

# New function to display all mailbox permissions
function Display-MailboxPermissions {
    Write-Host "--- Display All Mailbox Permissions ---" -ForegroundColor Cyan
    $mailbox = Read-Host "Enter the mailbox UPN or email address"
    
    Write-Host "Checking permissions for '$mailbox'..." -ForegroundColor Cyan

    try {
        $mailboxPerms = Get-MailboxPermission -Identity $mailbox -ErrorAction Stop
        Write-Host "✅ Mailbox Permissions for '$mailbox' found:" -ForegroundColor Green
        $mailboxPerms | Format-List
    } catch {
        Write-Host "❌ An error occurred. Mailbox '$mailbox' not found, or you do not have permissions to view its permissions." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the main menu." | Out-Null
}

# New function to display calendar folder statistics
function Display-CalendarFolderStatistics {
    Write-Host "--- Display Calendar Folder Statistics ---" -ForegroundColor Cyan
    $mailbox = Read-Host "Enter the mailbox UPN or email address"
    
    Write-Host "Checking calendar statistics for '$mailbox'..." -ForegroundColor Cyan
    
    try {
        Get-MailboxFolderStatistics -Identity $mailbox -ErrorAction Stop | Where-Object {$_.FolderType -eq "Calendar"} | Select-Object Name, FolderSize, ItemsInFolder
    } catch {
        Write-Host "❌ An error occurred. Mailbox '$mailbox' not found, or you do not have permissions to view its statistics." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the main menu." | Out-Null
}

# New function to add calendar permissions
function Add-CalendarPermissionsWithDelegate {
    Write-Host "--- Add Calendar Permissions (with Delegate) ---" -ForegroundColor Cyan
    $mailbox = Read-Host "Enter the mailbox UPN or email address to modify"
    $user = Read-Host "Enter the UPN or email address of the user to grant access"

    try {
        Add-MailboxFolderPermission -Identity "$mailbox`:\Calendar" -User $user -AccessRights Editor -SharingPermissionFlags Delegate -Confirm:$false -ErrorAction Stop
        Write-Host "✅ Successfully added Editor permissions (with Delegate) for '$user' on '$mailbox's calendar." -ForegroundColor Green
    } catch {
        Write-Host "❌ An error occurred. Please check the mailbox and user email addresses and try again." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the main menu." | Out-Null
}

# New function to add calendar permissions with Editor only
function Add-CalendarPermissionsEditorOnly {
    Write-Host "--- Add Calendar Permissions (Editor Only) ---" -ForegroundColor Cyan
    $mailbox = Read-Host "Enter the mailbox UPN or email address to modify"
    $user = Read-Host "Enter the UPN or email address of the user to grant access"

    try {
        Add-MailboxFolderPermission -Identity "$mailbox`:\Calendar" -User $user -AccessRights Editor -Confirm:$false -ErrorAction Stop
        Write-Host "✅ Successfully added Editor permissions for '$user' on '$mailbox's calendar." -ForegroundColor Green
    } catch {
        Write-Host "❌ An error occurred. Please check the mailbox and user email addresses and try again." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the main menu." | Out-Null
}

# New function to remove calendar permissions
function Remove-CalendarPermissions {
    Write-Host "--- Remove Calendar Permissions ---" -ForegroundColor Cyan
    $mailbox = Read-Host "Enter the mailbox UPN or email address to remove permissions from"
    $user = Read-Host "Enter the UPN or email address of the user to remove access for"
    
    try {
        Remove-MailboxFolderPermission -Identity "$mailbox`:\Calendar" -User $user -Confirm:$false -ErrorAction Stop
        Write-Host "✅ Successfully removed permissions for '$user' from '$mailbox's calendar." -ForegroundColor Green
    } catch {
        Write-Host "❌ An error occurred. The user's permissions might not exist or you do not have the required permissions." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the main menu." | Out-Null
}

# New function for detailed calendar folder statistics
function Display-DetailedCalendarStatistics {
    Write-Host "--- Display Detailed Calendar Folder Statistics ---" -ForegroundColor Cyan
    $mailbox = Read-Host "Enter the mailbox UPN or email address"

    try {
        Get-MailboxFolderStatistics -Identity $mailbox -FolderScope Calendar -ErrorAction Stop | 
        Select-Object Name, FolderPath, FolderSize, ItemsInFolder | 
        Format-Table -AutoSize
    } catch {
        Write-Host "❌ An error occurred. Mailbox '$mailbox' not found, or you do not have permissions to view its statistics." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the main menu." | Out-Null
}

# New function to add full access to a mailbox
function Add-FullMailboxAccess {
    Write-Host "--- Add Full Mailbox Access ---" -ForegroundColor Cyan
    $mailbox = Read-Host "Enter the mailbox UPN or email address to modify"
    $user = Read-Host "Enter the UPN or email address of the user to grant access"

    try {
        Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false -Confirm:$false -ErrorAction Stop
        Write-Host "✅ Successfully added Full Access for '$user' on '$mailbox's mailbox." -ForegroundColor Green
        Write-Host "Note: Auto-mapping is disabled. The user will need to manually add the mailbox." -ForegroundColor Yellow
    } catch {
        Write-Host "❌ An error occurred. Please check the mailbox and user email addresses and try again." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the main menu." | Out-Null
}

# New function to remove full access from a mailbox
function Remove-FullMailboxAccess {
    Write-Host "--- Remove Full Mailbox Access ---" -ForegroundColor Cyan
    $mailbox = Read-Host "Enter the mailbox UPN or email address to remove permissions from"
    $user = Read-Host "Enter the UPN or email address of the user to remove access for"

    try {
        Remove-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
        Write-Host "✅ Successfully removed Full Access for '$user' from '$mailbox's mailbox." -ForegroundColor Green
    } catch {
        Write-Host "❌ An error occurred. The user's permissions might not exist or you do not have the required permissions." -ForegroundColor Red
    }
    Read-Host "Press Enter to return to the main menu." | Out-Null
}

$choice = ""
do {
    # Display the menu
    Write-Host "`n====================================================" -ForegroundColor Cyan
    Write-Host "      Exchange Online Daily Tasks Menu" -ForegroundColor Cyan
    Write-Host "====================================================" -ForegroundColor Cyan
    Write-Host "  1) Add Calendar Permissions (with Delegate)" -ForegroundColor Cyan
    Write-Host "  2) Add Calendar Permissions (Editor Only)" -ForegroundColor Cyan
    Write-Host "  3) Remove Calendar Permissions" -ForegroundColor Cyan
    Write-Host "  4) Add Full Mailbox Access" -ForegroundColor Cyan
    Write-Host "  5) Remove Full Mailbox Access" -ForegroundColor Cyan
    Write-Host "  6) Display Calendar Folder Statistics (Summary)" -ForegroundColor Cyan
    Write-Host "  7) Display Calendar Folder Statistics (Detailed)" -ForegroundColor Cyan
    Write-Host "  8) Display Calendar Mailbox Permissions" -ForegroundColor Cyan
    Write-Host "  9) Display Mailbox Permissions (Full)" -ForegroundColor Cyan
    Write-Host "  X) Exit" -ForegroundColor Cyan
    Write-Host "  D) Disconnect" -ForegroundColor Cyan
    Write-Host "----------------------------------------------------" -ForegroundColor Cyan
    $choice = Read-Host "Enter your choice (1-9, D, X)"

    switch ($choice) {
        "1" {
            Add-CalendarPermissionsWithDelegate
        }
        "2" {
            Add-CalendarPermissionsEditorOnly
        }
        "3" {
            Remove-CalendarPermissions
        }
        "4" {
            Add-FullMailboxAccess
        }
        "5" {
            Remove-FullMailboxAccess
        }
        "6" {
            Display-CalendarFolderStatistics
        }
        "7" {
            Display-DetailedCalendarStatistics
        }
        "8" {
            Display-MailboxCalendarPermissions
        }
        "9" {
            Display-MailboxPermissions
        }
        "D" {
            Write-Host "➡️ Disconnecting from Exchange Online..." -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Host "✅ Successfully disconnected." -ForegroundColor Green
            $choice = "X" # Set choice to "X" to exit the loop
        }
        "X" {
            Write-Host "Exiting script. Goodbye!" -ForegroundColor Yellow
        }
        default {
            Write-Host "❌ Invalid choice. Please try again." -ForegroundColor Red
        }
    }
} while ($choice -ne "X")
