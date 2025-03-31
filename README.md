> [!NOTE]
> # Exchange Online

> [!IMPORTANT]
>>

## **1. Mailbox Permissions (`MailboxPermission`)**
**Purpose:**
- Grants or removes **Full Access** to a mailbox.
- Commonly used for assistants managing another user's emails or shared mailboxes.

**Commands:**
```powershell
# View existing permissions
Get-MailboxPermission -Identity "UserMailbox@domain.com" | Select-Object -Property * | ft

# Grant Full Access (without auto-mapping)
Add-MailboxPermission -Identity "UserMailbox@domain.com" -User "SupportUser@domain.com" -AccessRights FullAccess -AutoMapping $false -Confirm:$false

# Remove Full Access
Remove-MailboxPermission -Identity "UserMailbox@domain.com" -User "SupportUser@domain.com" -AccessRights FullAccess -Confirm:$false
```

---

## **2. Calendar Permissions (`MailboxFolderPermission`)**
**Purpose:**
- Grants access to a specific mailbox folder (e.g., Calendar).
- Used to allow users to **view or manage** someone else’s calendar.

**Commands:**
```powershell
# Grant Reviewer (Read-Only) access to the calendar
Add-MailboxFolderPermission -Identity "UserMailbox@domain.com:\Calendar" -User "Employee@domain.com" -AccessRights Reviewer

# Grant Editor (Modify & Create Items) access to the calendar
Add-MailboxFolderPermission -Identity "UserMailbox@domain.com:\Calendar" -User "Manager@domain.com" -AccessRights Editor

# View existing Calendar permissions
Get-MailboxFolderPermission -Identity "UserMailbox@domain.com:\Calendar"
```

---

## **3. Managing Distribution List (DL) Membership**
**Purpose:**
- Adds or removes users from distribution lists (DLs).
- Ensures proper email group communication.

**Commands:**
```powershell
# Add a user to a DL
Add-DistributionGroupMember -Identity "SalesTeam@domain.com" -Member "NewUser@domain.com"

# Remove a user from a DL
Remove-DistributionGroupMember -Identity "SalesTeam@domain.com" -Member "UserLeaving@domain.com"

# Check DL members
Get-DistributionGroupMember -Identity "SalesTeam@domain.com"
```

---

## **4. Assigning & Checking SendAs Permissions (`RecipientPermission`)**
**Purpose:**
- Allows a user to send emails as another user or shared mailbox.
- Used for roles where emails must appear to come from a team instead of individuals.

**Commands:**
```powershell
# Grant SendAs permission
Add-RecipientPermission -Identity "SharedMailbox@domain.com" -Trustee "User@domain.com" -AccessRights SendAs -Confirm:$false

# Remove SendAs permission
Remove-RecipientPermission -Identity "SharedMailbox@domain.com" -Trustee "User@domain.com" -AccessRights SendAs -Confirm:$false

# Check who has SendAs permission
Get-RecipientPermission -Identity "SharedMailbox@domain.com" | Where-Object {$_.AccessRights -eq "SendAs"}
```

---

## **5. Checking Mailbox Size & Usage**
**Purpose:**
- Helps identify large mailboxes that may need archiving or quota increases.
- Useful for troubleshooting **storage-related** issues.

**Commands:**
```powershell
# Check mailbox size for a specific user
Get-MailboxStatistics -Identity "UserMailbox@domain.com" | Select DisplayName, TotalItemSize, ItemCount

# List all mailboxes sorted by size
Get-MailboxStatistics -Database "MailboxDatabase" | Sort-Object TotalItemSize -Descending | Select DisplayName, TotalItemSize, ItemCount
```

---

## **6. Resetting Out-of-Office Messages**
**Purpose:**
- Enables IT to check and update a user’s **Automatic Replies (Out-of-Office messages)** when the user is unavailable.

**Commands:**
```powershell
# Check Out-of-Office message status for a user
Get-MailboxAutoReplyConfiguration -Identity "UserMailbox@domain.com"

# Enable & Set Out-of-Office message
Set-MailboxAutoReplyConfiguration -Identity "UserMailbox@domain.com" -AutoReplyState Enabled -InternalMessage "I am currently out of the office and will respond when I return." -ExternalMessage "Currently unavailable. Please contact support@domain.com."
```

---

## **7. Email Forwarding Management**
**Purpose:**
- Ensures emails are correctly **forwarded or removed** from forwarding when employees change roles or leave.

**Commands:**
```powershell
# Enable email forwarding for a user
Set-Mailbox -Identity "UserMailbox@domain.com" -ForwardingSMTPAddress "ForwardTo@domain.com" -DeliverToMailboxAndForward $true

# Disable email forwarding
Set-Mailbox -Identity "UserMailbox@domain.com" -ForwardingSMTPAddress $null
```

---

## **8. Troubleshooting Email Delivery Issues**
**Purpose:**
- Helps IT to track **why an email was not delivered** and troubleshoot issues faster.

**Commands:**
```powershell
# Trace email delivery for a specific sender
Get-MessageTrace -Sender "user@domain.com" -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date)

# Get detailed message trace for a specific email
Get-MessageTraceDetail -MessageTraceId "<MessageTraceId>"
```

---

## **9. Removing Stale or Disabled Users from Groups**
**Purpose:**
- Keeps **Distribution Groups and Security Groups up-to-date** by removing disabled users.

**Commands:**
```powershell
# Remove disabled users from a Distribution List
$DL = "SalesTeam@domain.com"
$DisabledUsers = Get-ADUser -Filter {Enabled -eq $false} | Select -ExpandProperty UserPrincipalName

foreach ($user in $DisabledUsers) {
    Remove-DistributionGroupMember -Identity $DL -Member $user -Confirm:$false
}
```

---

## **10. Creating Shared Mailboxes**
**Purpose:**
- Allows IT to **set up new shared mailboxes** efficiently.

**Commands:**
```powershell
# Create a shared mailbox
New-Mailbox -Shared -Name "HR Team" -DisplayName "HR Team Mailbox" -PrimarySmtpAddress "HR@domain.com"

# Assign Full Access to the shared mailbox
Add-MailboxPermission -Identity "HR@domain.com" -User "HRUser@domain.com" -AccessRights FullAccess -AutoMapping $false
```

---

### **Summary**
These PowerShell commands are essential for daily Exchange Online administration, helps efficiently manage mailbox access, calendar permissions, distribution lists.


> [!NOTE]
> # Display name using mailbox as function

> [!IMPORTANT]
>> Define a function to prompt for DisplayName input and perform the mailbox search
````
Clear
function Get-MailboxByDisplayName {
    param (
        [string]$PromptMessage = "Please enter part of the DisplayName to search for:"
    )
    
    # Prompt the user for input
    $searchString = Read-Host -Prompt $PromptMessage

    # Check if the user entered something
    if (-not $searchString) {
        Write-Host "No input provided. Exiting..." -ForegroundColor Yellow
        return
    }

    # Ensure wildcards are included for partial matches
    $filterValue = "*$searchString*"

    # Perform the mailbox search using Get-Mailbox with a filter
    try {
        $mailboxes = Get-Mailbox -Filter "DisplayName -like '$filterValue'" | Select-Object Office, Alias, DisplayName, PrimarySmtpAddress, GrantSendOnBehalfTo
        
        # Check if any results were found
        if ($mailboxes) {
            Write-Host "Matching mailboxes found:" -ForegroundColor Green
            $mailboxes | Format-Table -AutoSize
        } else {
            Write-Host "No matching mailboxes found." -ForegroundColor Red
        }
    } catch {
        Write-Host "An error occurred while retrieving mailboxes: $_" -ForegroundColor Red
    }
}

# Call the function to execute it
Get-MailboxByDisplayName
````

Download the DisplayName.ps1 and save into Downloads folder.
````
iwr -Uri "https://raw.githubusercontent.com/ulyweb/exchange/refs/heads/main/scripts/Get-DisplayName.ps1" -OutFile "$env:HOME/Downloads/Get-DisplayName.ps1"
````
