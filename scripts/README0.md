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
# Grant Full Access (without auto-mapping)
Add-MailboxPermission -Identity "UserMailbox@domain.com" -User "SupportUser@domain.com" -AccessRights FullAccess -AutoMapping $false

# Remove Full Access
Remove-MailboxPermission -Identity "UserMailbox@domain.com" -User "SupportUser@domain.com" -AccessRights FullAccess
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

### **Summary**
These PowerShell commands are essential for daily Exchange Online administration, helping IT efficiently manage mailbox access, calendar permissions, distribution lists, and email sending permissions.
