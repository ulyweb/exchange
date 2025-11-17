Clear-Host
Write-Host "Manage Distribution Group Members" -ForegroundColor Green
Write-Host "This script first checks if the group exists and is of universal type. " -ForegroundColor Yellow
Write-Host "If so, it retrieves the group owner and lists the group members." -ForegroundColor Yellow
Write-Host "If the group does not exist or is not a universal distribution group, it provides an appropriate message" -ForegroundColor Yellow
Write-Host "`n`n"

# Prompt for the distribution group name
$groupName = Read-Host "Enter the Distribution Group Name"

# Get the distribution group
$group = Get-DistributionGroup -Identity $groupName

# Prepare an array to hold the output
$output = @()

# Check if the group exists and is of universal type
if ($group -ne $null -and $group.GroupType -eq "Universal") {
    Write-Host "Distribution Group Name: $($group.DisplayName)"
    
    # Get the owner of the group
    $groupOwner = $group.ManagedBy
    if ($groupOwner -ne $null) {
        Write-Host "Group Owner: $groupOwner"
    } else {
        Write-Host "No group owner found."
    }

    # List group members
    $groupMembers = Get-DistributionGroupMember -Identity $groupName -ResultSize Unlimited
    if ($groupMembers.Count -gt 0) {
        Write-Host "Group Members:"
        $groupMembers | ForEach-Object {
            if ($_.RecipientType -eq "UserMailbox") {
                $user = Get-User -Identity $_.SamAccountName
                if ($user -ne $null) {
                    Write-Host " - Name: $($_.Name)"
                    Write-Host "   - Title: $($user.Title)"
                    Write-Host "   - Department: $($user.Department)"
                    Write-Host "   - Office: $($user.Office)"
                    Write-Host "   - StreetAddress: $($user.StreetAddress)"
                    Write-Host "   - HideDLMembership: $($_.HideDLMembership)"
                    Write-Host "   - PrimarySmtpAddress: $($_.PrimarySmtpAddress)"

                    # Add the member to the output array
                    $output += New-Object PSObject -Property @{
                        Name = $_.Name
                        Title = $user.Title
                        Department = $user.Department
                        Office = $user.Office
                        StreetAddress = $user.StreetAddress
                        HideDLMembership = $_.HideDLMembership
                        PrimarySmtpAddress = $_.PrimarySmtpAddress
                        GroupName = $group.DisplayName
                        GroupOwner = $groupOwner
                        Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                } else {
                    Write-Host " - $($_.Name): This member could no longer be with the company."
                }
            } else {
                Write-Host " - Name: $($_.Name)"
                Write-Host "   - RecipientType: $($_.RecipientType)"
                
                # Add the member to the output array
                $output += New-Object PSObject -Property @{
                    Name = $_.Name
                    RecipientType = $_.RecipientType
                    GroupName = $group.DisplayName
                    GroupOwner = $groupOwner
                    Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }
        
        # Print total number of members
        Write-Host "`nTotal Members: $($groupMembers.Count)"
    } else {
        Write-Host "No group members found."
    }
} else {
    Write-Host "The group '$groupName' does not exist or is not a Universal Distribution Group."
}

# Export the output to a .csv file
$groupName = $groupName -replace '[\\/:*?"<>|]', '_'  # Replace invalid characters
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$filename = "c:\ps\output\${groupName}_Members_${timestamp}.csv"
$output | Export-Csv -Path $filename -NoTypeInformation

Write-Host "Exported group members to $filename"
