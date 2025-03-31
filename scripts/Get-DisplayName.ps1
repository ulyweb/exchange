# Define a function to prompt for DisplayName input and perform the mailbox search
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
