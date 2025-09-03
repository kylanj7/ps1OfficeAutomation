# --- Script Configuration (EDIT THESE VALUES) ---

# The email subject to search for.
$searchSubject = "Insert"

# The recipient's email address for the automated message.
$recipientEmail = "user@mail.com"

# The subject of the automated message.
$autoSubject = "Your Subject"

# Read the body from text file
try {
    $autoBody = Get-Content -Path "message_body.txt" -Raw -ErrorAction Stop
    Write-Host "Message body loaded from file"
} catch {
    Write-Host "Warning: Could not read message_body.txt. Using default message."
    $autoBody = "Default message body"
}

# --- Script Logic (DO NOT EDIT BELOW THIS LINE) ---

$checkInterval = 10
$outlook = $null

try {
    Write-Host "Starting email monitor. Press Ctrl+C to stop."
    
    while ($true) {
        try {
            # Connect to Outlook if needed
            if (-not $outlook) {
                $outlook = New-Object -ComObject Outlook.Application
                Write-Host "Connected to Outlook"
            }

            # Get inbox and unread items
            $namespace = $outlook.GetNamespace("MAPI")
            $inbox = $namespace.GetDefaultFolder(6) # olFolderInbox = 6
            $unreadItems = $inbox.Items | Where-Object { $_.UnRead -eq $true }

            # Process unread emails
            $foundMatch = $false
            foreach ($item in $unreadItems) {
                if ($item.Class -eq 43 -and $item.Subject -like "*$searchSubject*") { # olMail = 43
                    Write-Host "Found: $($item.Subject)"
                    
                    # Send automated response
                    $newEmail = $outlook.CreateItem(0) # olMailItem = 0
                    $newEmail.To = $recipientEmail
                    $newEmail.Subject = $autoSubject
                    $newEmail.Body = $autoBody
                    $newEmail.Send()
                    
                    # Mark as read
                    $item.UnRead = $false
                    
                    Write-Host "Response sent and email marked as read"
                    $foundMatch = $true
                    break  # Exit after first match
                }
            }
            
            if (-not $foundMatch) {
                Write-Host "$(Get-Date -Format 'HH:mm:ss') - No new matches found"
            }
            
            Start-Sleep -Seconds $checkInterval
            
        } catch {
            Write-Host "Error: $($_.Exception.Message)"
            Start-Sleep -Seconds $checkInterval
        }
    }
} finally {
    # Clean up Outlook COM object
    if ($outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        Stop-Process -Name Outlook -Force
        Write-Host "Outlook connection closed"
    }
    [System.GC]::Collect()
}
