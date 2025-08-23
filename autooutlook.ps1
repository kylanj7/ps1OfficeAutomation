$searchSubject = "The script will search for your outlook unread mail header here."
$recipientEmail = "user.name@webpage.com"
$autoSubject = "Your Subject Here"
$autoBody = "this is what's going to be insdie of your email."

try {
    $outlook = New-Object -comobject Outlook.Application
} catch {
    Write-Host "Could not connect to Outlook. Please ensure Outlook is installed and running."
    exit
}

$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
$unreadItems = $inbox.Items | Where-Object { $_.UnRead -eq $true }

foreach ($item in $unreadItems) {
    if ($item.Class -eq [Microsoft.Office.Interop.Outlook.OlObjectClass]::olMail) {
        if ($item.Subject -like "*$searchSubject*") {
            Write-Host "Found new email with subject: '$($item.Subject)'"
            $newEmail = $outlook.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olMailItem)
            $newEmail.To = $recipientEmail
            $newEmail.Subject = $autoSubject
            $newEmail.Body = $autoBody
            $newEmail.Send()
            Write-Host "Automated message sent to '$recipientEmail' successfully."
            $item.UnRead = $false
            Write-Host "Original email marked as read."
        }
    }
}

Write-Host "Script finished. No new approval emails found or all were processed."
