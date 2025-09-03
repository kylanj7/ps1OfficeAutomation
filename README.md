# Outlook Email Automation

PowerShell script that automatically responds to emails with specific subjects.

## Setup

1. Install Microsoft Outlook
2. Edit these variables in the script:
   - `$searchSubject` - Email subject to look for
   - `$recipientEmail` - Where to send the reply
   - `$autoSubject` - Reply subject
   - `$autoBody` - Type your customized message in the "message_body.txt". This will be the message sent to your recipient.

## Usage

Run the script in PowerShell:
```powershell
.\autoemail.ps1
```

The script will:
- Check unread emails in your inbox
- Find emails containing the search subject
- Send automated replies
- Mark original emails as read

## Requirements

- Outlook must not be running
- PowerShell execution enabled
- ctrl+c to exit the program
