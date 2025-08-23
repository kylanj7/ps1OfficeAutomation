# Outlook Email Automation

PowerShell script that automatically responds to emails with specific subjects.

## Setup

1. Install Microsoft Outlook
2. Edit these variables in the script:
   - `$searchSubject` - Email subject to look for
   - `$recipientEmail` - Where to send the reply
   - `$autoSubject` - Reply subject
   - `$autoBody` - Reply message

## Usage

Run the script in PowerShell:
```powershell
.\script.ps1
```

The script will:
- Check unread emails in your inbox
- Find emails containing the search subject
- Send automated replies
- Mark original emails as read

## Requirements

- Outlook must be running
- PowerShell execution enabled
