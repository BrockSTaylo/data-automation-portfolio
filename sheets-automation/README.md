# Gmail → Google Sheets Logger

A Google Apps Script that reads Gmail messages by label and logs key fields into a Google Sheet automatically.

## What it does

- Reads all Gmail threads with a specified label (default: `to-log`)
- Extracts: date, sender, recipient, subject, body preview, and thread link
- Appends rows to a Google Sheet (default: `Email Log` tab)
- Removes the label after logging so threads aren't processed twice
- Runs hourly via a time-based trigger

## Setup

1. Open a Google Sheet
2. Go to **Extensions > Apps Script**
3. Paste the contents of `gmail_to_sheets.gs`
4. Update `LABEL_NAME` and `SHEET_NAME` at the top of the file
5. Run `setupTrigger()` once to start the hourly automation
6. Grant the required Gmail and Sheets permissions when prompted

## Customization

The script is designed to be easy to modify. Common changes:

- **Change the label:** Update `LABEL_NAME` to any Gmail label you use
- **Add more fields:** Extend the `rows.push([...])` array with additional message properties
- **Change frequency:** Modify `.everyHours(1)` in `setupTrigger()` to any interval
- **Filter by sender:** Add an `if` check on `firstMessage.getFrom()` inside the loop

## Use case

Ideal for small businesses that want to track incoming inquiries, support requests, or vendor emails in a shared spreadsheet without manual copy-paste.
