# Outlook Smart Alert Add-in

A modern Outlook add-in that uses Smart Alerts to warn users when sending emails containing the word "test".

## Features

- **Smart Alerts Integration**: Uses Office.js Smart Alerts for non-intrusive notifications
- **OnMessageSend Event**: Triggers automatically when user attempts to send an email
- **Keyword Detection**: Scans email body for the word "test" (case-insensitive, whole word)
- **Cross-Platform**: Works on Outlook Classic (Windows), New Outlook, and Outlook on the web
- **Offline Support**: Gracefully handles offline scenarios by allowing emails to go to Outbox
- **Two-Phase Send Process**: 
  1. Block initial send and show notification if "test" is found
  2. Allow send after user confirmation via "Send Anyway" button

## File Structure

```
outlook_smart_alert/
├── manifest.xml                    # Add-in manifest with Smart Alerts configuration
├── src/
│   └── commands/
│       ├── commands.html           # HTML host page for Office.js
│       └── commands.js             # Core add-in logic
└── README.md                       # This file
```

## Installation

1. Host these files on a web server (configured for GitHub Pages at `https://smarthome-rus.github.io/outlook-add-in/`)
2. Sideload the `manifest.xml` file in Outlook
3. The add-in will automatically activate for OnMessageSend events

## Technical Requirements

- **Mailbox API Version**: 1.12 or higher (required for Smart Alerts)
- **Permissions**: ReadWriteItem
- **Supported Platforms**: 
  - Outlook Classic (Windows)
  - New Outlook (Windows)
  - Outlook on the Web

## How It Works

### Smart Alerts Flow

1. **Send Attempt**: User clicks Send button
2. **Event Trigger**: OnMessageSend event fires
3. **State Check**: Check if send was already approved from previous attempt
4. **Offline Check**: If offline, allow send to Outbox
5. **Content Scan**: Extract email body and search for "test" keyword using regex `/\btest\b/i`
6. **Decision**:
   - **No "test" found**: Allow send immediately
   - **"test" found**: Show Smart Alert notification and block send
7. **User Interaction**: Smart Alert displays with "Send Anyway" button
8. **Approval**: User clicks "Send Anyway", setting approval flag
9. **Second Send**: User clicks Send again, approval flag allows email to proceed

### Key Technical Components

- **State Management**: Uses `Office.context.mailbox.item.sessionData` to track user approval
- **Smart Alerts**: Uses `item.notificationMessages.addAsync()` with action buttons
- **Action Association**: `Office.actions.associate()` connects button clicks to handler functions
- **Async/Await**: Modern JavaScript patterns for clean, readable code

## Configuration

All URLs in the manifest point to GitHub Pages:
- Base URL: `https://smarthome-rus.github.io/outlook-add-in/`
- Function File: `/src/commands/commands.html`
- Commands Script: `/src/commands/commands.js`

## Development Notes

- The `<SupportsSharedFolders>true</SupportsSharedFolders>` tag is required in the manifest for Smart Alerts functionality
- Error handling ensures legitimate emails are never permanently blocked
- Console logging is included for debugging purposes
- Session data automatically clears when compose window is closed