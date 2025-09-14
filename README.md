# Test Word Checker - Outlook Add-in

A modern Outlook add-in that scans emails for the word "test" and asks for confirmation before sending.

## Features

- **Event-Based Activation**: Triggers automatically when users attempt to send emails
- **Cross-Platform Compatibility**: Works on Outlook Classic (Windows), New Outlook (Windows/Mac), and Outlook on the web
- **Offline Support**: Gracefully handles offline scenarios without blocking email sending
- **User-Friendly Dialog**: Clean, modern confirmation dialog with clear options
- **Robust Error Handling**: Fail-safe approach ensures legitimate emails are never blocked

## Files Structure

```
outlook_sonnect/
├── manifest.xml                 # Office Add-in manifest
├── assets/                      # Icon files for the add-in
│   ├── icon-16.svg             # 16x16 icon
│   ├── icon-32.svg             # 32x32 icon  
│   ├── icon-64.svg             # 64x64 icon
│   ├── icon-80.svg             # 80x80 icon
│   └── README.md               # Icon documentation
├── src/
│   ├── commands/
│   │   ├── commands.html        # Function file for event handlers
│   │   └── commands.js          # Core event handling logic
│   └── dialog/
│       ├── dialog.html          # Confirmation dialog UI
│       └── dialog.js            # Dialog interaction logic
├── support/                     # Support documentation
│   └── index.html              # Support page
└── README.md                    # This file
```

## How It Works

1. **Email Send Detection**: When a user clicks "Send" in Outlook, the `OnMessageSend` event is triggered
2. **Offline Check**: The add-in first checks if the user is online using `navigator.onLine`
3. **Body Scanning**: If online, it retrieves the email body in plain text and searches for the word "test" (case-insensitive, whole word)
4. **User Confirmation**: If "test" is found, a blocking dialog appears with two options:
   - **Send Anyway**: Proceeds with sending the email
   - **Discard & Edit**: Cancels the send and returns to the compose window
5. **Action Completion**: The user's choice is communicated back to complete or cancel the send action

## Deployment Instructions

### 1. GitHub Pages Setup

1. Create a new GitHub repository for your add-in
2. Upload all files to the repository
3. Enable GitHub Pages in repository settings
4. Note your GitHub Pages URL: `https://your-username.github.io/your-repo-name/`

### 2. Update Manifest URLs

Replace all placeholder URLs in `manifest.xml` with your actual GitHub Pages URLs:

```xml
<!-- Replace these placeholders: -->
https://your-github-username.github.io/your-repo-name/

<!-- With your actual URLs, for example: -->
https://johndoe.github.io/outlook-test-checker/
```

### 3. Install the Add-in

#### For Development/Testing:
1. Open Outlook on the web (outlook.office.com)
2. Go to Settings > Mail > General > Manage add-ins
3. Click "Add a custom add-in" > "Add from file"
4. Upload your `manifest.xml` file

#### For Organization Deployment:
1. Upload the manifest to your Office 365 admin center
2. Deploy to users through the add-in management portal

## Technical Implementation Details

### Event-Based Activation
- Uses `VersionOverridesV1_1` with `LaunchEvent` extension point
- Registers for `OnMessageSend` event type
- Requires `ReadWriteItem` permissions to access email body

### Offline Handling
```javascript
if (!navigator.onLine) {
    console.log("User is offline - skipping add-in logic");
    event.completed({ allowEvent: true });
    return;
}
```

### Word Detection
```javascript
const testWordRegex = /\btest\b/i;
const containsTestWord = testWordRegex.test(emailBody);
```

### Dialog Communication
- Uses `Office.context.ui.messageParent()` to send messages from dialog to parent
- Implements fallback communication methods for compatibility

## Browser Compatibility

- **Edge WebView2**: Outlook Classic on Windows
- **Modern Browsers**: New Outlook and Outlook on the web
- **Safari**: Outlook on Mac

## Security Considerations

- All resources served over HTTPS (GitHub Pages)
- Minimal permissions requested (ReadWriteItem only)
- Fail-safe error handling prevents email blocking
- No sensitive data stored or transmitted

## Customization

### Changing the Search Word
Modify the regex in `commands.js`:
```javascript
const testWordRegex = /\byour-word\b/i;
```

### Updating Dialog Message
Edit the message in `dialog.html`:
```html
<div class="message">
    Your custom message here
</div>
```

### Styling the Dialog
Modify the CSS in `dialog.html` to match your organization's branding.

## Troubleshooting

### Add-in Not Loading
1. Verify all URLs in manifest.xml are correct and accessible
2. Check browser console for JavaScript errors
3. Ensure GitHub Pages is serving files over HTTPS

### Event Not Triggering
1. Confirm the manifest has correct `OnMessageSend` configuration
2. Check that Office.js is loading properly
3. Verify the function name matches between manifest and code

### Dialog Not Appearing
1. Check that the dialog URL is accessible
2. Verify Office.js is initialized in the dialog
3. Ensure popup blockers are not interfering

## Development and Testing

For local development:
1. Serve files using a local HTTPS server
2. Update manifest URLs to point to localhost
3. Sideload the manifest in Outlook for testing

## License

This add-in is provided as-is for educational and development purposes.