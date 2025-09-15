# Smart Alert Send Checker Outlook Add-in

This add-in intercepts the send event (OnMessageSend) in Outlook message compose and checks if the subject or body contains the word `test` (case-insensitive). If found, it prompts the user to either edit the message or send anyway.

## Files
- `manifest.xml` – Office Add-in manifest (host via raw URL or local for sideload)
- `function-file.html` – UI-less page hosting logic
- `function-file.js` – Implementation for OnMessageSend logic

## Hosting via GitHub Pages
1. Create (if not existing) a GitHub repository named `outlook-add-in` under the `smarthome-rus` account.
2. Place the three files (`manifest.xml`, `function-file.html`, `function-file.js`) and any icons under `assets/` if desired.
3. Enable GitHub Pages (Settings -> Pages) with branch `main` (or `master`) and root directory.
4. The files will be available at:
   - `https://smarthome-rus.github.io/outlook-add-in/manifest.xml`
   - `https://smarthome-rus.github.io/outlook-add-in/function-file.html`
   - `https://smarthome-rus.github.io/outlook-add-in/function-file.js`

Update URLs in `manifest.xml` if your path differs.

## Sideload in Outlook on the Web (OWA)
1. Go to Outlook on the web.
2. Open a new message compose window.
3. Select the gear icon -> "Manage add-ins" or "View all Outlook settings" -> Mail -> Customize actions -> Add-ins.
4. Choose "Add a custom add-in" -> "Add from file" and upload `manifest.xml` or use "Add from URL" and enter the GitHub Pages manifest URL.
5. Restart compose to ensure event handler loads.

## Behavior
- When you click Send:
  - The add-in retrieves the subject and plain text body.
  - If the word boundary match for `test` exists, a confirmation dialog (native `confirm`) appears.
  - Choosing OK = Send Anyway (allowEvent true)
  - Choosing Cancel = Edit Message (allowEvent false, returns to draft)

## Notes / Limitations
- A native Smart Alert with custom button labels would require richer UI surfaces not available in a pure UI-less function runtime; this implementation uses `confirm()` for two-button interaction while leveraging the event-based send interception.
- If APIs fail to load, the send is allowed to avoid blocking user flow.

## Logging
Open the browser developer tools (in OWA) or use Script Lab / Edge DevTools for Outlook (desktop preview) to see `console.log` output.

## Future Enhancements
- Replace `confirm()` with Smart Alerts extended button model if/when exposed for customization.
- Add localization resources.

## Validation
Ensure the requirement set Mailbox 1.13+ is available in your environment (new Outlook / OWA modern experience).
