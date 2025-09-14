/**
 * Test Word Checker Commands
 * Event-based activation for OnMessageSend to scan for the word "test"
 */

// Global variable to store the event object for dialog communication
let currentSendEvent = null;

/**
 * Office.js initialization
 */
Office.onReady(() => {
    console.log("Test Word Checker commands initialized");
});

/**
 * Event handler for OnMessageSend
 * This function is called when the user attempts to send an email
 * @param {Office.AddinCommands.Event} event - The event object
 */
function onMessageSendHandler(event) {
    console.log("OnMessageSend event triggered");
    
    // Store the event for later use in dialog communication
    currentSendEvent = event;
    
    // Critical: Check if the user is offline
    // If offline, skip all add-in logic and allow the email to proceed to Outbox
    if (!navigator.onLine) {
        console.log("User is offline - skipping add-in logic and allowing send");
        event.completed({ allowEvent: true });
        return;
    }
    
    try {
        // Get the email body in plain text format
        Office.context.mailbox.item.body.getAsync(
            Office.CoercionType.Text,
            { asyncContext: event },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const emailBody = result.value;
                    console.log("Email body retrieved successfully");
                    
                    // Search for the word "test" (case-insensitive, whole word only)
                    const testWordRegex = /\btest\b/i;
                    const containsTestWord = testWordRegex.test(emailBody);
                    
                    if (containsTestWord) {
                        console.log("Found 'test' in email body - showing confirmation dialog");
                        
                        // Block the send and show confirmation dialog
                        showTestWordDialog(event);
                    } else {
                        console.log("No 'test' word found - allowing send");
                        
                        // Allow the email to be sent
                        event.completed({ allowEvent: true });
                    }
                } else {
                    console.error("Failed to get email body:", result.error);
                    
                    // If we can't read the body, allow the send to proceed
                    // (fail-safe approach to avoid blocking legitimate emails)
                    event.completed({ allowEvent: true });
                }
            }
        );
    } catch (error) {
        console.error("Error in onMessageSendHandler:", error);
        
        // If any unexpected error occurs, allow the send to proceed
        event.completed({ allowEvent: true });
    }
}

/**
 * Shows the confirmation dialog when "test" is found in the email
 * @param {Office.AddinCommands.Event} event - The send event
 */
function showTestWordDialog(event) {
    // Block the send and show dialog
    event.completed({
        allowEvent: false,
        options: {
            url: "https://your-github-username.github.io/your-repo-name/src/dialog/dialog.html",
            width: 35,
            height: 20
        }
    });
}

/**
 * Handler for messages received from the dialog
 * This function is called when the user makes a choice in the dialog
 * @param {object} arg - Message from the dialog
 */
function handleDialogMessage(arg) {
    console.log("Received message from dialog:", arg.message);
    
    if (!currentSendEvent) {
        console.error("No current send event available");
        return;
    }
    
    try {
        if (arg.message === "send") {
            console.log("User chose to send anyway");
            
            // Allow the email to be sent
            currentSendEvent.completed({ allowEvent: true });
        } else if (arg.message === "discard") {
            console.log("User chose to discard and edit");
            
            // Cancel the send (email remains in compose window for editing)
            currentSendEvent.completed({ allowEvent: false });
        } else {
            console.warn("Unknown dialog message:", arg.message);
            
            // Default to canceling the send for safety
            currentSendEvent.completed({ allowEvent: false });
        }
    } catch (error) {
        console.error("Error handling dialog message:", error);
        
        // If there's an error, cancel the send for safety
        currentSendEvent.completed({ allowEvent: false });
    } finally {
        // Clean up the event reference
        currentSendEvent = null;
    }
}

/**
 * Initialize dialog message handling
 * This sets up the communication channel with the dialog
 */
function initializeDialogHandling() {
    try {
        // Register the dialog message handler
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogMessageReceived,
            handleDialogMessage,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Dialog message handler registered successfully");
                } else {
                    console.error("Failed to register dialog message handler:", result.error);
                }
            }
        );
    } catch (error) {
        console.error("Error initializing dialog handling:", error);
    }
}

// Initialize dialog handling when Office.js is ready
Office.onReady(() => {
    initializeDialogHandling();
});

/**
 * Alternative approach for dialog communication (if the above doesn't work)
 * This function can be called directly from the dialog
 */
function receiveDialogMessage(message) {
    console.log("Direct dialog message received:", message);
    handleDialogMessage({ message: message });
}

// Make functions globally available for Office Add-ins
if (typeof window !== "undefined") {
    window.onMessageSendHandler = onMessageSendHandler;
    window.handleDialogMessage = handleDialogMessage;
    window.receiveDialogMessage = receiveDialogMessage;
}

// For Node.js environments (if needed)
if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        onMessageSendHandler,
        handleDialogMessage,
        receiveDialogMessage
    };
}